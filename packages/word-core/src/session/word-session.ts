/** 功能：提供独立于旧 DocModel、React 和文件对话框的 Word 会话；时间：2026-11-08；作者：lq。 */
import { Editor, Extension, type JSONContent } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { parseDocx, buildBlankDocx, readSections, type ParsedDocFull, type SaveOptions, type BlankDocxOptions } from '../upstream/packages/docx-engine/src/index';
import { editorExtensions } from '../upstream/apps/docs/src/renderer/editor/extensions';
import { blocksToPmDoc, pmDocOptions, type PmNode } from '../upstream/apps/docs/src/renderer/editor/convert';
import { setActiveSubEditor, getActiveSubEditor, subscribeSubEditorState } from '../upstream/apps/docs/src/renderer/editor/active-editor';
import { collectDocFonts } from '../upstream/apps/docs/src/renderer/font-check';
import { prefetchFaces } from '../upstream/apps/docs/src/renderer/font-resolve';
import { adoptEmbeddedFonts } from '../upstream/apps/docs/src/renderer/embedded-fonts';
import { docStyleCss } from '../upstream/apps/docs/src/renderer/doc-style-css';
import { setDocFontTable, loadBundledStandIns } from '../upstream/apps/docs/src/renderer/line-metrics';
import { applyDocLayoutSettings } from '../upstream/apps/docs/src/renderer/document-settings';
import { parseDocxOffThread } from '../upstream/apps/docs/src/renderer/parse-off-thread';
import { serializeWordDocument } from '../serialization/document';

/** 功能：工作区只消费会话摘要，不复制原始包、DOM 或第二份正文；时间：2026-11-08；作者：lq。 */
export interface WordSessionState { revision: number; savedRevision: number; dirty: boolean; canUndo: boolean; canRedo: boolean; }
/** 功能：记录待写盘的确切版本；只有外部写盘成功后才允许确认；时间：2026-11-08；作者：lq。 */
export interface PreparedWordSave { readonly revision: number; readonly bytes: Uint8Array; }
/** 功能：挂载和事件回调仅绑定我们自己的 UI；时间：2026-11-08；作者：lq。 */
export interface WordSessionOptions { element?: HTMLElement; readonly?: boolean; onChange?: (state: WordSessionState) => void; }

/** 功能：承接初始化、事务、活动子编辑器和保存快照生命周期，不自动写入文件；时间：2026-11-08；作者：lq。 */
export class WordCoreSession {
  readonly editor: Editor;
  readonly parsed: ParsedDocFull;
  private revision = 0;
  private savedRevision = 0;
  private options: SaveOptions = {};
  private preparing = false;
  private disposed = false;
  private readonly issued = new WeakSet<PreparedWordSave>();
  private readonly listeners = new Set<(state: WordSessionState) => void>();
  private readonly unsubscribeSubEditor: () => void;
  private readonly viewOnly: boolean;
  private readonly focusMain = () => setActiveSubEditor(null);

  /** 功能：打开完整 OOXML 模型，复制输入字节避免调用者修改来源缓存；时间：2026-11-08；作者：lq。 */
  static async open(bytes: Uint8Array, options: WordSessionOptions = {}): Promise<WordCoreSession> {
    const parsed = await parseDocxOffThread(bytes.slice(), { owned: true });
    // 功能：模型转编辑 schema 前先采用本文档字体，避免上一会话的字体提示进入新 DOM；时间：2026-11-08；作者：lq。
    setDocFontTable(parsed.fontTable);
    if (!await adoptEmbeddedFonts(parsed.embeddedFonts)) throw new Error('Word 打开被较新的字体激活替代');
    await prefetchFaces(collectDocFonts(parsed));
    await loadBundledStandIns(collectDocFonts(parsed));
    return new WordCoreSession(parsed, options);
  }
  /** 功能：空白文档也使用上游的完整包与编号初始化；时间：2026-11-08；作者：lq。 */
  static async create(options: WordSessionOptions = {}, documentOptions?: BlankDocxOptions): Promise<WordCoreSession> {
    return WordCoreSession.open(await buildBlankDocx(documentOptions), options);
  }
  private constructor(parsed: ParsedDocFull, options: WordSessionOptions) {
    this.parsed = parsed;
    this.viewOnly = options.readonly === true;
    if (options.onChange) this.listeners.add(options.onChange);
    // 功能：只读在事务层生效，防止外部命令绕过 UI 的禁用状态；时间：2026-11-08；作者：lq。
    const readonly = this.viewOnly;
    const readonlyGuard = Extension.create({ name: 'miniOfficeReadonlyGuard', addProseMirrorPlugins: () => [new Plugin({ filterTransaction: transaction => !readonly || !transaction.docChanged })] });
    this.editor = new Editor({
      element: options.element ?? document.createElement('div'),
      extensions: [...editorExtensions, readonlyGuard],
      content: blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed)) as JSONContent,
      editable: !this.viewOnly,
      // 功能：首个 NodeView 绘制前采用原始编号与样式定义，不能等第一次编辑才修正列表；时间：2026-11-08；作者：lq。
      onBeforeCreate: ({ editor }) => {
        editor.storage.listNumbering.styles = parsed.styles;
        editor.storage.listNumbering.docDefaults = parsed.docDefaults;
        editor.storage.listNumbering.defs = parsed.numbering;
      },
      editorProps: { attributes: { class: 'doc-page', spellcheck: 'false' } },
      onTransaction: ({ transaction }) => { if (transaction.docChanged) this.revision++; this.emit(); },
    });
    applyDocLayoutSettings(this.editor, parsed);
    this.editor.view.dom.addEventListener('focusin', this.focusMain);
    this.unsubscribeSubEditor = subscribeSubEditorState(changed => {
      const active = getActiveSubEditor();
      if (!active || !this.editor.view.dom.contains(active.view.dom)) return;
      if (changed) this.revision++;
      this.emit();
    });
  }
  /** 功能：激活文档时使用其字体表和嵌入字体，等待真实字体度量，再生成当前文档 CSS；时间：2026-11-08；作者：lq。 */
  async activate(): Promise<string> {
    this.assertOpen();
    setDocFontTable(this.parsed.fontTable);
    if (!await adoptEmbeddedFonts(this.parsed.embeddedFonts)) throw new Error('Word 字体激活被更新的文档替代');
    await prefetchFaces(collectDocFonts(this.parsed));
    await loadBundledStandIns(collectDocFonts(this.parsed));
    applyDocLayoutSettings(this.editor, this.parsed);
    this.assertOpen();
    return docStyleCss(this.parsed);
  }
  /** 功能：状态随当前选区的可撤销/可重做能力返回；时间：2026-11-08；作者：lq。 */
  state(): WordSessionState {
    return { revision: this.revision, savedRevision: this.savedRevision, dirty: this.revision !== this.savedRevision, canUndo: this.editor.can().undo(), canRedo: this.editor.can().redo() };
  }
  /** 功能：供现有工具栏直接订阅状态，不增加通用命令转发框架；时间：2026-11-08；作者：lq。 */
  subscribe(listener: (state: WordSessionState) => void): () => void { this.listeners.add(listener); return () => this.listeners.delete(listener); }
  /** 功能：页面、样式和页眉等非正文属性也进入同一保存版本；时间：2026-11-08；作者：lq。 */
  updateSaveOptions(patch: Partial<SaveOptions>): void {
    this.assertEditable();
    this.options = { ...this.options, ...structuredClone(patch) };
    this.revision++;
    this.emit();
  }
  /** 功能：保存/快照只序列化当前不可变事务快照，不清除 dirty 或提前重置来源锚点；时间：2026-11-08；作者：lq。 */
  async prepareSave(): Promise<PreparedWordSave> {
    this.assertEditable();
    if (this.preparing) throw new Error('Word 会话正在生成保存快照，请串行保存');
    if (this.editor.view.composing) throw new Error('中文输入尚未确认，请在组合输入结束后保存');
    this.preparing = true;
    try {
      const revision = this.revision;
      const document = this.editor.getJSON() as PmNode;
      const options = structuredClone(this.options);
      const bytes = await serializeWordDocument(this.parsed, document, options);
      this.assertOpen();
      const prepared = Object.freeze({ revision, bytes });
      this.issued.add(prepared);
      return prepared;
    } finally { this.preparing = false; }
  }
  /** 功能：外部确实写盘后确认版本；保存期间新输入仍然保持 dirty；时间：2026-11-08；作者：lq。 */
  acknowledgeSave(prepared: PreparedWordSave): void {
    this.assertOpen();
    if (!this.issued.has(prepared)) throw new Error('保存结果不属于此 Word 会话或已经确认');
    if (prepared.revision < this.savedRevision) throw new Error('过期的保存结果不能覆盖较新的保存状态');
    this.issued.delete(prepared);
    this.savedRevision = prepared.revision;
    this.emit();
  }
  /** 功能：关闭时释放核心运行时及选区监听，不影响其他会话；时间：2026-11-08；作者：lq。 */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.unsubscribeSubEditor();
    this.editor.view.dom.removeEventListener('focusin', this.focusMain);
    const active = getActiveSubEditor();
    if (active && this.editor.view.dom.contains(active.view.dom)) setActiveSubEditor(null);
    this.editor.destroy();
    this.listeners.clear();
  }
  /** 功能：拒绝访问已释放的运行时；时间：2026-11-08；作者：lq。 */
  private assertOpen(): void { if (this.disposed) throw new Error('Word 会话已关闭'); }
  /** 功能：非正文设置同样尊重只读模式；时间：2026-11-08；作者：lq。 */
  private assertEditable(): void { this.assertOpen(); if (this.viewOnly) throw new Error('Word 文档为只读'); }
  /** 功能：广播轻量状态，主体内容始终只有新核心中的一份；时间：2026-11-08；作者：lq。 */
  private emit(): void { if (!this.disposed) for (const listener of this.listeners) listener(this.state()); }
}
