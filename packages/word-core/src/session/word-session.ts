/** 功能：提供独立于旧 DocModel、React 和文件对话框的 Word 会话；时间：2026-11-08；作者：lq。 */
import { Editor, Extension, type JSONContent } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
// 功能：目录页码仅作为与当前文档相配的排版缓存进入保存快照；时间：2026-11-11；作者：lq。
import { wordHeadings, wordTocSnapshot, type WordHeadingPages } from '../editing/toc';
import type { Node as WordNode } from '@tiptap/pm/model';
import { parseDocx, buildBlankDocx, readSections, type ParsedDocFull, type SaveOptions, type BlankDocxOptions } from '../upstream/packages/docx-engine/src/index';
import { editorExtensions, TABLE_TRAILING_SKIP } from '../upstream/apps/docs/src/renderer/editor/extensions';
// 功能：直接承接 GenOffice 的分批加载、尾部输入保护及修订忽略约定；时间：2026-11-13；作者：lq。
import { PHASE1_BLOCKS, PHASED_MIN_BLOCKS, setContentPhased, waitForFullContent, cancelPhasedContent, isPhasedContentPending, assertPhasedContentComplete } from '../upstream/apps/docs/src/renderer/phased-content';
import { PHASED_APPEND } from '../upstream/apps/docs/src/renderer/editor/streaming-tail-guard';
import { TRACK_IGNORE } from '../upstream/apps/docs/src/renderer/editor/revisions';
// 功能：保持同源分批正文和完整样式，尾部 DOM 使用一次片段提交；时间：2026-11-14；作者：lq。
import { batchWordTailDom } from '../layout/dom-batch';
import { blocksToPmDoc, pmDocOptions, type PmNode } from '../upstream/apps/docs/src/renderer/editor/convert';
import { setActiveSubEditor, getActiveSubEditor, subscribeSubEditorState } from '../upstream/apps/docs/src/renderer/editor/active-editor';
// 功能：文本框跟随主样式目录，不将回填记为新操作；时间：2026-11-11；作者：lq。
import { captureWordRegionRefresh, restoreWordRegionRefresh, syncWordRegionStyles } from '../editing/region-history';
import { collectDocFonts } from '../upstream/apps/docs/src/renderer/font-check';
import { prefetchFaces } from '../upstream/apps/docs/src/renderer/font-resolve';
import { adoptEmbeddedFonts } from '../upstream/apps/docs/src/renderer/embedded-fonts';
import { docStyleCss } from '../upstream/apps/docs/src/renderer/doc-style-css';
import { setDocFontTable, loadBundledStandIns } from '../upstream/apps/docs/src/renderer/line-metrics';
import { applyDocLayoutSettings } from '../upstream/apps/docs/src/renderer/document-settings';
import { parseDocxOffThread, disposeParseWorker } from '../upstream/apps/docs/src/renderer/parse-off-thread';
// 功能：旧 DOC 恢复数据在进入编辑 schema 前补充同源目录模型，开发和安装版本走同一路径；时间：2026-11-14；作者：lq。
import { adoptLegacyTocDisplay } from '../editing/legacy-toc';
import { serializeWordDocument } from '../serialization/document';
// 功能：会话管理列表资源，正文及历史仍由原生事务唯一持有；时间：2026-11-10；作者：lq。
import { WordListDefinitions } from '../editing/lists';
// 功能：页面目录进入正文原生历史，供布局与保存读取同一版本；时间：2026-11-10；作者：lq。
import { wordPageLayout, wordSections, wordHeaderFooterState, wordPageLayoutOf, commitWordPageLayout, type WordPageLayout } from '../editing/page-layout';
// 功能：样式定义与正文共用事务和保存快照，保持原始解析包不可变；时间：2026-11-10；作者：lq。
import { wordStyleCatalog, wordStyleCatalogOf, wordStyles, finishWordStyles, type StyleCatalog } from '../editing/styles';
// 功能：表格结构与属性回写保留未目标段落来源，不复制可编辑正文；时间：2026-11-11；作者：lq。
import { wordTableSource } from '../editing/table-source';
// 功能：正文与子编辑区共用三种粘贴规则，载荷按编辑器隔离；时间：2026-11-11；作者：lq。
import { wordPaste, ensureWordPaste } from '../editing/paste';
// 功能：属性面板与状态栏从唯一正文读取统计摘要；时间：2026-11-11；作者：lq。
import { wordDocumentStats } from '../editing/statistics';
// 功能：分节尺寸以原生装饰进入 DOM，不能触发段落重解析及选区丢失；时间：2026-11-11；作者：lq。
import { wordSectionWidths } from '../editing/section-width';
// 功能：跨页表格各列续行共享原生事务中的装饰，不复制可编辑正文；时间：2026-11-11；作者：lq。
import { wordCellGaps } from '../layout/cell-gaps';
// 功能：邮件面板配置只随保存快照输出，不制造正文 dirty 或第二套编辑模型；时间：2026-11-11；作者：lq。
import { patchWordMergeSettings, type WordMergeSettings } from '../serialization/mail-merge-settings';

/** 功能：工作区只消费会话摘要，不复制原始包、DOM 或第二份正文；时间：2026-11-08；作者：lq。 */
export interface WordSessionState { revision: number; savedRevision: number; dirty: boolean; canUndo: boolean; canRedo: boolean; loading: boolean; }
/** 功能：记录待写盘的确切版本；只有外部写盘成功后才允许确认；时间：2026-11-08；作者：lq。 */
export interface PreparedWordSave { readonly revision: number; readonly bytes: Uint8Array; }
/** 功能：挂载和事件回调仅绑定我们自己的 UI；时间：2026-11-08；作者：lq。 */
export interface WordSessionOptions { element?: HTMLElement; readonly?: boolean; onChange?: (state: WordSessionState) => void; }

// 功能：打开中的任务和已持有会话共同引用共享 worker，最后关闭才释放，不打断其他文档；时间：2026-11-11；作者：lq。
let runtimeUsers = 0;
function releaseRuntime(): void {
  if (--runtimeUsers === 0) { disposeParseWorker(); void adoptEmbeddedFonts(undefined); }
}

/** 功能：承接初始化、事务、活动子编辑器和保存快照生命周期，不自动写入文件；时间：2026-11-08；作者：lq。 */
export class WordCoreSession {
  readonly editor: Editor;
  readonly parsed: ParsedDocFull;
  private revision = 0;
  private savedRevision = 0;
  private options: SaveOptions = {};
  // 功能：输出预览和实际保存共用串行序列化队列，每次调用先捕获独立版本；时间：2026-11-11；作者：lq。
  private saveQueue: Promise<unknown> = Promise.resolve();
  private mergeSettings?: WordMergeSettings;
  private measuredPageCount?: number;
  private headingPages?: { doc: WordNode; pages: WordHeadingPages };
  // 功能：滚动、选区及分页装饰复用不可变正文的统计，编辑/撤销后按新正文重新计算；时间：2026-11-11；作者：lq。
  private statisticsCache?: { doc: WordNode; value: ReturnType<typeof wordDocumentStats> };
  private readonly initialDocument: WordNode;
  private disposed = false;
  private readonly issued = new WeakSet<PreparedWordSave>();
  private readonly listeners = new Set<(state: WordSessionState) => void>();
  private readonly unsubscribeSubEditor: () => void;
  private readonly viewOnly: boolean;
  private readonly lists: WordListDefinitions;
  /** 功能：页眉页脚是正文同级的受控区域，仅显式登记的子编辑器可接收本会话命令；时间：2026-11-11；作者：lq。 */
  private readonly regions = new Set<Editor>();
  /** 功能：子编辑器的 focusin 会冒泡，不能把其焦点误判为正文焦点；时间：2026-11-10；作者：lq。 */
  private readonly focusMain = (event: FocusEvent) => {
    const active = getActiveSubEditor();
    if (active?.view.dom.contains(event.target as Node)) return;
    setActiveSubEditor(null);
    this.emit();
  };

  /** 功能：打开完整 OOXML 模型，复制输入字节避免调用者修改来源缓存；时间：2026-11-08；作者：lq。 */
  static async open(bytes: Uint8Array, options: WordSessionOptions = {}): Promise<WordCoreSession> {
    runtimeUsers++;
    try {
      const parsed = await parseDocxOffThread(bytes.slice(), { owned: true });
      await adoptLegacyTocDisplay(parsed);
      // 功能：模型转编辑 schema 前先采用本文档字体，避免上一会话的字体提示进入新 DOM；时间：2026-11-08；作者：lq。
      setDocFontTable(parsed.fontTable);
      if (!await adoptEmbeddedFonts(parsed.embeddedFonts)) throw new Error('Word 打开被较新的字体激活替代');
      await prefetchFaces(collectDocFonts(parsed));
      await loadBundledStandIns(collectDocFonts(parsed));
      return new WordCoreSession(parsed, options);
    } catch (error) { releaseRuntime(); throw error; }
  }
  /** 功能：空白文档也使用上游的完整包与编号初始化；时间：2026-11-08；作者：lq。 */
  static async create(options: WordSessionOptions = {}, documentOptions?: BlankDocxOptions): Promise<WordCoreSession> {
    return WordCoreSession.open(await buildBlankDocx(documentOptions), options);
  }
  private constructor(parsed: ParsedDocFull, options: WordSessionOptions) {
    this.parsed = parsed;
    this.lists = new WordListDefinitions(parsed.numbering);
    // 功能：受保护文档在事务层保持只读，未经解锁不能由开发入口绕过保护；时间：2026-11-09；作者：lq。
    this.viewOnly = options.readonly === true || parsed.protection?.enforced === true;
    if (options.onChange) this.listeners.add(options.onChange);
    // 功能：只读在事务层生效，防止外部命令绕过 UI 的禁用状态；时间：2026-11-08；作者：lq。
    const readonly = this.viewOnly;
    // 功能：只读文档也允许系统装载来源正文；用户事务仍不能绕过只读；时间：2026-11-13；作者：lq。
    const readonlyGuard = Extension.create({ name: 'miniOfficeReadonlyGuard', addProseMirrorPlugins: () => [new Plugin({ filterTransaction: transaction => !readonly || !transaction.docChanged || transaction.getMeta(PHASED_APPEND) === true })] });
    const fullContent = blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed));
    const phased = (fullContent.content?.length ?? 0) > PHASED_MIN_BLOCKS;
    this.editor = new Editor({
      element: options.element ?? document.createElement('div'),
      extensions: [...editorExtensions, readonlyGuard, wordStyleCatalog, wordPageLayout, wordTableSource, wordPaste, wordSectionWidths(parsed), wordCellGaps],
      // 功能：首屏只创建上游规定的 64 个块，剩余正文由同一模型分帧追加；时间：2026-11-13；作者：lq。
      content: (phased ? { ...fullContent, content: fullContent.content!.slice(0, PHASE1_BLOCKS) } : fullContent) as JSONContent,
      editable: !this.viewOnly,
      // 功能：首个 NodeView 绘制前采用原始编号与样式定义，不能等第一次编辑才修正列表；时间：2026-11-08；作者：lq。
      onBeforeCreate: ({ editor }) => {
        editor.storage.listNumbering.styles = parsed.styles;
        editor.storage.listNumbering.docDefaults = parsed.docDefaults;
        editor.storage.listNumbering.defs = parsed.numbering;
      },
      editorProps: { attributes: { class: 'doc-page', spellcheck: 'false' } },
      onTransaction: ({ transaction, editor }) => {
        // 功能：来源追加不算编辑，不逐块刷新工具栏、样式投影或自动保存；时间：2026-11-13；作者：lq。
        if (transaction.getMeta(PHASED_APPEND)) return;
        // 功能：样式改名、撤销与重做即时更新渲染目录，不改原始样式定义；时间：2026-11-10；作者：lq。
        if (transaction.docChanged) { this.revision++; editor.storage.listNumbering.styles = wordStyles(editor.state.doc, parsed); syncWordRegionStyles(editor); }
        restoreWordRegionRefresh(editor);
        // 功能：分页/字体压缩的纯装饰事务不改变工具栏摘要，避免长文档首次排版触发整套 React 工具栏反复渲染；时间：2026-11-11；作者：lq。
        if (transaction.docChanged || transaction.selectionSet || transaction.storedMarksSet) this.emit();
      },
    });
    this.editor.on('beforeTransaction', ({ editor, transaction }) => { if (!transaction.getMeta(PHASED_APPEND)) captureWordRegionRefresh(editor); });
    this.initialDocument = phased ? this.editor.schema.nodeFromJSON(fullContent) : this.editor.state.doc;
    applyDocLayoutSettings(this.editor, parsed);
    this.editor.view.dom.addEventListener('focusin', this.focusMain);
    this.unsubscribeSubEditor = subscribeSubEditorState(changed => {
      const active = getActiveSubEditor();
      if (active && this.regions.has(active)) { this.emit(); return; }
      if (!active || !this.editor.view.dom.contains(active.view.dom)) return;
      if (changed) this.revision++;
      this.emit();
    });
    if (phased) {
      let first = true;
      const replace = (nodes: PmNode[], from: number, to: number) => {
        const tr = this.editor.state.tr.replaceWith(from, to, nodes.map(node => this.editor.schema.nodeFromJSON(node)))
          .setMeta('addToHistory', false).setMeta(PHASED_APPEND, true).setMeta(TRACK_IGNORE, true).setMeta(TABLE_TRAILING_SKIP, true);
        if (from === to) batchWordTailDom(this.editor.view.dom, nodes.length, () => this.editor.view.dispatch(tr));
        else this.editor.view.dispatch(tr);
        if (this.editor.state.doc !== tr.doc) throw new Error('Word 正文分批加载被拒绝');
      };
      setContentPhased({
        // 功能：构造时已装载首屏；上游失败回退仍可完整替换正文；时间：2026-11-13；作者：lq。
        setContent: doc => { if (first) { first = false; return; } replace(doc.content ?? [], 0, this.editor.state.doc.content.size); },
        appendNodes: nodes => replace(nodes, this.editor.state.doc.content.size, this.editor.state.doc.content.size),
        isDestroyed: () => this.disposed || this.editor.isDestroyed,
        isComposing: () => this.editor.view.composing || this.commandEditor().view.composing || [...this.regions].some(region => !region.isDestroyed && region.view.composing),
        // 功能：新会话历史本为空且所有追加排除历史，无需重建插件；用户首屏编辑历史保持原样；时间：2026-11-13；作者：lq。
        resetHistory: () => {},
        setLoading: () => this.emit(),
        getDirty: () => this.revision !== this.savedRevision,
        // 功能：追加事务从源头不增加修订号，不能覆盖加载期间的用户编辑版本；时间：2026-11-13；作者：lq。
        setDirty: () => {},
      }, fullContent, undefined, this.editor.view.dom);
    }
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
    return this.styleCss();
  }
  /** 功能：当前样式摘要/渲染使用事务目录，序列化仍以原始包为来源；时间：2026-11-10；作者：lq。 */
  formattingSource(): ParsedDocFull { return { ...this.parsed, styles: wordStyles(this.editor.state.doc, this.parsed) }; }
  /** 功能：样式变化时只重新生成文档 CSS，不重新打开正文或丢失历史；时间：2026-11-10；作者：lq。 */
  styleCss(): string { return docStyleCss(this.formattingSource()); }
  /** 功能：只读统计不创建编辑版本或历史；时间：2026-11-11；作者：lq。 */
  statistics() {
    const doc = this.editor.state.doc;
    if (this.statisticsCache?.doc !== doc) this.statisticsCache = { doc, value: wordDocumentStats(doc) };
    return this.statisticsCache.value;
  }
  /** 功能：状态随当前选区的可撤销/可重做能力返回；时间：2026-11-08；作者：lq。 */
  state(): WordSessionState {
    // 功能：历史按钮跟随当前活动编辑区，其他文档的全局子编辑器不能串入本会话；时间：2026-11-10；作者：lq。
    const editor = this.commandEditor();
    return { revision: this.revision, savedRevision: this.savedRevision, dirty: this.revision !== this.savedRevision, canUndo: editor.can().undo(), canRedo: editor.can().redo(), loading: this.contentPending() };
  }
  /** 功能：保存、查找和输出可等待本会话的完整来源，取消/关闭后拒绝半份正文；时间：2026-11-13；作者：lq。 */
  contentPending(): boolean { return !this.disposed && isPhasedContentPending(this.editor.view.dom); }
  async whenContentReady(): Promise<void> { this.assertOpen(); await waitForFullContent(this.editor.view.dom); this.assertOpen(); }
  /** 功能：工具栏只操作属于当前文档的活动子编辑器，其余情况使用正文；时间：2026-11-10；作者：lq。 */
  commandEditor(): Editor {
    const active = getActiveSubEditor();
    if (active && (this.regions.has(active) || this.editor.view.dom.contains(active.view.dom))) { ensureWordPaste(active); return active; }
    return this.editor;
  }
  /** 功能：供现有工具栏直接订阅状态，不增加通用命令转发框架；时间：2026-11-08；作者：lq。 */
  subscribe(listener: (state: WordSessionState) => void): () => void { this.listeners.add(listener); return () => this.listeners.delete(listener); }
  /** 功能：受控区域将内容即时提交到主历史，活动选区通知不能重复增加编辑版本；时间：2026-11-11；作者：lq。 */
  registerRegion(editor: Editor): () => void {
    this.assertEditable(); this.regions.add(editor);
    return () => { this.regions.delete(editor); if (getActiveSubEditor() === editor) setActiveSubEditor(null); };
  }
  /** 功能：列表命令只作用于本会话的活动编辑器，不产生独立资源历史；时间：2026-11-10；作者：lq。 */
  toggleList(kind: 'bullet' | 'ordered', editor = this.commandEditor()): boolean {
    if (this.viewOnly || (editor !== this.editor && editor !== this.commandEditor())) return false;
    return this.lists.toggle(editor, kind, this.editor);
  }
  /** 功能：样式的列表资源由会话分配，不跨文档复用旧编号；时间：2026-11-11；作者：lq。 */
  styleNumbering(kind: 'bullet' | 'ordered', numId?: string): string { this.assertEditable(); return this.lists.forStyle(this.editor, kind, numId); }
  /** 功能：页面和页眉设置只暴露原生摘要，提交进入同一撤销栈；时间：2026-11-10；作者：lq。 */
  sections() { return wordSections(this.editor.state.doc, this.parsed); }
  pageLayout() { return wordPageLayoutOf(this.editor.state.doc); }
  headerFooterState() { return wordHeaderFooterState(this.editor.state.doc, this.parsed); }
  evenAndOddHeaders(): boolean { return this.pageLayout().evenAndOddHeaders ?? !!this.parsed.evenAndOddHeaders; }
  commitPageLayout(next: WordPageLayout, historyGroup?: number, styles?: StyleCatalog, closeGroup = false): boolean { this.assertEditable(); return commitWordPageLayout(this.editor, this.parsed, next, historyGroup, styles, closeGroup); }
  /** 功能：分页记录精确文档版本，保存时拒绝使用编辑前的旧标题位置；时间：2026-11-11；作者：lq。 */
  recordHeadingPages(pages: WordHeadingPages): void { this.headingPages = { doc: this.editor.state.doc, pages: new Map(pages) }; }
  /** 功能：真实分页总数仅用作动态域缓存，不产生编辑版本或撤销项；时间：2026-11-10；作者：lq。 */
  recordPageCount(count: number): void { if (Number.isInteger(count) && count > 0) this.measuredPageCount = count; }
  /** 功能：页面、样式和页眉等非正文属性也进入同一保存版本；时间：2026-11-08；作者：lq。 */
  updateSaveOptions(patch: Partial<SaveOptions>): void {
    this.assertEditable();
    this.options = { ...this.options, ...structuredClone(patch) };
    this.revision++;
    this.emit();
  }
  /** 功能：保存用户实际调整的邮件面板配置，挂载/预览不调用此方法；时间：2026-11-11；作者：lq。 */
  updateMergeSettings(settings: WordMergeSettings): void { this.assertEditable(); this.mergeSettings = structuredClone(settings); }
  /** 功能：邮件合并等内存生成文档在首次实际保存前保持未保存状态，撤销正文也不能误认为已有磁盘副本；时间：2026-11-11；作者：lq。 */
  markUnsavedOrigin(): void { this.assertEditable(); if (this.savedRevision === 0) { this.savedRevision = -1; this.emit(); } }
  /** 功能：保存/快照只序列化当前不可变事务快照，不清除 dirty 或提前重置来源锚点；时间：2026-11-08；作者：lq。 */
  async prepareSave(): Promise<PreparedWordSave> {
    this.assertEditable();
    // 功能：分帧加载未结束时不得序列化截断正文；小文档保持同步捕获版本的原有语义；时间：2026-11-13；作者：lq。
    if (this.contentPending()) { await this.whenContentReady(); this.assertEditable(); }
    assertPhasedContentComplete(this.editor.view.dom);
    // 功能：受控页眉页脚及活动文本框的未确认输入也不能进入保存快照；时间：2026-11-11；作者：lq。
    if (this.editor.view.composing || [...this.regions].some(region => !region.isDestroyed && region.view.composing) || this.commandEditor().view.composing) throw new Error('中文输入尚未确认，请在组合输入结束后保存');
    const revision = this.revision;
    let document = this.editor.getJSON() as PmNode;
    // 功能：原文未改时保留来源缓存及原始字节；编辑后仅采用本版本的真实分页；时间：2026-11-11；作者：lq。
    if (!this.editor.state.doc.eq(this.initialDocument) && this.headingPages?.doc === this.editor.state.doc) {
      document = wordTocSnapshot(document, wordHeadings(this.editor.state.doc, this.formattingSource().styles), this.headingPages.pages);
    }
    const options = structuredClone(this.options);
    const mergeSettings = this.mergeSettings ? structuredClone(this.mergeSettings) : undefined;
    // 功能：已测量的实际总页数进入本次保存快照，未挂载时沿用来源缓存；时间：2026-11-10；作者：lq。
    if (this.measuredPageCount !== undefined) options.pageCount = this.measuredPageCount;
    // 功能：捕获定义和正文同一版本；撤销后尚未引用的新定义不会泄露到保存包；时间：2026-11-10；作者：lq。
    const styles = structuredClone(wordStyleCatalogOf(this.editor.state.doc));
    if (styles.definitions.length) options.styleUpserts = [...options.styleUpserts ?? [], ...styles.definitions.map(item => item.upsert)];
    // 功能：编号定义与本次正文快照配对，保存时继续编辑不会污染该版本；时间：2026-11-10；作者：lq。
    const numbering = this.lists.saveOptions(this.editor.state.doc);
    if (numbering.newDefs?.length) options.numbering = { ...options.numbering, newDefs: [...options.numbering?.newDefs ?? [], ...numbering.newDefs] };
    // 功能：排队前已捕获正文/定义/配置，后续输入不污染旧请求；失败不阻断下一次请求；时间：2026-11-11；作者：lq。
    const serialization = this.saveQueue.catch(() => {}).then(async () => {
      this.assertOpen();
      let bytes = await finishWordStyles(await serializeWordDocument(this.parsed, document, options), styles);
      if (mergeSettings) bytes = await patchWordMergeSettings(bytes, mergeSettings);
      return bytes;
    });
    this.saveQueue = serialization;
    const bytes = await serialization;
    this.assertOpen();
    const prepared = Object.freeze({ revision, bytes });
    this.issued.add(prepared);
    return prepared;
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
    // 功能：仅取消本标签的待追加帧，唤醒等待方后由已关闭检查拒绝保存；时间：2026-11-13；作者：lq。
    cancelPhasedContent(this.editor.view.dom);
    this.unsubscribeSubEditor();
    for (const region of this.regions) { if (getActiveSubEditor() === region) setActiveSubEditor(null); if (!region.isDestroyed) region.destroy(); }
    this.regions.clear();
    this.editor.view.dom.removeEventListener('focusin', this.focusMain);
    const active = getActiveSubEditor();
    if (active && this.editor.view.dom.contains(active.view.dom)) setActiveSubEditor(null);
    this.editor.destroy();
    this.listeners.clear();
    releaseRuntime();
  }
  /** 功能：拒绝访问已释放的运行时；时间：2026-11-08；作者：lq。 */
  private assertOpen(): void { if (this.disposed) throw new Error('Word 会话已关闭'); }
  /** 功能：非正文设置同样尊重只读模式；时间：2026-11-08；作者：lq。 */
  private assertEditable(): void { this.assertOpen(); if (this.viewOnly) throw new Error('Word 文档为只读'); }
  /** 功能：广播轻量状态，主体内容始终只有新核心中的一份；时间：2026-11-08；作者：lq。 */
  private emit(): void { if (!this.disposed) for (const listener of this.listeners) listener(this.state()); }
}
