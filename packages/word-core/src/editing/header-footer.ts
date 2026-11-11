/** 功能：页眉页脚富文本是主文档属性的受控投影，内容/撤销/保存均由唯一会话负责；时间：2026-11-11；作者：lq。 */
import { Editor, Extension } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import type { WordCoreSession } from '../session/word-session';
import { hfDocJson, hfValueFromDoc, hfSubExtensions } from '../upstream/apps/docs/src/renderer/editor/hf-editor';
import { setActiveSubEditor, notifySubEditorState } from '../upstream/apps/docs/src/renderer/editor/active-editor';
import { resolveHf, withHfEdit, type HfKind } from '../upstream/apps/docs/src/renderer/hf-sections';
import type { HfView } from '../upstream/apps/docs/src/renderer/doc-state';
import type { PmNode } from '../upstream/apps/docs/src/renderer/editor/convert';
import { wordPaste } from './paste';
// 功能：区域样式目录由主文档投影，创建和套用在同一事务中回写；时间：2026-11-11；作者：lq。
import { wordStyleCatalog, wordStyleCatalogOf } from './styles';

let nextHistoryGroup = -1000;
/** 功能：区域标识按分节/类型/变体绑定，不随分页 DOM 的重绘丢失；时间：2026-11-11；作者：lq。 */
export interface WordHeaderFooterTarget { section: number; kind: HfKind; variant: HfView; }
/** 功能：实时回写后的序列化无需提交事件或依赖失焦，取消输入法及只读在事务层保护；时间：2026-11-11；作者：lq。 */
export class WordHeaderFooterEditor {
  readonly editor: Editor;
  private readonly unregister: () => void;
  private readonly unsubscribe: () => void;
  private syncing = false;
  private disposed = false;
  private lastValue: string;
  private lastStyles: string;
  private lastTime = 0;
  private group = --nextHistoryGroup;
  private pasteGroup?: { time: number; group: number };
  private base;
  private initial: PmNode;

  constructor(private readonly core: WordCoreSession, host: HTMLElement, readonly target: WordHeaderFooterTarget, private readonly onError: (error: unknown) => void = () => {}) {
    if (!core.editor.isEditable) throw new Error('只读文档不能编辑页眉页脚');
    this.base = resolveHf(core.headerFooterState(), target.section, target.kind, target.variant).value ?? { text: '' };
    this.initial = { ...hfDocJson(this.base), attrs: { miniOfficeStyles: wordStyleCatalogOf(core.editor.state.doc) } };
    this.lastStyles = JSON.stringify(wordStyleCatalogOf(core.editor.state.doc));
    this.lastValue = JSON.stringify(this.base);
    // 功能：区域快捷键和工具栏历史均转入主文档，禁止建立第二个撤销栈；时间：2026-11-11；作者：lq。
    const history = Extension.create({
      name: 'miniOfficeRegionHistory',
      addCommands: () => ({
        undo: () => ({ dispatch, tr }) => { if (!core.editor.can().undo()) return false; tr.setMeta('preventDispatch', true); if (dispatch) core.editor.commands.undo(); return true; },
        redo: () => ({ dispatch, tr }) => { if (!core.editor.can().redo()) return false; tr.setMeta('preventDispatch', true); if (dispatch) core.editor.commands.redo(); return true; },
      }),
      addKeyboardShortcuts: () => ({ 'Mod-z': () => this.editor.commands.undo(), 'Shift-Mod-z': () => this.editor.commands.redo(), 'Mod-y': () => this.editor.commands.redo() }),
    });
    this.editor = new Editor({
      element: host,
      extensions: [...hfSubExtensions.filter(extension => extension.name !== 'undoRedo'), history, wordPaste, wordStyleCatalog],
      content: this.initial,
      editorProps: { attributes: { class: 'page-hf-editor', spellcheck: 'false', 'aria-label': target.kind === 'header' ? '直接编辑页眉' : '直接编辑页脚' } },
      onFocus: () => setActiveSubEditor(this.editor),
      onTransaction: ({ transaction, editor }) => {
        if (this.syncing || this.disposed) return;
        if (transaction.docChanged) {
          const previousGroup = this.group;
          const now = Date.now();
          if (now - this.lastTime > 500 || (!transaction.getMeta('appendedTransaction') && transaction.steps.some(step => step.toJSON().stepType !== 'replace')) || transaction.getMeta('uiEvent') === 'paste') this.group = --nextHistoryGroup;
          // 功能：粘贴选项延迟重放沿用原时间的历史组，仍一次撤销整个载荷；时间：2026-11-11；作者：lq。
          if (transaction.getMeta('uiEvent') === 'paste') {
            if (this.pasteGroup?.time === transaction.time) this.group = this.pasteGroup.group;
            else this.pasteGroup = { time: transaction.time, group: this.group };
          }
          this.lastTime = now;
          // 功能：仅改名/导入目录时保留原部件字节，真正内容变化才重写区域；时间：2026-11-11；作者：lq。
          const contentChanged = !transaction.before.content.eq(transaction.doc.content);
          const value = contentChanged ? hfValueFromDoc(editor.getJSON() as PmNode, this.base, this.initial) : undefined;
          const next = value ? withHfEdit(core.headerFooterState(), target.section, target.kind, target.variant, value) : {};
          const styles = wordStyleCatalogOf(editor.state.doc);
          try {
            if (!next) throw new Error('页眉页脚所属分节已改变');
            this.syncing = true;
            core.commitPageLayout({ ...core.pageLayout(), ...next }, this.group, styles, previousGroup !== this.group);
            this.lastStyles = JSON.stringify(styles);
            this.lastValue = JSON.stringify(resolveHf(core.headerFooterState(), target.section, target.kind, target.variant).value);
          } catch (error) {
            this.onError(error);
            queueMicrotask(() => this.syncFromCore(true));
          } finally { this.syncing = false; }
        } else if (transaction.selectionSet) this.lastTime = 0;
        notifySubEditorState(false);
      },
    });
    this.initial = this.editor.getJSON() as PmNode;
    this.unregister = core.registerRegion(this.editor);
    this.unsubscribe = core.subscribe(() => this.syncFromCore());
  }
  /** 功能：主历史撤销或外部设置后回填区域，保留合法光标，不把回填重新记为编辑；时间：2026-11-11；作者：lq。 */
  private syncFromCore(force = false): void {
    if (this.syncing || this.disposed || this.editor.isDestroyed) return;
    const value = resolveHf(this.core.headerFooterState(), this.target.section, this.target.kind, this.target.variant).value ?? { text: '' };
    const signature = JSON.stringify(value);
    const styles = wordStyleCatalogOf(this.core.editor.state.doc);
    const styleSignature = JSON.stringify(styles);
    if (!force && signature === this.lastValue && styleSignature === this.lastStyles) return;
    this.syncing = true;
    try {
      const selection = this.editor.state.selection;
      this.base = value;
      if (force || signature !== this.lastValue) this.editor.commands.setContent({ ...hfDocJson(value), attrs: { miniOfficeStyles: styles } }, { emitUpdate: false });
      else this.editor.view.dispatch(this.editor.state.tr.setDocAttribute('miniOfficeStyles', styles));
      this.initial = this.editor.getJSON() as PmNode;
      const limit = this.editor.state.doc.content.size;
      const from = Math.min(selection.from, limit), to = Math.min(selection.to, limit);
      this.editor.view.dispatch(this.editor.state.tr.setSelection(TextSelection.near(this.editor.state.doc.resolve(Math.max(0, from)))));
      if (from !== to && this.editor.state.doc.resolve(to).parent.isTextblock) this.editor.commands.setTextSelection({ from: this.editor.state.selection.from, to });
      this.lastValue = signature; this.lastStyles = styleSignature; this.lastTime = 0;
    } finally { this.syncing = false; }
  }
  /** 功能：标签卸载/退出只清理投影，已提交内容继续留在会话中；时间：2026-11-11；作者：lq。 */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true; this.unsubscribe(); this.unregister();
    if (!this.editor.isDestroyed) this.editor.destroy();
  }
}
