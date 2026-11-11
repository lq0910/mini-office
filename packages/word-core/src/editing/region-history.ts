/** 功能：嵌入区域把目录和撤销路由给主文档，禁止另建样式或编辑历史；时间：2026-11-11；作者：lq。 */
import { Extension, type Editor } from '@tiptap/core';
import type { EditorView } from '@tiptap/pm/view';
import { undo, redo } from '@tiptap/pm/history';
import { getActiveSubEditor, setActiveSubEditor } from '../upstream/apps/docs/src/renderer/editor/active-editor';
let sequence = -100000;
type Region = { editor: import('@tiptap/core').Editor; identity: () => string };
const regions = new WeakMap<Editor, Set<Region>>();
const refresh = new WeakMap<Editor, { editor: Region['editor']; identity: string; from: number; to: number; editable: boolean; focused: boolean }>();
/** 功能：文档属性事务会令 ProseMirror 重建根视图；预先保存区域交互状态而不复制正文；时间：2026-11-11；作者：lq。 */
export function captureWordRegionRefresh(owner: Editor): void {
  const active = getActiveSubEditor();
  const region = [...regions.get(owner) ?? []].find(item => item.editor === active);
  if (!region || !active || active.isDestroyed) { refresh.delete(owner); return; }
  refresh.set(owner, { editor: active, identity: region.identity(), from: active.state.selection.from, to: active.state.selection.to, editable: active.isEditable, focused: active.isFocused });
}
/** 功能：根视图重建后恢复当前区域、选区和编辑模式，旧实例不能继续接受工具栏命令；时间：2026-11-11；作者：lq。 */
export function restoreWordRegionRefresh(owner: Editor): void {
  const saved = refresh.get(owner); refresh.delete(owner);
  if (!saved?.editor.isDestroyed) return;
  const editor = [...regions.get(owner) ?? []].find(item => item.identity() === saved.identity && !item.editor.isDestroyed)?.editor;
  if (!editor) return;
  editor.setEditable(saved.editable);
  editor.view.dom.closest('.doc-protected')?.classList.toggle('doc-content-editing', saved.editable);
  const limit = Math.max(1, editor.state.doc.content.size - 1);
  editor.commands.setTextSelection({ from: Math.min(saved.from, limit), to: Math.min(saved.to, limit) });
  setActiveSubEditor(editor);
  if (saved.focused) editor.view.focus();
}
/** 功能：在节点视图创建时立即登记，释放只清除对应实例；时间：2026-11-11；作者：lq。 */
export function registerWordRegion(owner: Editor, editor: Region['editor'], identity: () => string): () => void {
  const items = regions.get(owner) ?? new Set<Region>(); const item = { editor, identity };
  items.add(item); regions.set(owner, items);
  return () => items.delete(item);
}
/** 功能：主目录变化即时投影到所有文本框，投影事务不反向提交或增加历史；时间：2026-11-11；作者：lq。 */
export function syncWordRegionStyles(owner: Editor): void {
  const view = owner.view;
  const styles = view.state.doc.attrs.miniOfficeStyles;
  for (const { editor } of regions.get(owner) ?? []) {
    if (!editor.isDestroyed && JSON.stringify(editor.state.doc.attrs.miniOfficeStyles) !== JSON.stringify(styles)) {
      editor.view.dispatch(editor.state.tr.setDocAttribute('miniOfficeStyles', styles).setMeta('miniOfficeRegionProjection', true).setMeta('addToHistory', false));
    }
  }
}
/** 功能：每次独立区域操作分配唯一组合身份，连续输入由宿主复用；时间：2026-11-11；作者：lq。 */
export function nextWordRegionGroup(): number { return --sequence; }
/** 功能：文本框 schema 使用主目录属性并将快捷键/工具栏历史派发到主视图；时间：2026-11-11；作者：lq。 */
export function wordRegionHistory(view: EditorView) {
  return Extension.create({
    name: 'miniOfficeRegionHistory',
    addGlobalAttributes: () => [{ types: ['doc'], attributes: { miniOfficeStyles: { default: null, rendered: false } } }],
    addCommands: () => ({
      undo: () => ({ tr, dispatch }) => { if (!undo(view.state)) return false; tr.setMeta('preventDispatch', true); if (dispatch) undo(view.state, view.dispatch); return true; },
      redo: () => ({ tr, dispatch }) => { if (!redo(view.state)) return false; tr.setMeta('preventDispatch', true); if (dispatch) redo(view.state, view.dispatch); return true; },
    }),
    addKeyboardShortcuts() { return { 'Mod-z': () => this.editor.commands.undo(), 'Shift-Mod-z': () => this.editor.commands.redo(), 'Mod-y': () => this.editor.commands.redo() }; },
  });
}
