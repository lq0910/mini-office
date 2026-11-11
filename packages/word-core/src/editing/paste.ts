/** 功能：每个原生编辑区独立处理三种粘贴模式、选项切换和同一步历史；时间：2026-11-11；作者：lq。 */
import { Extension, type Editor } from '@tiptap/core';
import { DOMParser as PmDOMParser, Slice, type Node } from '@tiptap/pm/model';
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state';
import { closeHistory } from '@tiptap/pm/history';
import type { EditorView } from '@tiptap/pm/view';
import { capturePasteRestore, revertPaste, caretStyleMark, caretParaFormat, mergeFormattingFragment, type PasteMode, type PasteRestore } from '../upstream/apps/docs/src/renderer/editor/paste-options';
import { pasteTextSlice, singleCellPasteText } from '../upstream/apps/docs/src/renderer/editor/paste-text';
import { isForeignPasteHtml, pushDownWebInlineStyles, caretFontSlots, fillMissingRunFonts } from '../upstream/apps/docs/src/renderer/editor/paste-web-html';
import { readWordPictureFile } from './picture-files';
import { insertWordPictures } from './objects';

/** 功能：载荷只属于当前编辑器及精确快照，后续正文编辑使其失效；时间：2026-11-11；作者：lq。 */
interface Payload { html: string; text: string; mode: PasteMode; }
export interface WordPasteRecord extends Payload { doc: Node; restore: PasteRestore; time: number; }
/** 功能：历史边界独立于格式浮层，图片/内部/纯文本粘贴及关闭浮层后也必须单独撤销；时间：2026-11-11；作者：lq。 */
interface PasteState { record: WordPasteRecord | null; pasted: boolean; }
const key = new PluginKey<PasteState>('miniOfficePasteOptions');
const APPLY = 'miniOfficePasteApply';
/** 功能：沿用网页格式下推，清理 Office 包装和列表内多余段落；时间：2026-11-11；作者：lq。 */
export function cleanWordPasteHtml(html: string): string {
  const cleaned = html.replace(/<!--\[if[\s\S]*?<!\[endif\]-->/g, '').replace(/<o:p>[\s\S]*?<\/o:p>/g, '')
    .replace(/<li([^>]*)>\s*<p[^>]*>([\s\S]*?)<\/p>\s*<\/li>/g, '<li$1>$2</li>');
  return isForeignPasteHtml(html) ? pushDownWebInlineStyles(cleaned) : cleaned;
}
/** 功能：HTML 单独存在时以块边界生成纯文本；时间：2026-11-11；作者：lq。 */
function htmlText(html: string): string {
  const body = new window.DOMParser().parseFromString(html, 'text/html').body;
  for (const block of body.querySelectorAll('p,div,li,tr,h1,h2,h3,h4,h5,h6,br')) block.after(body.ownerDocument.createTextNode('\n'));
  return (body.textContent ?? '').replace(/\n{2,}/g, '\n').trim();
}
/** 功能：纯文本读取粘贴前 state 的字号和段落属性，切换时不读取已贴入的来源格式；时间：2026-11-11；作者：lq。 */
const textSlice = (state: EditorState, text: string) => pasteTextSlice(text, state.selection.$from, { state } as EditorView);
/** 功能：图片原子粘贴到正文末尾后保留可编辑尾段，避免 Chromium 将末尾对象选区回退成光标而使剪切失效；时间：2026-11-11；作者：lq。 */
function finishPaste(tr: Transaction): Transaction {
  if (tr.doc.lastChild?.type.name === 'docProtected' && tr.doc.lastChild.attrs.blockType === 'image' && tr.doc.type.schema.nodes.docParagraph) tr.insert(tr.doc.content.size, tr.doc.type.schema.nodes.docParagraph.create());
  return tr;
}
/** 功能：目标模式保留强调和结构，源模式补齐缺失的字体槽；时间：2026-11-11；作者：lq。 */
function formatSlice(slice: Slice, state: EditorState, mode: PasteMode): Slice {
  const { schema, selection, storedMarks } = state;
  const marks = storedMarks ?? selection.$from.marks();
  if (mode === 'merge') return new Slice(mergeFormattingFragment(slice.content, schema, caretStyleMark(marks, schema), caretParaFormat(selection.$from, schema)), slice.openStart, slice.openEnd);
  const slots = schema.marks.docTextStyle ? caretFontSlots(marks, schema.marks.docTextStyle) : null;
  return slots ? new Slice(fillMissingRunFonts(slice.content, slots, schema), slice.openStart, slice.openEnd) : slice;
}
/** 功能：空正文段保留来源标题；普通插入按 slice 合并，单列无格式内容走文本通道；时间：2026-11-11；作者：lq。 */
function insertPayload(tr: Transaction, state: EditorState, payload: Payload, parsed?: Slice): Transaction {
  const { html, text, mode } = payload;
  const cellText = html ? singleCellPasteText(html) : null;
  if (!html || mode === 'text' || cellText !== null) return tr.replaceSelection(textSlice(state, cellText ?? (text || htmlText(html))));
  const body = new window.DOMParser().parseFromString(cleanWordPasteHtml(html), 'text/html').body;
  const parser = PmDOMParser.fromSchema(state.schema);
  const { $from, empty } = state.selection;
  if (isForeignPasteHtml(html) && empty && $from.depth === 1 && $from.parent.isTextblock && !$from.parent.content.size) {
    const content = parser.parse(body).content;
    if (content.size) return tr.replaceWith($from.before(1), $from.after(1), formatSlice(new Slice(content, 0, 0), state, mode).content);
  }
  const slice = parsed ?? parser.parseSlice(body, { preserveWhitespace: true, context: $from });
  return tr.replaceSelection(isForeignPasteHtml(html) ? formatSlice(slice, state, mode) : slice);
}
/** 功能：闭包隔离各文档载荷，不依赖上游模块级暂存；时间：2026-11-11；作者：lq。 */
function pastePlugin(editor: Editor): Plugin<PasteState> {
  let liveView: EditorView | undefined;
  return new Plugin<PasteState>({
    key,
    view(view) { liveView = view; return { destroy() { liveView = undefined; } }; },
    state: { init: () => ({ record: null, pasted: false }), apply: (tr, previous, before) => {
      const record = previous.record;
      const paste = tr.getMeta(APPLY) || tr.getMeta('uiEvent') === 'paste' || tr.getMeta('appendedTransaction')?.getMeta(APPLY);
      const pasted = tr.docChanged ? !!paste : previous.pasted;
      if (tr.getMeta(key) !== undefined) return { record: tr.getMeta(key), pasted };
      if (!tr.docChanged) return previous;
      // 功能：追加的尾段/表格修复也属于粘贴；按真正应用顺序累计逆步骤，不能只撤销第一条事务；时间：2026-11-11；作者：lq。
      if (record && tr.getMeta('appendedTransaction')?.getMeta(APPLY)) {
        const extra = capturePasteRestore(before, [tr]);
        if (extra) return { record: { ...record, doc: tr.doc, restore: { ...record.restore, steps: [...extra.steps, ...record.restore.steps] } }, pasted };
      }
      return { record: null, pasted };
    } },
    filterTransaction(tr, state) {
      // 功能：下一次普通输入独立撤销，不能与前一次粘贴捆绑；时间：2026-11-11；作者：lq。
      if (tr.docChanged && state === liveView?.state && key.getState(state)?.pasted && !tr.getMeta(APPLY) && !tr.getMeta('appendedTransaction')) closeHistory(tr);
      return true;
    },
    props: {
      transformPastedHTML: cleanWordPasteHtml,
      clipboardTextParser: (text, context, _plain, view) => pasteTextSlice(text, context, view),
      handlePaste(view, event, slice) {
        if (!view.editable || view.composing) return true;
        const data = event.clipboardData;
        if (!data) return false;
        const html = data.getData('text/html'); const text = data.getData('text/plain');
        // 功能：系统截图和图片文件无需 HTML；一次解码全部图片，版本变化或离开编辑区时拒绝迟到插入；时间：2026-11-11；作者：lq。
        const files = Array.from(data.files ?? []).filter(file => file.type.startsWith('image/'));
        if (files.length && !text.trim() && (!html || !htmlText(html)) && !/data:image\//i.test(html) && editor.schema.nodes.docProtected) {
          const doc = view.state.doc; const selection = view.state.selection; const bookmark = selection.getBookmark();
          event.preventDefault();
          void Promise.all(files.map(readWordPictureFile)).then(pictures => {
            if (view.isDestroyed || editor.isDestroyed) return;
            if (view.state.doc !== doc || !view.state.selection.eq(selection) || view.composing || !view.editable || !view.hasFocus()) throw new Error('图片加载期间文档或焦点已变化，请重新粘贴。');
            view.dispatch(view.state.tr.setSelection(bookmark.resolve(doc)));
            if (!insertWordPictures(editor, pictures)) throw new Error('当前区域不能粘贴这些图片。');
            view.focus();
          }).catch(error => { if (!view.isDestroyed) view.dom.dispatchEvent(new CustomEvent('mo:word-paste-error', { bubbles: true, detail: error instanceof Error ? error.message : String(error) })); });
          return true;
        }
        if (!html && !text) return false;
        const payload: Payload = { html, text, mode: 'source' };
        const tr = closeHistory(finishPaste(insertPayload(view.state.tr, view.state, payload, slice))).setMeta(APPLY, true).setMeta('uiEvent', 'paste').scrollIntoView();
        if (html && isForeignPasteHtml(html) && htmlText(html)) {
          const restore = capturePasteRestore(view.state, [tr]);
          if (restore) tr.setMeta(key, { ...payload, doc: tr.doc, restore, time: tr.time });
        }
        view.dispatch(tr); return true;
      },
    },
  });
}
export const wordPaste = Extension.create({ name: 'miniOfficePaste', addProseMirrorPlugins() { return [pastePlugin(this.editor)]; } });
/** 功能：精简文本框编辑器按需安装相同规则，保留其 schema；时间：2026-11-11；作者：lq。 */
export function ensureWordPaste(editor: Editor): void {
  if (!editor.isDestroyed && !editor.state.plugins.some(plugin => plugin.spec.key === key)) editor.registerPlugin(pastePlugin(editor));
}
/** 功能：只公开当前可切换载荷，关闭面板不修改正文；时间：2026-11-11；作者：lq。 */
export const wordPasteRecord = (editor: Editor) => key.getState(editor.state)?.record ?? null;
export function dismissWordPaste(editor: Editor): void { if (!editor.isDestroyed) editor.view.dispatch(editor.state.tr.setMeta(key, null)); }
/** 功能：从原插入点重放模式，同一历史时间确保延迟点击后仍一次撤销整个粘贴；时间：2026-11-11；作者：lq。 */
export function applyWordPasteMode(editor: Editor, mode: PasteMode): boolean {
  const record = wordPasteRecord(editor);
  if (!record || record.doc !== editor.state.doc || editor.isDestroyed || !editor.isEditable || editor.view.composing) return false;
  if (record.mode === mode) return true;
  const tr = revertPaste(editor.state, record.restore);
  if (!tr) return false;
  // 功能：逆向快照校验同样经过插件过滤，必须提前标明重放，不能意外开启新历史组；时间：2026-11-11；作者：lq。
  tr.setMeta(APPLY, true);
  const before = editor.state.apply(tr);
  const pasteOnly = finishPaste(insertPayload(before.tr, before, { ...record, mode }));
  const restore = capturePasteRestore(before, [pasteOnly]);
  if (!restore) return false;
  pasteOnly.steps.forEach(step => tr.step(step));
  tr.setSelection(pasteOnly.selection.getBookmark().resolve(tr.doc)).setStoredMarks(pasteOnly.storedMarks).setTime(record.time).setMeta(APPLY, true).setMeta('uiEvent', 'paste');
  tr.setMeta(key, { ...record, mode, restore, doc: tr.doc });
  editor.view.dispatch(tr.scrollIntoView()); editor.view.focus(); return true;
}
export type { PasteMode };
