/** 功能：原生 Word 字段插入和当前字段摘要，复用同源 MERGEFIELD 指令与唯一历史；时间：2026-11-11；作者：lq。 */
import type { Editor } from '@tiptap/core';
import type { Node } from '@tiptap/pm/model';
import { closeHistory } from '@tiptap/pm/history';
import { mergeFieldInstruction, mergeFieldName } from '../upstream/packages/docx-engine/src/mail-merge';
import type { ParsedDocFull, Run, TextboxDisplay } from '../upstream/packages/docx-engine/src/index';
import { inspectWordMergeXml } from '../serialization/merge-fields';
import { wordHeaderFooterState } from './page-layout';
import { resolveHf } from '../upstream/apps/docs/src/renderer/hf-sections';

/** 功能：字段继承当前字符格式，指令只标记占位文本，之后普通输入不会延续域标记；时间：2026-11-11；作者：lq。 */
export function insertWordMergeField(editor: Editor, name: string, next = false): boolean {
  if (editor.isDestroyed || !editor.isEditable || editor.view.composing || !editor.schema.marks.instrField || (!next && (!name.trim() || name.length > 200 || /["\r\n\u0000-\u001f]/.test(name)))) return false;
  const marks = (editor.state.storedMarks ?? editor.state.selection.$from.marks()).filter(mark => mark.type.name !== 'instrField');
  const field = editor.schema.marks.instrField.create({ instr: next ? 'NEXT' : mergeFieldInstruction(name) });
  const tr = closeHistory(editor.state.tr).replaceSelectionWith(editor.schema.text(next ? '\u200b' : `«${name}»`, [...marks, field]), false).setStoredMarks(marks).scrollIntoView();
  editor.view.dispatch(tr); editor.view.dispatch(closeHistory(editor.state.tr)); return true;
}
/** 功能：同一个连续域缓存可能拆成多种字符格式，统计时不重复计数 NEXT；时间：2026-11-11；作者：lq。 */
export function wordMergeFields(doc: Node, parsed?: ParsedDocFull): { fields: string[]; nextBlocks: number } {
  const fields = new Set<string>(); let next = 0, previousEnd = -1, previousInstruction = '';
  const sourceBlocks: string[] = [];
  // 功能：文本框保存即时更新 rich-run，字段摘要读取当前对象而非初始来源缓存；时间：2026-11-11；作者：lq。
  const runs = (items: readonly Run[]) => { for (const run of items) { const name = mergeFieldName(run.instrField ?? ''); if (name) fields.add(name); } };
  doc.descendants((node, pos) => {
    // 功能：受保护来源块和嵌套表格仍按其实际 XML 统计；删除这些块后不沿用初始包的陈旧字段；时间：2026-11-11；作者：lq。
    if (parsed && ['docProtected', 'docNestedTable'].includes(node.type.name)) {
      const source = typeof node.attrs.genXml === 'string' ? node.attrs.genXml : parsed.blocks.find(block => block.docxIndex === node.attrs.docxIndex)?.originalXml;
      const boxes = node.attrs.textboxes as TextboxDisplay[] | undefined;
      if (boxes) for (const box of boxes) for (const paragraph of box.paras) runs(paragraph.runs);
      if (source) sourceBlocks.push(boxes ? source.replace(/<w:txbxContent\b[^>]*>[\s\S]*?<\/w:txbxContent>/g, '') : source);
      return false;
    }
    if (!node.isText) return;
    const instruction = String(node.marks.find(mark => mark.type.name === 'instrField')?.attrs.instr ?? '');
    const name = mergeFieldName(instruction); if (name) fields.add(name);
    if (/^\s*NEXT\s*$/i.test(instruction) && !(previousEnd === pos && previousInstruction === instruction)) next++;
    previousEnd = pos + node.nodeSize; previousInstruction = instruction;
  });
  // 功能：TOC 等普通复杂域可跨多个来源段落，汇总完整序列后再配对，摘要失败不阻断编辑界面；时间：2026-11-11；作者：lq。
  if (sourceBlocks.length) {
    try { const info = inspectWordMergeXml(sourceBlocks.join('')); info.fields.forEach(name => fields.add(name)); next += info.next; }
    catch { /* 功能：损坏来源域保持原内容，真正合并时由完整包校验给出错误；时间：2026-11-11；作者：lq。 */ }
  }
  // 功能：页眉页脚字段跟随继承、变体及主历史中的实时编辑，去重不重复计算 NEXT 正文槽位；时间：2026-11-11；作者：lq。
  if (parsed) {
    const state = wordHeaderFooterState(doc, parsed);
    state.sections.forEach((_, section) => { for (const kind of ['header','footer'] as const) for (const variant of ['default','first','even'] as const) {
      for (const paragraph of resolveHf(state, section, kind, variant).value?.paras ?? []) { runs(paragraph.runs); for (const cell of paragraph.cells ?? []) for (const items of cell.paras) runs(items); }
    } });
  }
  return { fields: [...fields], nextBlocks: next ? next + 1 : 0 };
}
