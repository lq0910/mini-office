/** 功能：按来源字符区间替换简单/复杂/嵌套合并域，未知字段和非目标 XML 原样保留；时间：2026-11-11；作者：lq。 */
import { controlEffect, mergeFieldName, mergeFieldText, splitInstruction, type MergeRecord } from '../upstream/packages/docx-engine/src/mail-merge';
import { decodeEntities } from '../upstream/packages/docx-engine/src/parse-xml-text';
import { applyXmlEdits, wordXmlAttr, wordXmlEscape, wordXmlNode, xmlSpans, type XmlEdit, type XmlSpan } from './xml-spans';

interface Field { simple?: XmlSpan; begin?: XmlSpan; end?: XmlSpan; nodes: XmlSpan[]; texts: XmlSpan[]; results: XmlSpan[]; segments: Array<string | Field>; children: Field[]; result: boolean; }
export interface WordMergeCursor { index: number; skipped: boolean; }
/** 功能：脚注/批注引用按其所在记录块分配，NEXT 后不会继续采用首条数据；时间：2026-11-11；作者：lq。 */
export interface WordMergeReferences { id(kind: string, original: string, record: number): string; }
/** 功能：字段可跨 run/段落，复杂域用栈配对；损坏域不得部分合并；时间：2026-11-11；作者：lq。 */
function fieldsOf(xml: string): { fields: Field[]; nodes: XmlSpan[] } {
  const nodes = xmlSpans(xml), fields: Field[] = [], stack: Field[] = [];
  const attach = (field: Field) => { const parent = stack.at(-1); (parent?.children ?? fields).push(field); if (parent && !parent.result) parent.segments.push(field); stack.push(field); };
  const visit = (node: XmlSpan) => {
    if (wordXmlNode(node, 'fldSimple')) {
      const field: Field = { simple: node, nodes: [], texts: [], results: [], segments: [wordXmlAttr(node, 'instr') ?? ''], children: [], result: true };
      attach(field); node.children.forEach(visit); if (stack.pop() !== field) throw new Error('嵌套域不完整，已取消合并。'); return;
    }
    if (wordXmlNode(node, 'fldChar')) {
      const type = wordXmlAttr(node, 'fldCharType');
      if (type === 'begin') { attach({ begin: node, nodes: [node], texts: [], results: [], segments: [], children: [], result: false }); return; }
      const field = stack.at(-1);
      if (field && type === 'separate') { field.result = true; field.nodes.push(node); return; }
      if (field && type === 'end') { if (field.simple) throw new Error('复杂域边界交叉，已取消合并。'); field.end = node; field.nodes.push(node); stack.pop(); return; }
    }
    if (wordXmlNode(node, 'instrText') && stack.length) {
      const field = stack.at(-1)!; field.nodes.push(node);
      if (!field.result) field.segments.push(decodeEntities(xml.slice(node.openEnd, node.closeStart))); return;
    }
    if (wordXmlNode(node, 't')) for (const field of stack) { field.texts.push(node); if (field.result) field.results.push(node); }
    node.children.forEach(visit);
  };
  nodes.filter(node => !node.parent).forEach(visit);
  if (stack.length) throw new Error('复杂域缺少结束标记，已取消合并。');
  return { fields, nodes };
}
/** 功能：嵌套 MERGEFIELD 以字段引用进入 IF/SKIPIF 表达式，沿用核心条件规则；时间：2026-11-11；作者：lq。 */
function instructionOf(field: Field, record: MergeRecord): string {
  return field.segments.map(segment => {
    if (typeof segment === 'string') return segment;
    const instr = instructionOf(segment, record); const name = mergeFieldName(instr);
    if (name) return `«${name}»`;
    const effect = controlEffect(instr, record); return effect?.kind === 'if' ? `"${effect.text.replace(/"/g, '')}"` : '';
  }).join('');
}
/** 功能：预检整个原始部件的合并列及 NEXT 数，包含表格/嵌套表格/文本框中的真实域；时间：2026-11-11；作者：lq。 */
export function inspectWordMergeXml(xml: string): { fields: string[]; next: number } {
  const names = new Set<string>(); let next = 0;
  const visit = (field: Field) => { const instr = instructionOf(field, {}); const name = mergeFieldName(instr); if (name) names.add(name); if (splitInstruction(instr)[0] === 'NEXT') next++; field.children.forEach(visit); };
  fieldsOf(xml).fields.forEach(visit); return { fields: [...names], next };
}
/** 功能：多行和制表字段值生成真正的 Word 换行/制表节点，保留目标结果 run 的格式；时间：2026-11-11；作者：lq。 */
function textXml(value: string, prefix: string): string {
  const p = prefix ? `${prefix}:` : '';
  return `<${p}t xml:space="preserve">${wordXmlEscape(value).replace(/\r\n?|\n/g, `</${p}t><${p}br/><${p}t xml:space="preserve">`).replace(/\t/g, `</${p}t><${p}tab/><${p}t xml:space="preserve">`)}</${p}t>`;
}
/** 功能：一次模板遍历共享 NEXT 游标；超出记录末尾的槽位留空，未知域保持动态语义；时间：2026-11-11；作者：lq。 */
export function mergeWordXml(xml: string, records: readonly MergeRecord[], cursor: WordMergeCursor, allowNext = true, references?: WordMergeReferences): string {
  const { fields, nodes } = fieldsOf(xml); const edits: XmlEdit[] = []; const controls = new Set<XmlSpan>();
  const initialRecord = cursor.index, advances: Array<{ end: number; index: number }> = [];
  const remove = (node: XmlSpan) => edits.push({ start: node.start, end: node.end, text: '' });
  const visit = (field: Field) => {
    const record = records[cursor.index] ?? {}; const instr = instructionOf(field, record); const keyword = splitInstruction(instr)[0];
    const effect = controlEffect(instr, record);
    if (keyword !== 'MERGEFIELD' && !effect) { field.children.forEach(visit); return; }
    if (effect?.kind === 'next') { if (!allowNext) throw new Error('NEXT 只能用于主文档记录块，页眉页脚和注释内不能推进记录。'); cursor.index++; advances.push({ end: (field.simple ?? field.end!).end, index: cursor.index }); }
    if (effect?.kind === 'skip') cursor.skipped = true;
    const value = keyword === 'MERGEFIELD' ? mergeFieldText(instr, record) : effect?.kind === 'if' ? effect.text : '';
    const allNodes = (item: Field): XmlSpan[] => [...item.nodes, ...item.children.flatMap(allNodes)];
    const allFields = (item: Field): Field[] => [item, ...item.children.flatMap(allFields)];
    const removedNodes = new Set(allNodes(field));
    removedNodes.forEach(remove);
    for (const item of allFields(field)) if (item.simple) {
      edits.push({ start: item.simple.start, end: item.simple.openEnd, text: item === field && item.simple.end === item.simple.openEnd && value ? `<${item.simple.prefix}:r>${textXml(value, item.simple.prefix)}</${item.simple.prefix}:r>` : '' });
      if (item.simple.closeStart < item.simple.end) edits.push({ start: item.simple.closeStart, end: item.simple.end, text: '' });
    }
    const target = field.results[0];
    for (const node of new Set(field.texts)) edits.push({ start: node.start, end: node.end, text: node === target ? textXml(value, node.prefix) : '' });
    if (!target && value) {
      if (field.end) { const edit = edits.find(edit => edit.start === field.end!.start && edit.end === field.end!.end)!; edit.text = textXml(value, field.end.prefix); }
      else if (field.simple) { const edit = edits.find(edit => edit.start === field.simple!.closeStart && edit.end === field.simple!.end); if (edit) edit.text = `<${field.simple.prefix}:r>${textXml(value, field.simple.prefix)}</${field.simple.prefix}:r>`; }
    }
    // 功能：清除旧结果中的换行、图片和脚注等可见缓存，避免新值后仍残留旧图/旧行；时间：2026-11-11；作者：lq。
    const start = (field.simple ?? field.begin!).start, end = (field.simple ?? field.end!).end;
    const cached = nodes.filter(node => node.start >= start && node.end <= end && ['br', 'cr', 'tab', 'ptab', 'drawing', 'pict', 'object', 'sym', 'noBreakHyphen', 'softHyphen', 'footnoteReference', 'endnoteReference'].some(local => wordXmlNode(node, local)));
    for (const node of cached.filter(node => !cached.some(parent => parent !== node && parent.start <= node.start && parent.end >= node.end))) {
      for (let i = edits.length - 1; i >= 0; i--) if (edits[i].start >= node.start && edits[i].end <= node.end) edits.splice(i, 1);
      remove(node);
    }
    if (keyword === 'NEXT' || keyword === 'SKIPIF') {
      let paragraph = field.simple ?? field.begin; while (paragraph && !wordXmlNode(paragraph, 'p')) paragraph = paragraph.parent;
      if (paragraph) controls.add(paragraph);
    }
  };
  fields.forEach(visit);
  if (references) for (const node of nodes) {
    if (!/^(?:footnote|endnote|comment)(?:Reference|RangeStart|RangeEnd)$/.test(node.local) || !wordXmlNode(node, node.local)) continue;
    if (edits.some(edit => edit.start <= node.start && edit.end >= node.end && !edit.text)) continue;
    const original = wordXmlAttr(node, 'id'); if (original === undefined) continue;
    const index = advances.filter(advance => advance.end <= node.start).at(-1)?.index ?? initialRecord;
    const id = references.id(node.local.replace(/Reference|RangeStart|RangeEnd/g, ''), original, index);
    const tag = xml.slice(node.start, node.openEnd).replace(/((?:\w+:)?id\s*=\s*)(["'])(.*?)\2/, (_m, before, quote) => `${before}${quote}${wordXmlEscape(id)}${quote}`);
    edits.push({ start: node.start, end: node.openEnd, text: tag });
  }
  // 功能：仅控制域组成的独立段落不占标签高度；保留书签/分节及单元格最后一段；时间：2026-11-11；作者：lq。
  for (const paragraph of controls) {
    const contained = edits.filter(edit => edit.start >= paragraph.start && edit.end <= paragraph.end);
    const fragment = applyXmlEdits(xml.slice(paragraph.start, paragraph.end), contained.map(edit => ({ ...edit, start: edit.start - paragraph.start, end: edit.end - paragraph.start })));
    const visible = /<(?:\w+:)?(?:t|drawing|pict|object|fldChar|instrText|bookmarkStart|bookmarkEnd|sectPr|br|tab)(?=[\s/>])/.test(fragment);
    const lastCellParagraph = paragraph.parent && wordXmlNode(paragraph.parent, 'tc') && paragraph.parent.children.filter(node => wordXmlNode(node, 'p')).length === 1;
    if (!visible && !lastCellParagraph) {
      for (let i = edits.length - 1; i >= 0; i--) if (edits[i].start >= paragraph.start && edits[i].end <= paragraph.end) edits.splice(i, 1);
      remove(paragraph);
    }
  }
  // 功能：域指令之外的正文、属性及所有未知节点不参与替换；时间：2026-11-11；作者：lq。
  return applyXmlEdits(xml, edits);
}
