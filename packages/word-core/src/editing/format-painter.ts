/** 功能：复制核心选区格式并以单个原生事务刷到目标，保留链接、语言及来源 XML；时间：2026-11-10；作者：lq。 */
import type { Editor } from '@tiptap/core';
import type { Mark } from '@tiptap/pm/model';
import { TextSelection, type Transaction } from '@tiptap/pm/state';
import { closeHistory } from '@tiptap/pm/history';
import type { ParsedDocFull } from '../upstream/packages/docx-engine/src/index';
import { readWordFormatting, selectedParagraphs, type WordParagraphFormat, type WordTextFormat } from './formatting';
import { wordRunEffectsXml } from '../serialization/run-effects';

/** 功能：仅持有格式值，不保留源正文、选区或编辑器引用；时间：2026-11-10；作者：lq。 */
export interface WordFormatSnapshot { text: WordTextFormat; paragraph: WordParagraphFormat;
  /** 功能：只复制已建模的字符效果和列表身份，未知 XML/语言/链接仍属于目标；时间：2026-11-11；作者：lq。 */
  effects?: Record<string, unknown>; list?: { kind: 'bullet' | 'ordered'; level: number; numId?: string };
}
const EFFECT_KEYS = ['charSpacingTwips', 'charScalePct', 'charScaleEm', 'charScaleX', 'kern', 'kernHalfPoints', 'textOutline', 'textEffect', 'dstrike', 'glow', 'positionHalfPoints', 'bdr', 'vertAlign', 'em', 'caps'] as const;

/** 功能：取源选区起点的实际格式，混合选区不把后续片段的格式误并入；时间：2026-11-10；作者：lq。 */
export function captureWordFormat(editor: Editor, parsed: ParsedDocFull): WordFormatSnapshot {
  const state = editor.state.selection.empty ? editor.state : editor.state.apply(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, editor.state.selection.from)));
  const format = readWordFormatting({ state } as Editor, parsed);
  const para = state.selection.$from.parent;
  const attrs = (state.storedMarks ?? state.selection.$from.marks()).find(mark => mark.type.name === 'docTextStyle')?.attrs ?? {};
  const inherited = parsed.styles.get(String(para.attrs.styleId ?? 'Normal'))?.display;
  const effects: Record<string, unknown> = Object.fromEntries(EFFECT_KEYS.map(key => [key, attrs[key] ?? (inherited as Record<string, unknown> | undefined)?.[key] ?? null]));
  for (const key of ['bdr', 'textOutline'] as const) if (effects[key] && typeof effects[key] === 'object') effects[key] = JSON.stringify(effects[key]);
  effects.miniOfficeEffectXml = wordRunEffectsXml({ ...attrs, ...effects });
  const p = format.paragraph.value;
  const borderSides = inherited?.borderSides;
  p.borders = para.attrs.borders ?? (borderSides ? Object.entries(borderSides).filter(([, line]) => line && !line.none).map(([side]) => side).join('') : '');
  p.borderLines = para.attrs.borderLines ?? (borderSides ? JSON.stringify(borderSides) : null);
  p.shadingFill = para.attrs.shadingFill ?? inherited?.shadingFill ?? null;
  p.tabStops = para.attrs.tabStops ?? (inherited?.tabStops ? JSON.stringify(inherited.tabStops) : null);
  return structuredClone({ text: format.text.value, paragraph: p, effects,
    ...(para.type.name === 'docListItem' ? { list: { kind: para.attrs.kind as 'bullet' | 'ordered', numId: String(para.attrs.numId), level: Number(para.attrs.ilvl) || 0 } } : para.attrs.listMarker?.kind ? { list: { kind: para.attrs.listMarker.kind, numId: para.attrs.listMarker.numId, level: para.attrs.listMarker.ilvl ?? 0 } } : {}) });
}

/** 功能：在调用者的同一事务中合并格式；整段样式与局部格式刷共用字符规则；时间：2026-11-10；作者：lq。 */
export function appendWordFormat(tr: Transaction, editor: Editor, format: WordFormatSnapshot, wholeParagraph = false): void {
  const paragraphs = selectedParagraphs(editor);
  const ranges = wholeParagraph ? paragraphs.map(({ node, pos }) => ({ from: pos + 1, to: pos + node.nodeSize - 1 }))
    : editor.state.selection.ranges.map(({ $from, $to }) => ({ from: $from.pos, to: $to.pos }));
  const t = format.text;
  const patch: Record<string, unknown> = { ...format.effects };
  if ('font' in t || 'fontAscii' in t) Object.assign(patch, { font: t.font ?? null, eastAsiaFont: t.font ?? null, fontAscii: t.fontAscii ?? t.font ?? null, eaSlotEmpty: null, themeRFonts: null });
  if ('sizeHalfPoints' in t) patch.sizeHalfPoints = t.sizeHalfPoints ?? null;
  if ('color' in t) Object.assign(patch, { color: t.color?.replace(/^#/, '') ?? 'auto', themeColor: null });
  if ('highlight' in t) Object.assign(patch, { highlight: null, shading: t.highlight?.replace(/^#/, '') ?? null, shadingDisplay: null });
  if ('bold' in t) patch.boldOff = !t.bold || null;
  if ('italic' in t) patch.italicOff = !t.italic || null;
  if ('underline' in t) patch.underlineOff = !t.underline || null;
  if ('strike' in t) patch.strikeOff = !t.strike || null;
  const styleType = editor.schema.marks.docTextStyle;
  /** 功能：只替换格式标记，超链接、批注、域及未知 rPr 按目标片段保留；时间：2026-11-10；作者：lq。 */
  const marksFor = (marks: readonly Mark[]) => {
    let result = [...marks];
    if (styleType) result = [...result.filter(mark => mark.type !== styleType), styleType.create({ ...marks.find(mark => mark.type === styleType)?.attrs, ...patch })];
    for (const name of ['bold', 'italic', 'underline', 'strike'] as const) {
      const type = editor.schema.marks[name];
      if (!type || !(name in t)) continue;
      result = result.filter(mark => mark.type !== type);
      if (t[name]) result.push(type.create());
    }
    return result;
  };
  for (const { from, to } of ranges) editor.state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) return;
    const start = Math.max(from, pos), end = Math.min(to, pos + node.nodeSize);
    const next = marksFor(node.marks);
    for (const mark of node.marks) if (!next.some(item => item.eq(mark))) tr.removeMark(start, end, mark.type);
    for (const mark of next) tr.addMark(start, end, mark);
  });
  if (editor.state.selection.empty) tr.setStoredMarks(marksFor(editor.state.storedMarks ?? editor.state.selection.$from.marks()));
  for (const { node, pos } of paragraphs) tr.setNodeMarkup(pos, undefined, { ...tr.doc.nodeAt(pos)!.attrs, ...format.paragraph,
    // 功能：长度型缩进/间距覆盖旧字符单位和自动间距，避免显示与保存冲突；时间：2026-11-10；作者：lq。
    ...('indentFirstLine' in format.paragraph ? { indentFirstLineChars: null } : {}),
    ...('spaceBefore' in format.paragraph ? { spaceBeforeAuto: false } : {}),
    ...('spaceAfter' in format.paragraph ? { spaceAfterAuto: false } : {}),
  });
  // 功能：区域列表使用原生 marker 属性，不能把精简 schema 替换成正文列表节点；时间：2026-11-11；作者：lq。
  if (format.list?.numId && !editor.schema.nodes.docListItem && editor.schema.nodes.docParagraph?.spec.attrs?.listMarker) for (const { pos } of paragraphs) {
    const current = tr.doc.nodeAt(pos)!;
    tr.setNodeMarkup(pos, undefined, { ...current.attrs, listMarker: { kind: format.list.kind, numId: format.list.numId, ilvl: format.list.level, text: format.list.kind === 'bullet' ? '•' : '1.', indentLeft: 720 * (format.list.level + 1), hanging: 360 } });
  }
  if (format.list?.numId && editor.schema.nodes.docListItem) for (const { pos } of paragraphs) {
    const current = tr.doc.nodeAt(pos)!;
    tr.setNodeMarkup(pos, editor.schema.nodes.docListItem, { ...current.attrs, kind: format.list.kind, numId: format.list.numId, ilvl: format.list.level });
  }
}

/** 功能：目标拖选完成后一次提交字符和段落格式，撤销一步恢复；时间：2026-11-10；作者：lq。 */
export function applyWordFormatPainter(editor: Editor, format: WordFormatSnapshot): boolean {
  if (!editor.isEditable || editor.view.composing || editor.state.selection.empty) return false;
  const tr = closeHistory(editor.state.tr);
  appendWordFormat(tr, editor, format);
  if (!tr.docChanged) return false;
  editor.view.dispatch(tr);
  return true;
}
