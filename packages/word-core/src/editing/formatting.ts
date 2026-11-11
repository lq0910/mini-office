/** 功能：读取选区格式并以单个事务修改字符/段落，不经过旧正文模型；时间：2026-11-10；作者：lq。 */
import type { Editor } from '@tiptap/core';
import type { Mark, Node } from '@tiptap/pm/model';
import { closeHistory } from '@tiptap/pm/history';
import type { ParsedDocFull } from '../upstream/packages/docx-engine/src/index';
import { HIGHLIGHT_CSS } from '../upstream/apps/docs/src/renderer/editor/marks';

/** 功能：只暴露工具栏需要的格式摘要，单位沿用内核的半磅和 twips；时间：2026-11-10；作者：lq。 */
export interface WordTextFormat {
  font?: string; fontAscii?: string; sizeHalfPoints?: number;
  bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean;
  color?: string; highlight?: string;
}
/** 功能：原生段落属性补丁，未提供的属性保持原值；时间：2026-11-10；作者：lq。 */
export interface WordParagraphFormat {
  /** 功能：段落边框、底纹和制表位摘要可供样式和格式刷复用；时间：2026-11-11；作者：lq。 */
  borders?: string | null; borderLines?: string | null; shadingFill?: string | null; tabStops?: string | null;
  align?: string | null; indentLeft?: number | null; indentRight?: number | null;
  indentFirstLine?: number | null; indentFirstLineChars?: number | null;
  spaceBefore?: number | null; spaceAfter?: number | null;
  spaceBeforeAuto?: boolean | null; spaceAfterAuto?: boolean | null;
  lineRule?: string | null; lineRawTwips?: number | null; lineSpacing?: number | null;
  pageBreakBefore?: boolean; keepNext?: boolean | null; keepLines?: boolean | null;
  widowControl?: boolean | null; snapToGrid?: boolean | null;
}

/** 功能：枚举真实选区相交的文本段，支持表格多格选区且不误改选区末端下一段；时间：2026-11-10；作者：lq。 */
export function selectedParagraphs(editor: Editor): Array<{ node: Node; pos: number }> {
  const found = new Map<number, Node>();
  const { selection, doc } = editor.state;
  if (selection.empty && selection.$from.parent.isTextblock) {
    found.set(selection.$from.before(), selection.$from.parent);
  } else for (const { $from, $to } of selection.ranges) {
    doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.isTextblock && pos + 1 < $to.pos && pos + node.nodeSize > $from.pos) found.set(pos, node);
    });
  }
  return [...found].map(([pos, node]) => ({ pos, node }));
}

/** 功能：把各文本片段的相同值与混合字段分开返回，禁止用选区第一个字符伪装统一格式；时间：2026-11-10；作者：lq。 */
function common<T extends object>(items: T[]): { value: T; mixed: Array<keyof T> } {
  const value = { ...items[0] } as T;
  const mixed: Array<keyof T> = [];
  for (const key of Object.keys(value) as Array<keyof T>) {
    if (items.some(item => item[key] !== value[key])) { delete value[key]; mixed.push(key); }
  }
  return { value, mixed };
}

/** 功能：从模型样式链、直接属性与字符标记读取格式，不读取 DOM/CSS；时间：2026-11-10；作者：lq。 */
export function readWordFormatting(editor: Editor, parsed: ParsedDocFull) {
  const paragraphs = selectedParagraphs(editor);
  const texts: WordTextFormat[] = [];
  const paras: WordParagraphFormat[] = [];
  const defaults = parsed.docDefaults;
  for (const { node, pos } of paragraphs) {
    const a = node.attrs;
    const display = parsed.styles.get(String(a.styleId ?? 'Normal'))?.display;
    const inherited = { ...defaults, ...display };
    // 功能：只比较当前行距模式的有效值，忽略被固定行距覆盖的继承倍数；时间：2026-11-10；作者：lq。
    const rule = a.lineRule ?? inherited.lineRule ?? 'auto';
    paras.push({
      align: a.align ?? display?.align ?? 'left',
      indentLeft: a.indentLeft ?? inherited.indentLeftTwips ?? 0,
      indentRight: a.indentRight ?? inherited.indentRightTwips ?? 0,
      indentFirstLine: a.indentFirstLine ?? inherited.indentFirstLineTwips ?? 0,
      spaceBefore: a.spaceBefore ?? inherited.spaceBeforeTwips ?? 0,
      spaceAfter: a.spaceAfter ?? inherited.spaceAfterTwips ?? 0,
      lineRule: rule,
      lineRawTwips: rule === 'auto' ? null : a.lineRawTwips ?? inherited.lineRawTwips ?? null,
      lineSpacing: rule === 'auto' ? a.lineSpacing ?? inherited.lineSpacing ?? 1 : null,
      pageBreakBefore: a.pageBreakBefore === true,
      keepNext: a.keepNext ?? display?.keepNext ?? false,
      keepLines: a.keepLines ?? display?.keepLines ?? false,
      widowControl: a.widowControl ?? inherited.widowControl ?? true,
      snapToGrid: a.snapToGrid ?? true,
    });
    /** 功能：字符样式覆盖段落样式，显式取消粗斜体优先于继承值；时间：2026-11-10；作者：lq。 */
    const read = (marks: readonly Mark[]) => {
      const attr = marks.find(mark => mark.type.name === 'docTextStyle')?.attrs ?? {};
      const character = parsed.styles.get(String(attr.styleId ?? ''))?.display;
      const base = { ...inherited, ...character };
      let table: Record<string, unknown> = {};
      try { table = JSON.parse(String(attr.tblRun ?? '{}')); } catch { /* 功能：无效的显示缓存不影响正文；时间：2026-11-10；作者：lq。 */ }
      const has = (name: string) => marks.some(mark => mark.type.name === name);
      const color = attr.color ?? base.color;
      texts.push({
        font: attr.font ?? attr.eastAsiaFont ?? base.font ?? base.eastAsiaFont ?? defaults?.eastAsiaFont,
        fontAscii: attr.fontAscii ?? base.fontAscii ?? defaults?.asciiFont,
        sizeHalfPoints: attr.sizeHalfPoints ?? table.sizeHalfPoints ?? base.sizeHalfPoints ?? 22,
        bold: attr.boldOff ? false : has('bold') || base.bold === true,
        italic: attr.italicOff ? false : has('italic') || base.italic === true,
        underline: attr.underlineOff ? false : has('underline') || base.underline === true, strike: attr.strikeOff ? false : has('strike') || base.strike === true,
        color: color && color !== 'auto' ? `#${String(color).replace(/^#/, '')}` : undefined,
        highlight: attr.highlight ? HIGHLIGHT_CSS[String(attr.highlight)] ?? String(attr.highlight) : attr.shading ? `#${attr.shading}` : base.shading ? `#${base.shading}` : undefined,
      });
    };
    if (editor.state.selection.empty) read(editor.state.storedMarks ?? editor.state.selection.$from.marks());
    else node.descendants((child, offset) => {
      if (child.isText && editor.state.selection.ranges.some(({ $from, $to }) => pos + 1 + offset < $to.pos && pos + 1 + offset + child.nodeSize > $from.pos)) read(child.marks);
    });
    if (!node.content.size && !editor.state.selection.empty) read([]);
  }
  return { text: common(texts), paragraph: common(paras) };
}

/** 功能：逐个文本片段合并字符属性，保留各自语言、主题、未知 rPr 及其他未改字段；时间：2026-11-10；作者：lq。 */
export function applyWordTextFormat(editor: Editor, patch: Record<string, unknown>): boolean {
  if (!editor.isEditable || editor.view.composing) return false;
  const { state } = editor;
  const type = state.schema.marks.docTextStyle;
  if (!type) return false;
  const tr = closeHistory(state.tr);
  if (state.selection.empty) {
    const marks = state.storedMarks ?? state.selection.$from.marks();
    tr.addStoredMark(type.create({ ...marks.find(mark => mark.type === type)?.attrs, ...patch }));
  } else for (const { $from, $to } of state.selection.ranges) {
    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (!node.isText) return;
      tr.addMark(Math.max(pos, $from.pos), Math.min(pos + node.nodeSize, $to.pos), type.create({ ...node.marks.find(mark => mark.type === type)?.attrs, ...patch }));
    });
  }
  editor.view.dispatch(tr);
  return true;
}

/** 功能：段落补丁一次提交并保留来源定位，单位换算由现有 UI 回调负责；时间：2026-11-10；作者：lq。 */
export function applyWordParagraphFormat(editor: Editor, patch: WordParagraphFormat): boolean {
  if (!editor.isEditable || editor.view.composing) return false;
  const tr = closeHistory(editor.state.tr);
  for (const { node, pos } of selectedParagraphs(editor)) tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...patch });
  if (!tr.docChanged) return false;
  editor.view.dispatch(tr);
  return true;
}

/** 功能：粗斜体切换清理显式关闭标记，保留选区内各片段的未知字符属性；时间：2026-11-10；作者：lq。 */
export function toggleWordInlineFormat(editor: Editor, parsed: ParsedDocFull, name: 'bold' | 'italic' | 'underline' | 'strike'): boolean {
  if (!editor.isEditable || editor.view.composing) return false;
  const enabled = readWordFormatting(editor, parsed).text.value[name] !== true;
  const { state } = editor;
  const tr = closeHistory(state.tr);
  const type = state.schema.marks[name];
  const styleType = state.schema.marks.docTextStyle;
  const off = ({ bold: 'boldOff', italic: 'italicOff', underline: 'underlineOff', strike: 'strikeOff' } as const)[name];
  /** 功能：每个片段保留自身 rPr，避免整段覆盖字体及语言；时间：2026-11-10；作者：lq。 */
  const style = (marks: readonly Mark[]) => styleType.create({ ...marks.find(mark => mark.type === styleType)?.attrs, ...(off ? { [off]: !enabled || null } : {}) });
  if (state.selection.empty) {
    const marks = state.storedMarks ?? state.selection.$from.marks();
    if (enabled) tr.addStoredMark(type.create()); else tr.removeStoredMark(type);
    if (off) tr.addStoredMark(style(marks));
  } else for (const { $from, $to } of state.selection.ranges) {
    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (!node.isText) return;
      const start = Math.max(pos, $from.pos), end = Math.min(pos + node.nodeSize, $to.pos);
      if (enabled) tr.addMark(start, end, type.create()); else tr.removeMark(start, end, type);
      if (off) tr.addMark(start, end, style(node.marks));
    });
  }
  editor.view.dispatch(tr);
  return true;
}
