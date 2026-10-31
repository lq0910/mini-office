import {
  type DocBlock,
  type DocComment,
  type DocImage,
  type DocModel,
  type DocParagraph,
  type DocRun,
  type DocStyleDefinition,
  type DocTable,
  type DocTableCell,
  type DocTableRow,
  type FieldMarker,
  type HeaderFooter,
  type ListInfo,
  type PageSetup,
  type ParagraphStyle,
  type RunControl,
  type RunStyle,
  type TableBorderSide,
  type TableBorders,
  type WordTabStop,
} from '@mini-office/document-model';
import {
  cssColor,
  firstChild,
  firstDescendant,
  getAttr,
  getAttrLocal,
  loadOoxmlPackage,
  localName,
  mimeForPath,
  numericAttr,
  parseXmlDom,
  elementText,
  readThemeColors,
  readThemeFonts,
  relationshipMap,
  toggleValue,
  twipsToPx,
  emuToPx,
  type OoxmlPackage,
  type Relationship,
} from '@mini-office/ooxml';
import { countWords, uid } from '@mini-office/shared';
// 功能：解析 Word 原生合并域（含 NEXT/NEXTIF/SKIPIF/MERGEREC/MERGESEQ 控制域与格式开关）；时间：2026-10-24；作者：lq。
import { parseMergeDirective, parseMergeField, type MergeDirective } from './merge-fields';

const HIGHLIGHT_COLORS: Record<string, string | undefined> = {
  black: '#000000',
  blue: '#0000FF',
  cyan: '#00FFFF',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGray: '#808080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  green: '#00FF00',
  lightGray: '#C0C0C0',
  magenta: '#FF00FF',
  none: undefined,
  red: '#FF0000',
  white: '#FFFFFF',
  yellow: '#FFFF00',
};

interface NumberingLevel {
  ilvl: number;
  numFmt: string;
  lvlText: string;
  start: number;
  indentLeftPx?: number;
  firstLineIndentPx?: number;
}

interface ResolvedStyle {
  run: RunStyle;
  para: ParagraphStyle;
  numId?: string;
  ilvl?: number;
}

interface RawStyle {
  id: string;
  type: string;
  /** 功能：样式显示名（w:name），供样式库展示与识别；时间：2026-10-29；作者：lq。 */
  name?: string;
  basedOn?: string;
  isDefault?: boolean;
  run: RunStyle;
  para: ParagraphStyle;
  numId?: string;
  ilvl?: number;
}

interface StyleSheet {
  docDefaults: { run: RunStyle; para: ParagraphStyle };
  styles: Map<string, ResolvedStyle>;
  defaultParagraph?: ResolvedStyle;
  /** 功能：styleId → 显示名 / 类型，用于把文档里实际用到的样式还原到样式库；时间：2026-10-29；作者：lq。 */
  names: Map<string, string>;
  types: Map<string, string>;
}

interface ParseContext {
  pkg: OoxmlPackage;
  rels: Map<string, Relationship>;
  styles: StyleSheet;
  numbering: Map<string, NumberingLevel[]>;
  counters: Map<string, Map<number, number>>;
  themeColors: Record<string, string>;
  themeFonts: { major: string; minor: string };
  defaultFont: string;
  defaultSizePt: number;
  imageUrls: Map<string, string>;
  usedImages: Map<string, { src: string; part?: string }>;
  pageWidthPx: number;
  blockCounter: number;
  hasTrackedChanges: boolean;
  hasContentControls: boolean;
  /** 功能：文档级复杂域栈——TOC 等域的 begin/end 常跨段落，靠它串联；时间：2026-10-30。 */
  fieldStack: FieldStackEntry[];
  /** 功能：页眉页脚按内容去重缓存——重模板的 125 个部件往往只有十几种内容；时间：2026-10-31。 */
  headerFooterCache?: Map<string, HeaderFooter>;
}

function mergeRun(base: RunStyle, override: RunStyle | undefined): RunStyle {
  if (!override) return { ...base };
  const result = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value !== undefined) (result as Record<string, unknown>)[key] = value;
  }
  return result;
}

function mergeParagraph(base: ParagraphStyle, override: ParagraphStyle | undefined): ParagraphStyle {
  if (!override) return { ...base };
  const result = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value !== undefined) (result as Record<string, unknown>)[key] = value;
  }
  return result;
}

function attrValue(element: Element | null | undefined, name: string): string | undefined {
  return getAttr(element, name);
}

function colorFromElement(element: Element | null, theme: Record<string, string>): string | undefined {
  if (!element) return undefined;
  const value = attrValue(element, 'w:val');
  if (value && value.toLowerCase() !== 'auto') {
    const hex = value.replace(/^#/, '');
    if (/^[0-9a-fA-F]{6}$/.test(hex)) return `#${hex.toUpperCase()}`;
    if (/^[0-9a-fA-F]{8}$/.test(hex)) return `#${hex.slice(0, 6).toUpperCase()}`;
  }
  const themeColor = attrValue(element, 'w:themeColor');
  if (!themeColor) return undefined;
  const schemeMap: Record<string, string> = {
    text1: 'tx1',
    background1: 'bg1',
    text2: 'tx2',
    background2: 'bg2',
    dark1: 'dk1',
    light1: 'lt1',
    dark2: 'dk2',
    light2: 'lt2',
    accent1: 'accent1',
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hyperlink: 'hlink',
    followedHyperlink: 'folHlink',
  };
  const base = theme[schemeMap[themeColor] ?? themeColor];
  if (!base) return undefined;
  const tintHex = attrValue(element, 'w:themeTint');
  const shadeHex = attrValue(element, 'w:themeShade');
  const toPercent = (hex?: string) => {
    if (!hex) return undefined;
    const parsed = parseInt(hex, 16);
    return Number.isFinite(parsed) ? (parsed / 255) * 100000 : undefined;
  };
  return cssColor({ hex: base, tint: toPercent(tintHex), shade: toPercent(shadeHex) }, theme);
}

function applyRunStyle(base: RunStyle, rPr: Element | null, ctx: ParseContext): RunStyle {
  if (!rPr) return { ...base };
  const style: RunStyle = { ...base };
  const fonts = firstChild(rPr, 'w:rFonts');
  // 功能：Word 中西文字体分槽（w:ascii / w:eastAsia），分别记录避免中文样式被西文覆盖；时间：2026-10-30。
  const asciiFont =
    attrValue(fonts, 'w:ascii') ?? attrValue(fonts, 'w:hAnsi') ?? attrValue(fonts, 'w:cs');
  const eastAsiaAttr = attrValue(fonts, 'w:eastAsia');
  const font = asciiFont ?? eastAsiaAttr;
  if (font) style.fontFamily = font;
  if (eastAsiaAttr) style.eastAsiaFont = eastAsiaAttr;

  const size = numericAttr(firstChild(rPr, 'w:sz'), 'w:val');
  const sizeCs = numericAttr(firstChild(rPr, 'w:szCs'), 'w:val');
  if (size) style.fontSizePt = size / 2;
  else if (sizeCs) style.fontSizePt = sizeCs / 2;

  const bold = firstChild(rPr, 'w:b');
  if (bold) style.bold = toggleValue(bold);
  const italic = firstChild(rPr, 'w:i');
  if (italic) style.italic = toggleValue(italic);
  const underline = firstChild(rPr, 'w:u');
  if (underline) {
    const value = attrValue(underline, 'w:val');
    style.underline = value ? !['none', '0', 'false'].includes(value.toLowerCase()) : false;
  }
  const strike = firstChild(rPr, 'w:strike');
  if (strike) style.strike = toggleValue(strike);

  const color = colorFromElement(firstChild(rPr, 'w:color'), ctx.themeColors);
  if (color) style.color = color;
  const highlight = firstChild(rPr, 'w:highlight');
  if (highlight) {
    const value = attrValue(highlight, 'w:val');
    if (value && HIGHLIGHT_COLORS[value]) style.highlight = HIGHLIGHT_COLORS[value];
    else if (value === 'none') style.highlight = undefined;
  }
  const shading = firstChild(rPr, 'w:shd');
  const fill = attrValue(shading, 'w:fill');
  if (fill && fill.toLowerCase() !== 'auto' && /^[0-9a-fA-F]{6}$/.test(fill)) {
    style.highlight = `#${fill.toUpperCase()}`;
  }

  const vertAlign = attrValue(firstChild(rPr, 'w:vertAlign'), 'w:val');
  if (vertAlign === 'superscript') style.verticalAlign = 'super';
  else if (vertAlign === 'subscript') style.verticalAlign = 'sub';
  else if (vertAlign === 'baseline') style.verticalAlign = 'baseline';

  const spacing = numericAttr(firstChild(rPr, 'w:spacing'), 'w:val');
  if (spacing !== undefined) style.letterSpacingPx = twipsToPx(spacing);
  if (firstChild(rPr, 'w:caps')) style.allCaps = toggleValue(firstChild(rPr, 'w:caps'));
  if (firstChild(rPr, 'w:smallCaps')) style.smallCaps = toggleValue(firstChild(rPr, 'w:smallCaps'));
  return style;
}

function applyParagraphStyle(base: ParagraphStyle, pPr: Element | null, ctx: ParseContext): ParagraphStyle {
  if (!pPr) return { ...base };
  const style: ParagraphStyle = { ...base };
  const alignment = attrValue(firstChild(pPr, 'w:jc'), 'w:val');
  if (alignment) {
    const map: Record<string, ParagraphStyle['alignment']> = {
      left: 'left',
      start: 'left',
      right: 'right',
      end: 'right',
      center: 'center',
      both: 'justify',
      justify: 'justify',
      distribute: 'justify',
    };
    style.alignment = map[alignment] ?? base.alignment;
  }

  const indent = firstChild(pPr, 'w:ind');
  if (indent) {
    const left = numericAttr(indent, 'w:left') ?? numericAttr(indent, 'w:start');
    const right = numericAttr(indent, 'w:right') ?? numericAttr(indent, 'w:end');
    const firstLine = numericAttr(indent, 'w:firstLine');
    const hanging = numericAttr(indent, 'w:hanging');
    // 功能：WPS/中文模板常用字符单位缩进（leftChars/firstLineChars/hangingChars），按字号换算为 px，避免缩进丢失；时间：2026-10-30。
    const charPx = (ctx.defaultSizePt ?? 10.5) * (96 / 72);
    const leftChars = numericAttr(indent, 'w:leftChars') ?? numericAttr(indent, 'w:startChars');
    const rightChars = numericAttr(indent, 'w:rightChars') ?? numericAttr(indent, 'w:endChars');
    const firstLineChars = numericAttr(indent, 'w:firstLineChars');
    const hangingChars = numericAttr(indent, 'w:hangingChars');
    if (left !== undefined) style.indentLeftPx = twipsToPx(left);
    else if (leftChars !== undefined) style.indentLeftPx = (leftChars / 100) * charPx;
    if (right !== undefined) style.indentRightPx = twipsToPx(right);
    else if (rightChars !== undefined) style.indentRightPx = (rightChars / 100) * charPx;
    if (firstLine !== undefined) style.firstLineIndentPx = twipsToPx(firstLine);
    else if (hanging !== undefined) style.firstLineIndentPx = -twipsToPx(hanging);
    else if (firstLineChars !== undefined) style.firstLineIndentPx = (firstLineChars / 100) * charPx;
    else if (hangingChars !== undefined) style.firstLineIndentPx = -(hangingChars / 100) * charPx;
  }

  const tabs = firstChild(pPr, 'w:tabs');
  if (tabs) {
    const tabStops: WordTabStop[] = [];
    for (const child of Array.from(tabs.childNodes)) {
      if (!(child instanceof Element) || localName(child) !== 'tab') continue;
      const pos = numericAttr(child, 'w:pos');
      if (pos === undefined) continue;
      const val = attrValue(child, 'w:val');
      const kind: WordTabStop['kind'] =
        val === 'center' || val === 'right' || val === 'decimal' ? val : 'left';
      // 功能：解析制表符前导符（dot/hyphen/underline…），供目录等场景渲染「标题…页码」；时间：2026-10-30。
      const leader = attrValue(child, 'w:leader') as WordTabStop['leader'] | undefined;
      const tabStop: WordTabStop = { posPx: twipsToPx(pos), kind };
      if (leader) tabStop.leader = leader;
      tabStops.push(tabStop);
    }
    if (tabStops.length) style.tabStops = tabStops;
  }

  const spacing = firstChild(pPr, 'w:spacing');
  if (spacing) {
    const before = numericAttr(spacing, 'w:before');
    const after = numericAttr(spacing, 'w:after');
    const line = numericAttr(spacing, 'w:line');
    const rule = attrValue(spacing, 'w:lineRule') ?? 'auto';
    if (before !== undefined) style.spaceBeforePx = twipsToPx(before);
    if (after !== undefined) style.spaceAfterPx = twipsToPx(after);
    if (line) {
      if (rule === 'auto') style.lineHeight = line / 240;
      else if (rule === 'atLeast') style.lineHeightAtLeastPx = twipsToPx(line);
      else style.lineHeightExactPx = twipsToPx(line);
    }
  }

  const keepNext = firstChild(pPr, 'w:keepNext');
  if (keepNext) style.keepNext = toggleValue(keepNext);
  const pageBreakBefore = firstChild(pPr, 'w:pageBreakBefore');
  if (pageBreakBefore) style.pageBreakBefore = toggleValue(pageBreakBefore);
  // 功能：段落显式关闭「对齐到网格」时记下来，行高不再向上取整到 docGrid 行距；时间：2026-10-30。
  const snapToGrid = firstChild(pPr, 'w:snapToGrid');
  if (snapToGrid && !toggleValue(snapToGrid)) style.snapToGrid = false;
  // 功能：段落显式关闭孤寡控制时记下来（Word 默认开启）；时间：2026-10-31。
  const widowControl = firstChild(pPr, 'w:widowControl');
  if (widowControl && !toggleValue(widowControl)) style.widowControl = false;
  const outline = numericAttr(firstChild(pPr, 'w:outlineLvl'), 'w:val');
  if (outline !== undefined) style.outlineLevel = outline;

  const shading = firstChild(pPr, 'w:shd');
  const fill = attrValue(shading, 'w:fill');
  if (fill && /^[0-9a-fA-F]{6}$/.test(fill)) style.backgroundColor = `#${fill.toUpperCase()}`;
  // 功能：解析段落边框，恢复 Word 中常见的横线、签名线和分隔线；时间：2026-10-27；作者：lq。
  const paragraphBorders = firstChild(pPr, 'w:pBdr');
  if (paragraphBorders) style.borders = parseTableBorders(paragraphBorders);
  return style;
}

function parseProjectedParagraphNumbering(pPr: Element | null): { numId?: string; ilvl?: number } {
  const numPr = firstChild(pPr, 'w:numPr');
  if (!numPr) return {};
  return {
    numId: attrValue(firstChild(numPr, 'w:numId'), 'w:val'),
    ilvl: numericAttr(firstChild(numPr, 'w:ilvl'), 'w:val'),
  };
}

function parseStyleSheet(stylesXml: string, ctx: ParseContext): StyleSheet {
  const docDefaults: { run: RunStyle; para: ParagraphStyle } = { run: {}, para: {} };
  const rawStyles = new Map<string, RawStyle>();
  const names = new Map<string, string>();
  const types = new Map<string, string>();
  let defaultParagraphId: string | undefined;

  if (stylesXml) {
    const doc = parseXmlDom(stylesXml);
    const stylesRoot = firstDescendant(doc, 'w:styles') ?? doc.documentElement;
    const defaults = firstChild(stylesRoot, 'w:docDefaults');
    if (defaults) {
      const rPrDefault = firstChild(firstChild(defaults, 'w:rPrDefault') ?? defaults, 'w:rPr');
      const pPrDefault = firstChild(firstChild(defaults, 'w:pPrDefault') ?? defaults, 'w:pPr');
      docDefaults.run = applyRunStyle({}, rPrDefault, ctx);
      docDefaults.para = applyParagraphStyle({}, pPrDefault, ctx);
    }
    for (const styleElement of Array.from(stylesRoot.childNodes)) {
      if (!(styleElement instanceof Element) || localName(styleElement) !== 'style') continue;
      const id = attrValue(styleElement, 'w:styleId');
      if (!id) continue;
      const type = attrValue(styleElement, 'w:type') ?? 'paragraph';
      const basedOn = attrValue(firstChild(styleElement, 'w:basedOn'), 'w:val');
      const isDefault = attrValue(styleElement, 'w:default') === '1';
      const run = applyRunStyle({}, firstChild(styleElement, 'w:rPr'), ctx);
      const para = applyParagraphStyle({}, firstChild(styleElement, 'w:pPr'), ctx);
      const num = parseProjectedParagraphNumbering(firstChild(styleElement, 'w:pPr'));
      const styleName = attrValue(firstChild(styleElement, 'w:name'), 'w:val');
      if (styleName) names.set(id, styleName);
      types.set(id, type);
      rawStyles.set(id, { id, type, name: styleName, basedOn, isDefault, run, para, ...num });
      if (type === 'paragraph' && isDefault) defaultParagraphId = id;
    }
  }

  const resolved = new Map<string, ResolvedStyle>();
  const resolving = new Set<string>();

  const resolve = (id: string, stack: string[] = []): ResolvedStyle => {
    const cached = resolved.get(id);
    if (cached) return cached;
    if (resolving.has(id) || stack.length > 30) {
      return { run: { ...docDefaults.run }, para: { ...docDefaults.para } };
    }
    resolving.add(id);
    const raw = rawStyles.get(id);
    if (!raw) {
      resolving.delete(id);
      return { run: { ...docDefaults.run }, para: { ...docDefaults.para } };
    }
    const base = raw.basedOn ? resolve(raw.basedOn, [...stack, id]) : { run: { ...docDefaults.run }, para: { ...docDefaults.para } };
    const result: ResolvedStyle = {
      run: mergeRun(base.run, raw.run),
      para: mergeParagraph(base.para, raw.para),
      numId: raw.numId ?? base.numId,
      ilvl: raw.ilvl ?? base.ilvl,
    };
    resolved.set(id, result);
    resolving.delete(id);
    return result;
  };

  rawStyles.forEach((_style, id) => resolve(id));
  const defaultParagraph = defaultParagraphId
    ? resolved.get(defaultParagraphId)
    : resolved.get('Normal');
  return { docDefaults, styles: resolved, defaultParagraph, names, types };
}

/**
 * 功能：把文档中实际使用到的段落样式还原为样式库定义，让导入文档的样式可直接复用；时间：2026-10-29；作者：lq。
 * 只收段落样式且最多 40 个，避免带一大堆无用样式进内存与存储。
 */
function collectDocumentStyles(
  blocks: DocBlock[],
  header: HeaderFooter | undefined,
  footer: HeaderFooter | undefined,
  sheet: StyleSheet,
): DocStyleDefinition[] {
  const used = new Set<string>();
  const visit = (items: DocBlock[]) => {
    for (const block of items) {
      if (block.type === 'paragraph') {
        if (block.styleId) used.add(block.styleId);
      } else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
      }
    }
  };
  visit(blocks);
  if (header) visit(header.blocks);
  if (footer) visit(footer.blocks);

  const result: DocStyleDefinition[] = [];
  for (const id of used) {
    if (result.length >= 40) break;
    if (id === 'Normal') continue;
    if (sheet.types.get(id) !== 'paragraph') continue;
    const resolved = sheet.styles.get(id);
    if (!resolved) continue;
    const name = sheet.names.get(id) ?? id;
    result.push({
      id,
      name,
      paragraph: { ...resolved.para },
      run: { ...resolved.run },
      origin: 'imported',
    });
  }
  return result;
}

function parseNumbering(numberingXml: string): Map<string, NumberingLevel[]> {
  const result = new Map<string, NumberingLevel[]>();
  if (!numberingXml) return result;
  const doc = parseXmlDom(numberingXml);
  const root = firstDescendant(doc, 'w:numbering') ?? doc.documentElement;
  const abstractLevels = new Map<string, NumberingLevel[]>();
  for (const abstractNum of Array.from(root.childNodes)) {
    if (!(abstractNum instanceof Element) || localName(abstractNum) !== 'abstractNum') continue;
    const id = attrValue(abstractNum, 'w:abstractNumId');
    if (!id) continue;
    const levels: NumberingLevel[] = [];
    for (const lvl of Array.from(abstractNum.childNodes)) {
      if (!(lvl instanceof Element) || localName(lvl) !== 'lvl') continue;
      const ilvl = numericAttr(lvl, 'w:ilvl') ?? levels.length;
      const pPr = firstChild(lvl, 'w:pPr');
      const ind = firstChild(pPr, 'w:ind');
      const left = numericAttr(ind, 'w:left') ?? numericAttr(ind, 'w:start');
      const hanging = numericAttr(ind, 'w:hanging');
      const firstLine = numericAttr(ind, 'w:firstLine');
      levels.push({
        ilvl,
        numFmt: attrValue(firstChild(lvl, 'w:numFmt'), 'w:val') ?? 'decimal',
        lvlText: attrValue(firstChild(lvl, 'w:lvlText'), 'w:val') ?? '',
        start: numericAttr(firstChild(lvl, 'w:start'), 'w:val') ?? 1,
        indentLeftPx: left !== undefined ? twipsToPx(left) : undefined,
        firstLineIndentPx:
          firstLine !== undefined
            ? twipsToPx(firstLine)
            : hanging !== undefined
              ? -twipsToPx(hanging)
              : undefined,
      });
    }
    abstractLevels.set(id, levels);
  }
  for (const num of Array.from(root.childNodes)) {
    if (!(num instanceof Element) || localName(num) !== 'num') continue;
    const numId = attrValue(num, 'w:numId');
    const abstractId = attrValue(firstChild(num, 'w:abstractNumId'), 'w:val');
    if (!numId || !abstractId) continue;
    result.set(numId, abstractLevels.get(abstractId) ?? []);
  }
  return result;
}

function toChineseNumber(value: number): string {
  if (value <= 0) return String(value);
  const digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
  const units = ['', '十', '百', '千'];
  if (value < 10) return digits[value];
  if (value < 20) return value === 10 ? '十' : `十${digits[value % 10]}`;
  if (value < 10000) {
    const text = String(value);
    let result = '';
    let zero = false;
    for (let index = 0; index < text.length; index += 1) {
      const digit = Number(text[index]);
      const unit = units[text.length - index - 1];
      if (digit === 0) {
        zero = true;
        continue;
      }
      if (zero && result) result += '零';
      zero = false;
      result += `${digits[digit]}${unit}`;
    }
    return result;
  }
  return String(value);
}

function toRoman(value: number, upper: boolean): string {
  const numerals: Array<[number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let remaining = Math.max(1, Math.min(3999, Math.floor(value)));
  let text = '';
  for (const [number, symbol] of numerals) {
    while (remaining >= number) {
      text += symbol;
      remaining -= number;
    }
  }
  return upper ? text : text.toLowerCase();
}

function toLetters(value: number, upper: boolean): string {
  let n = Math.max(1, Math.floor(value));
  let text = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    text = String.fromCharCode((upper ? 65 : 97) + rem) + text;
    n = Math.floor((n - 1) / 26);
  }
  return text;
}

function formatListNumber(value: number, format: string): string {
  switch (format) {
    case 'lowerLetter':
      return toLetters(value, false);
    case 'upperLetter':
      return toLetters(value, true);
    case 'lowerRoman':
      return toRoman(value, false);
    case 'upperRoman':
      return toRoman(value, true);
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'japaneseCounting':
      return toChineseNumber(value);
    case 'decimalZero':
      return String(value).padStart(2, '0');
    case 'decimal':
    default:
      return String(value);
  }
}

function listInfoFor(ctx: ParseContext, numId: string, ilvl: number): ListInfo | undefined {
  const levels = ctx.numbering.get(numId);
  if (!levels || levels.length === 0) return undefined;
  let counters = ctx.counters.get(numId);
  if (!counters) {
    counters = new Map();
    ctx.counters.set(numId, counters);
  }
  const level = levels.find((item) => item.ilvl === ilvl) ?? levels[0];
  const current = (counters.get(ilvl) ?? level.start - 1) + 1;
  counters.set(ilvl, current);
  for (const key of [...counters.keys()]) {
    if (key > ilvl) counters.delete(key);
  }
  const values = new Map<number, number>();
  for (const item of levels) {
    values.set(item.ilvl, counters.get(item.ilvl) ?? item.start);
  }
  values.set(ilvl, current);
  // 功能：numFmt="none" 的编号定义在 Word 里显示为无编号——标记置空，渲染与分段都不再产生数字；时间：2026-10-31。
  if (level.numFmt === 'none') {
    return { numId, level: ilvl, marker: '', ordered: false };
  }
  let marker = level.lvlText;
  if (level.numFmt === 'bullet') {
    marker = marker || '•';
  } else {
    marker = marker.replace(/%([1-9])/g, (_match, index: string) => {
      const target = Number(index) - 1;
      const targetLevel = levels.find((item) => item.ilvl === target);
      return formatListNumber(values.get(target) ?? targetLevel?.start ?? 1, targetLevel?.numFmt ?? 'decimal');
    });
    if (!marker) marker = `${formatListNumber(current, level.numFmt)}.`;
  }
  return { numId, level: ilvl, marker, ordered: level.numFmt !== 'bullet' && level.numFmt !== 'none' };
}

function parseTableBorder(element: Element | null): TableBorderSide | undefined {
  if (!element) return undefined;
  const value = attrValue(element, 'w:val') ?? 'single';
  if (['nil', 'none'].includes(value)) {
    return { style: 'none', widthPx: 0, color: 'transparent' };
  }
  const size = numericAttr(element, 'w:sz', 4) ?? 4;
  const colorRaw = attrValue(element, 'w:color');
  const color =
    colorRaw && colorRaw.toLowerCase() !== 'auto' && /^[0-9a-fA-F]{6}$/.test(colorRaw)
      ? `#${colorRaw.toUpperCase()}`
      : '#000000';
  const styleMap: Record<string, TableBorderSide['style']> = {
    single: 'single',
    double: 'double',
    dashed: 'dashed',
    dotted: 'dotted',
    thick: 'single',
  };
  return {
    style: styleMap[value] ?? 'single',
    widthPx: Math.max(0.5, size / 6),
    color,
  };
}

function parseTableBorders(element: Element | null): TableBorders {
  if (!element) return {};
  return {
    top: parseTableBorder(firstChild(element, 'w:top')),
    right: parseTableBorder(firstChild(element, 'w:right')),
    bottom: parseTableBorder(firstChild(element, 'w:bottom')),
    left: parseTableBorder(firstChild(element, 'w:left')),
    insideH: parseTableBorder(firstChild(element, 'w:insideH')),
    insideV: parseTableBorder(firstChild(element, 'w:insideV')),
  };
}

async function loadImageUrl(ctx: ParseContext, part: string): Promise<string | undefined> {
  const cached = ctx.imageUrls.get(part);
  if (cached) return cached;
  if (!ctx.pkg.has(part)) return undefined;
  try {
    const bytes = await ctx.pkg.bytes(part);
    const buffer = new Uint8Array(bytes.byteLength);
    buffer.set(bytes);
    const url = URL.createObjectURL(new Blob([buffer], { type: mimeForPath(part) }));
    ctx.imageUrls.set(part, url);
    return url;
  } catch {
    return undefined;
  }
}

function vmlStyleSize(style: string | undefined): { width: number; height: number } | undefined {
  if (!style) return undefined;
  const width = /(?:^|;)\s*width\s*:\s*([\d.]+)(pt|px|in|cm)/i.exec(style);
  const height = /(?:^|;)\s*height\s*:\s*([\d.]+)(pt|px|in|cm)/i.exec(style);
  if (!width || !height) return undefined;
  const convert = (match: RegExpExecArray) => {
    const value = Number(match[1]);
    switch (match[2].toLowerCase()) {
      case 'pt':
        return value * (4 / 3);
      case 'in':
        return value * 96;
      case 'cm':
        return (value / 2.54) * 96;
      default:
        return value;
    }
  };
  return { width: convert(width), height: convert(height) };
}

/** 功能：解析 VML rect 的绝对定位信息（margin-left/margin-top/width/height），单位 CSS px；支持负值 margin；时间：2026-10-30。 */
function parseVmlRectStyle(style: string): { left: number; top: number; width: number; height: number; xRelative: 'page' | 'text' } | undefined {
  const position = /(?:^|;)\s*position\s*:\s*absolute/i.test(style);
  if (!position) return undefined;
  const left = /(?:^|;)\s*margin-left\s*:\s*(-?[\d.]+)(pt|px|in|cm)/i.exec(style);
  const top = /(?:^|;)\s*margin-top\s*:\s*(-?[\d.]+)(pt|px|in|cm)/i.exec(style);
  const width = /(?:^|;)\s*width\s*:\s*(-?[\d.]+)(pt|px|in|cm)/i.exec(style);
  const height = /(?:^|;)\s*height\s*:\s*(-?[\d.]+)(pt|px|in|cm)/i.exec(style);
  if (!left || !top || !width || !height) return undefined;
  const convert = (match: RegExpExecArray) => {
    const value = Number(match[1]);
    switch (match[2].toLowerCase()) {
      case 'pt':
        return value * (4 / 3);
      case 'in':
        return value * 96;
      case 'cm':
        return (value / 2.54) * 96;
      default:
        return value;
    }
  };
  // 功能：mso-position-horizontal-relative:page 表示横坐标相对页面左缘，缺省相对内容区左缘；时间：2026-10-30。
  const horizontalRelative = /mso-position-horizontal-relative\s*:\s*([a-z]+)/i.exec(style)?.[1]?.toLowerCase();
  return {
    left: convert(left),
    top: convert(top),
    width: convert(width),
    height: convert(height),
    xRelative: horizontalRelative === 'page' ? 'page' : 'text',
  };
}

async function parseDrawing(drawing: Element, ctx: ParseContext): Promise<DocRun[]> {
  // 功能：DrawingML 文本框（wps:txbx）——提取内容段落与锚定几何，只读渲染；时间：2026-10-31。
  const txbx = firstDescendant(drawing, 'wps:txbx') ?? firstDescendant(drawing, 'w:txbxContent');
  if (txbx || firstDescendant(drawing, 'w:txbxContent')) {
    const content = firstDescendant(drawing, 'w:txbxContent');
    if (content) {
      const blocks = await parseBlockContainer(content, ctx);
      const extent = firstDescendant(drawing, 'wp:extent');
      const width = Math.max(8, emuToPx(numericAttr(extent, 'cx', 1_828_800) ?? 1_828_800));
      const height = Math.max(8, emuToPx(numericAttr(extent, 'cy', 1_371_600) ?? 1_371_600));
      const positionH = firstDescendant(drawing, 'wp:positionH');
      const positionV = firstDescendant(drawing, 'wp:positionV');
      const hRelative = attrValue(positionH, 'relativeFrom');
      // 功能：wp:posOffset 的偏移是元素文本（EMU），不是属性；时间：2026-10-31。
      const leftEmu = numericAttr(positionH, 'x') ?? (Number.parseInt(elementText(firstDescendant(positionH ?? null, 'wp:posOffset')) || '0', 10) || 0);
      const topEmu = numericAttr(positionV, 'y') ?? (Number.parseInt(elementText(firstDescendant(positionV ?? null, 'wp:posOffset')) || '0', 10) || 0);
      return [{
        text: '',
        style: {},
        textbox: {
          left: emuToPx(leftEmu),
          top: emuToPx(topEmu),
          width,
          height,
          xRelative: hRelative === 'page' ? 'page' : 'text',
          blocks,
        },
      }];
    }
  }
  const blip = firstDescendant(drawing, 'a:blip');
  const rid = attrValue(blip, 'r:embed') ?? attrValue(blip, 'r:link');
  if (!rid) return [];
  const rel = ctx.rels.get(rid);
  const part = rel?.resolved;
  if (!part) return [];
  const src = await loadImageUrl(ctx, part);
  if (!src) return [];
  const extent = firstDescendant(drawing, 'wp:extent');
  const widthPx = Math.max(8, emuToPx(numericAttr(extent, 'cx', 1_828_800) ?? 1_828_800));
  const heightPx = Math.max(8, emuToPx(numericAttr(extent, 'cy', 1_371_600) ?? 1_371_600));
  const docPr = firstDescendant(drawing, 'wp:docPr');
  const alt = attrValue(docPr, 'descr') ?? attrValue(docPr, 'title');
  const name = attrValue(docPr, 'name');
  const image: DocImage = {
    id: uid('img'),
    src,
    widthPx,
    heightPx,
    alt,
    name,
    relId: rid,
    part,
    // 功能：识别 Word 图片的内嵌/浮动形态，重新打开后保持文字环绕；时间：2026-10-24；作者：lq。
    wrap: drawing.getElementsByTagName('wp:anchor').length
      ? (firstDescendant(drawing, 'wp:wrapTopAndBottom') ? 'top-bottom' : attrValue(firstDescendant(drawing, 'wp:wrapSquare'), 'wrapText') === 'right' ? 'square-right' : 'square-left')
      : 'inline',
  };
  ctx.usedImages.set(rid, { src, part });
  return [{ text: '', style: {}, image }];
}

/**
 * 功能：从 w:sdtPr 读出内容控件语义——复选框勾选状态、下拉/组合框候选、日期格式与别名；
 * 未识别的控件按普通文本处理；时间：2026-10-30。
 */
function parseSdtControl(sdt: Element): RunControl | undefined {
  const pr = firstChild(sdt, 'w:sdtPr');
  if (!pr) return undefined;
  const alias = attrValue(firstChild(pr, 'w:alias'), 'w:val');
  const tag = attrValue(firstChild(pr, 'w:tag'), 'w:val');
  const locked = toggleValue(firstChild(pr, 'w:lock'));
  const base: RunControl = {
    kind: 'text',
    ...(alias ? { alias } : {}),
    ...(tag ? { tag } : {}),
    ...(locked ? { locked } : {}),
  };
  // 功能：复选框控件（w14:checkbox）读取 w14:checked；时间：2026-10-30。
  const checkbox = firstDescendant(pr, 'checkbox') ?? firstChild(pr, 'w:checkBox');
  if (checkbox) {
    const checkedNode =
      firstDescendant(checkbox, 'checked') ?? firstChild(checkbox, 'w:checked');
    return { ...base, kind: 'checkbox', checked: toggleValue(checkedNode) };
  }
  const dropdown = firstChild(pr, 'w:dropDownList');
  if (dropdown) {
    const choices = Array.from(dropdown.getElementsByTagName('w:listItem'))
      .map((item) => attrValue(item, 'w:displayText') ?? attrValue(item, 'w:value'))
      .filter((value): value is string => !!value);
    return { ...base, kind: 'dropdown', choices };
  }
  const comboBox = firstChild(pr, 'w:comboBox');
  if (comboBox) {
    const choices = Array.from(comboBox.getElementsByTagName('w:listItem'))
      .map((item) => attrValue(item, 'w:displayText') ?? attrValue(item, 'w:value'))
      .filter((value): value is string => !!value);
    return { ...base, kind: 'comboBox', choices };
  }
  const date = firstChild(pr, 'w:date');
  if (date) {
    const format =
      attrValue(firstChild(date, 'w:dateFormat'), 'w:val') ??
      attrValue(firstChild(date, 'w:dateFormat'), 'w:valFormat');
    return { ...base, kind: 'date', ...(format ? { dateFormat: format } : {}) };
  }
  const text = firstChild(pr, 'w:text');
  if (text) return { ...base, kind: 'text' };
  // 未带具体类型的 sdt（如富文本控件）也保留包装语义，便于写回。
  return base;
}

/** 功能：VML 方框内是否已画勾（v:textbox 内的 ✓），用于恢复勾选状态；时间：2026-10-30。 */
function vmlRectChecked(rect: Element): boolean {
  const textbox = firstDescendant(rect, 'v:textbox');
  if (!textbox) return false;
  const text = textbox.textContent ?? '';
  return /[✓✔☑√]/.test(text);
}

async function parseVmlPicture(pict: Element, ctx: ParseContext): Promise<DocRun[]> {
  // 功能：VML 文本框（v:shape 或无定位的 v:rect 内嵌 v:textbox）——解析内容与几何，只读渲染；
  // 带定位的 v:rect 是复选框语义（勾选态存在 v:textbox 里），走下方 vmlRect 分支；时间：2026-10-31。
  const rectForTextbox = firstDescendant(pict, 'v:rect');
  const rectIsCheckbox = rectForTextbox ? !!parseVmlRectStyle(attrValue(rectForTextbox, 'style') ?? '') : false;
  const textShape = firstDescendant(pict, 'v:shape');
  const textHost = (textShape ?? (rectForTextbox && !rectIsCheckbox ? rectForTextbox : null)) as Element | null;
  if (textHost && firstDescendant(textHost, 'v:textbox')) {
    const content = firstDescendant(textHost, 'w:txbxContent');
    if (content) {
      const blocks = await parseBlockContainer(content, ctx);
      const style = attrValue(textHost, 'style') ?? '';
      const rectStyle = parseVmlRectStyle(style);
      const size = rectStyle ?? vmlStyleSize(style);
      return [{
        text: '',
        style: {},
        textbox: {
          left: rectStyle?.left ?? 0,
          top: rectStyle?.top ?? 0,
          width: size?.width ?? 120,
          height: size?.height ?? 80,
          xRelative: rectStyle?.xRelative ?? 'text',
          blocks,
        },
      }];
    }
  }
  // 功能：VML 复选框（v:rect）按定位信息渲染；时间：2026-10-30。
  const rect = firstDescendant(pict, 'v:rect');
  if (rect) {
    const style = attrValue(rect, 'style') ?? '';
    const rectStyle = parseVmlRectStyle(style);
    // 有定位信息的 rect 渲染为绝对定位浮层；仅无定位信息的才退化为 □ 字符；
    if (rectStyle) {
      const strokeWeight = /strokeweight\s*:\s*(-?[\d.]+)(pt|px)/i.exec(style);
      const strokeColor = attrValue(rect, 'strokecolor');
      return [
        {
          text: '',
          style: {},
          image: {
            id: uid('img'),
            src: '',
            widthPx: rectStyle.width,
            heightPx: rectStyle.height,
            vmlRect: { ...rectStyle, ...(vmlRectChecked(rect) ? { checked: true } : {}) },
            strokeColor: strokeColor && /^#?[0-9a-fA-F]{6}$/.test(strokeColor)
              ? `#${strokeColor.replace('#', '').toUpperCase()}`
              : undefined,
            // 功能：描边粗细按 strokeweight 换算 px（pt→px 乘 4/3），缺省 0.72pt；时间：2026-10-30。
            strokeWidthPx: strokeWeight
              ? Number(strokeWeight[1]) * (strokeWeight[2].toLowerCase() === 'pt' ? 4 / 3 : 1)
              : 1,
            wrap: 'none',
          },
        },
      ];
    }
    return [{ text: '□', style: {} }];
  }
  const imageData = firstDescendant(pict, 'v:imagedata');
  const rid = attrValue(imageData, 'r:id') ?? attrValue(imageData, 'r:embed');
  if (!rid) return [];
  const rel = ctx.rels.get(rid);
  const part = rel?.resolved;
  if (!part) return [];
  const src = await loadImageUrl(ctx, part);
  if (!src) return [];
  const shape = firstDescendant(pict, 'v:shape');
  const size = vmlStyleSize(attrValue(shape, 'style'));
  const image: DocImage = {
    id: uid('img'),
    src,
    widthPx: size?.width ?? 160,
    heightPx: size?.height ?? 120,
    part,
    relId: rid,
    name: attrValue(shape, 'alt') ?? undefined,
  };
  ctx.usedImages.set(rid, { src, part });
  return [{ text: '', style: {}, image }];
}

interface ParseInlineState {
  /** 功能：文档级复杂域栈，跨段落串联 TOC 等域的 begin/end；时间：2026-10-30。 */
  fieldStack: FieldStackEntry[];
  /** 功能：本段落已产出的 run 数，用于把域标记定位到段内位置；时间：2026-10-30。 */
  runCount?: number;
  /** 功能：本段落的域字符标记，保存时写回 fldChar/instrText；时间：2026-10-30。 */
  markers?: FieldMarker[];
  /** 功能：累积本段落内出现的书签名（w:bookmarkStart），解析完挂到 DocParagraph.bookmarks；时间：2026-10-30。 */
  bookmarks?: string[];
  /** 功能：当前打开的批注范围 id（w:commentRangeStart/End 之间）；时间：2026-10-31。 */
  activeComments?: Set<string>;
  /** 功能：是否处于 w:del 删除修订内（run 标记为删除）；时间：2026-10-31。 */
  deletedRevision?: { author?: string; date?: string } | null;
  /** 功能：是否处于 w:ins 插入修订内；时间：2026-10-31。 */
  insertedRevision?: { author?: string; date?: string } | null;
}

/** 功能：文档级复杂域的栈项，记录指令文本与已登记的域标记；时间：2026-10-30。 */
interface FieldStackEntry {
  instr: string;
  phase: 'instr' | 'result';
  mergeEmitted?: boolean;
  /** 功能：begin 标记对象——指令文本在 separate/end 时才完整，届时回填；时间：2026-10-30。 */
  beginMarker?: FieldMarker;
  /** 功能：本域已登记的标记（begin/separate），命中自表达域时整体跳过写回；时间：2026-10-30。 */
  markers?: FieldMarker[];
}

/** 功能：判断域指令是否已由 fldSimple 自表达（合并域、页码域），这类域不再写回复杂域结构；时间：2026-10-30。 */
function isSelfExpressedField(instr: string): boolean {
  if (parseMergeField(instr) || parseMergeDirective(instr)) return true;
  return /\bPAGE\b/i.test(instr) || /\bNUMPAGES\b/i.test(instr) || /\bSECTIONPAGES\b/i.test(instr);
}

function pushControlFieldRun(runs: DocRun[], directive: MergeDirective, style: RunStyle): void {
  switch (directive.kind) {
    case 'next':
      runs.push({ text: '«NEXT»', mergeNext: true, style });
      break;
    case 'nextIf':
      runs.push({ text: '«NEXTIF»', mergeNextIf: directive.condition, style });
      break;
    case 'skipIf':
      runs.push({ text: '«SKIPIF»', mergeSkipIf: directive.condition, style });
      break;
    case 'record':
      runs.push({ text: '«MERGEREC»', mergeRecord: true, style });
      break;
    case 'seq':
      runs.push({ text: '«MERGESEQ»', mergeSeq: true, style });
      break;
    case 'mergeField':
      runs.push({
        text: `«${directive.name}»`,
        mergeField: directive.name,
        mergeFormat: directive.format,
        style,
      });
      break;
  }
}

async function parseRunElement(
  element: Element,
  base: RunStyle,
  ctx: ParseContext,
  state: ParseInlineState,
): Promise<DocRun[]> {
  const rPr = firstChild(element, 'w:rPr');
  const charStyleId = attrValue(firstChild(rPr, 'w:rStyle'), 'w:val');
  const charStyle = charStyleId ? ctx.styles.styles.get(charStyleId)?.run : undefined;
  const style = applyRunStyle(mergeRun(base, charStyle), rPr, ctx);
  const runs: DocRun[] = [];
  const topField = () => state.fieldStack.at(-1);

  for (const child of Array.from(element.childNodes)) {
    if (!(child instanceof Element)) continue;
    switch (localName(child)) {
      case 't': {
        const text = child.textContent ?? '';
        if (text) {
          const run: DocRun = { text, style };
          const field = topField();
          if (field && field.phase === 'result') {
            // 功能：跨多个文本 run 的字段只生成一个原子占位符；时间：2026-10-24；作者：lq。
            const name = parseMergeField(field.instr);
            if (name) {
              if (!field.mergeEmitted) {
                const directive = parseMergeDirective(field.instr);
                runs.push({
                  text: `«${name}»`,
                  mergeField: name,
                  ...(directive?.kind === 'mergeField' && directive.format ? { mergeFormat: directive.format } : {}),
                  style,
                });
              }
              field.mergeEmitted = true;
              break;
            }
            const directive = parseMergeDirective(field.instr);
            if (directive && !field.mergeEmitted) {
              field.mergeEmitted = true;
              pushControlFieldRun(runs, directive, style);
              break;
            }
            if (/\bPAGE\b/i.test(field.instr) && !/NUMPAGES/i.test(field.instr)) run.field = 'PAGE';
            else if (/NUMPAGES/i.test(field.instr)) run.field = 'NUMPAGES';
            else if (/SECTIONPAGES/i.test(field.instr)) run.field = 'SECTIONPAGES';
          }
          runs.push(run);
        }
        break;
      }
      case 'tab':
        runs.push({ text: '', style, tab: true });
        break;
      case 'delText': {
        const delText = child.textContent ?? '';
        if (delText) runs.push({ text: delText, style, revision: { type: 'delete', ...(state.deletedRevision?.author ? { author: state.deletedRevision.author } : {}), ...(state.deletedRevision?.date ? { date: state.deletedRevision.date } : {}) } });
        break;
      }
      case 'br': {
        const type = attrValue(child, 'w:type');
        // 功能：分页符强制断页、分栏符切到下一栏，其余按普通换行；时间：2026-10-30；作者：lq。
        runs.push({
          text: '',
          style,
          pageBreak: type === 'page',
          columnBreak: type === 'column',
          lineBreak: type !== 'page' && type !== 'column',
        });
        break;
      }
      case 'cr':
        runs.push({ text: '', style, lineBreak: true });
        break;
      case 'noBreakHyphen':
        runs.push({ text: '‑', style });
        break;
      case 'softHyphen':
        runs.push({ text: '\u00AD', style });
        break;
      case 'sym': {
        const charCode = attrValue(child, 'w:char');
        if (charCode) {
          const code = parseInt(charCode, 16);
          if (Number.isFinite(code)) runs.push({ text: String.fromCodePoint(code), style });
        }
        break;
      }
      case 'drawing': {
        runs.push(...(await parseDrawing(child, ctx)));
        break;
      }
      case 'pict': {
        runs.push(...(await parseVmlPicture(child, ctx)));
        break;
      }
      case 'AlternateContent': {
        // 功能：mc:AlternateContent——优先 mc:Choice 的现代绘图（drawing/pict），无产出时回退 mc:Fallback；时间：2026-10-31。
        const parseBranch = async (branch: Element | null): Promise<DocRun[]> => {
          const produced: DocRun[] = [];
          for (const sub of Array.from(branch?.childNodes ?? [])) {
            if (!(sub instanceof Element)) continue;
            const name = localName(sub);
            if (name === 'drawing') produced.push(...(await parseDrawing(sub, ctx)));
            else if (name === 'pict') produced.push(...(await parseVmlPicture(sub, ctx)));
          }
          return produced;
        };
        let produced = await parseBranch(firstDescendant(child, 'mc:Choice'));
        if (!produced.some((run) => run.text || run.image || run.textbox)) {
          produced = await parseBranch(firstDescendant(child, 'mc:Fallback'));
        }
        if (produced.some((run) => run.text || run.image || run.textbox)) runs.push(...produced);
        break;
      }
      case 'fldChar': {
        const type = attrValue(child, 'w:fldCharType');
        if (type === 'begin') {
          // 功能：登记复杂域的 begin 标记，保存时写回 fldChar+instrText，避免目录等域退化成死文本；时间：2026-10-30。
          const marker: FieldMarker = { role: 'begin', afterRun: state.runCount ?? 0 };
          state.fieldStack.push({ instr: '', phase: 'instr', beginMarker: marker, markers: [marker] });
          state.markers?.push(marker);
        } else if (type === 'separate' && topField()) {
          const field = topField()!;
          field.phase = 'result';
          if (field.beginMarker) field.beginMarker.instr = field.instr;
          const marker: FieldMarker = { role: 'separate', afterRun: state.runCount ?? 0 };
          field.markers?.push(marker);
          state.markers?.push(marker);
        } else if (type === 'end') {
          const field = topField();
          if (field) {
            const name = parseMergeField(field.instr);
            if (name && !field.mergeEmitted) {
              const directive = parseMergeDirective(field.instr);
              runs.push({
                text: `«${name}»`,
                mergeField: name,
                ...(directive?.kind === 'mergeField' && directive.format ? { mergeFormat: directive.format } : {}),
                style,
              });
              field.mergeEmitted = true;
            } else if (!field.mergeEmitted) {
              const directive = parseMergeDirective(field.instr);
              if (directive) {
                field.mergeEmitted = true;
                pushControlFieldRun(runs, directive, style);
              }
            }
            // 功能：合并域/页码域由 fldSimple 自表达，复杂域结构不再写回以免重复；时间：2026-10-30。
            if (isSelfExpressedField(field.instr)) {
              for (const marker of field.markers ?? []) marker.skip = true;
            } else if (field.beginMarker) {
              field.beginMarker.instr = field.instr;
            }
          }
          const marker: FieldMarker = { role: 'end', afterRun: state.runCount ?? 0 };
          if (field && isSelfExpressedField(field.instr)) marker.skip = true;
          state.markers?.push(marker);
          state.fieldStack.pop();
        }
        break;
      }
      case 'instrText': {
        const field = topField();
        if (field && field.phase === 'instr') field.instr += child.textContent ?? '';
        break;
      }
      case 'delText':
        break;
      default:
        break;
    }
  }
  // 功能：把当前打开的批注范围与修订标记附加到本 run 的所有产出；时间：2026-10-31。
  for (const run of runs) {
    if (state.activeComments?.size) run.commentIds = [...state.activeComments];
    if (state.insertedRevision) run.revision = { type: 'insert', ...(state.insertedRevision.author ? { author: state.insertedRevision.author } : {}), ...(state.insertedRevision.date ? { date: state.insertedRevision.date } : {}) };
    else if (state.deletedRevision) run.revision = { type: 'delete', ...(state.deletedRevision.author ? { author: state.deletedRevision.author } : {}), ...(state.deletedRevision.date ? { date: state.deletedRevision.date } : {}) };
  }

  return runs;
}

async function parseInlineChildren(
  parent: Element,
  base: RunStyle,
  ctx: ParseContext,
  state: ParseInlineState,
): Promise<DocRun[]> {
  const runs: DocRun[] = [];
  for (const child of Array.from(parent.childNodes)) {
    if (!(child instanceof Element)) continue;
    switch (localName(child)) {
      case 'r': {
        const produced = await parseRunElement(child, base, ctx, state);
        // 功能：累计本段已产出的 run 数，供域标记定位插入位置；时间：2026-10-30。
        state.runCount = (state.runCount ?? 0) + produced.length;
        runs.push(...produced);
        break;
      }
      case 'hyperlink': {
        // 功能：保留目录项等内部书签链接（w:anchor），供编辑区点击跳转到目标标题；时间：2026-10-30。
        const anchor = attrValue(child, 'w:anchor');
        const inner = await parseInlineChildren(child, base, ctx, state);
        if (anchor) for (const run of inner) run.anchor = anchor;
        runs.push(...inner);
        break;
      }
      case 'ins':
      case 'smartTag':
      case 'customXml':
      case 'sdt': {
        const container =
          localName(child) === 'sdt' ? firstChild(child, 'w:sdtContent') ?? child : child;
        if (localName(child) === 'ins') {
          ctx.hasTrackedChanges = true;
          // 功能：插入修订——产出 run 带 insert 标记（审阅窗格可接受/拒绝）；时间：2026-10-31。
          const prevInserted = state.insertedRevision;
          state.insertedRevision = { author: getAttr(child, 'w:author') ?? undefined, date: getAttr(child, 'w:date') ?? undefined };
          const inner = await parseInlineChildren(container, base, ctx, state);
          state.insertedRevision = prevInserted ?? null;
          runs.push(...inner);
          break;
        }
        if (localName(child) === 'sdt') ctx.hasContentControls = true;
        const inner = await parseInlineChildren(container, base, ctx, state);
        // 功能：内联内容控件保留控件语义（复选框/下拉/日期），空内容标记为占位待填；时间：2026-10-30。
        const control = localName(child) === 'sdt' ? parseSdtControl(child) : undefined;
        if (control) {
          const text = inner.map((run) => run.text ?? '').join('');
          if (!inner.length || !text.trim()) control.placeholder = true;
          if (inner.length) {
            inner[0] = { ...inner[0], control };
            runs.push(...inner);
          } else {
            runs.push({ text: '', style: base, control });
          }
          break;
        }
        runs.push(...inner);
        break;
      }
      case 'fldSimple': {
        const instr = attrValue(child, 'w:instr') ?? '';
        // 功能：兼容简式 MERGEFIELD，保持第一个结果字符的样式；时间：2026-10-24；作者：lq。
        const directive = parseMergeDirective(instr);
        if (directive?.kind === 'mergeField') {
          const result = await parseInlineChildren(child, base, ctx, state);
          runs.push({
            text: `«${directive.name}»`,
            mergeField: directive.name,
            mergeFormat: directive.format,
            style: result[0]?.style ?? base,
          });
          break;
        }
        if (directive) {
          pushControlFieldRun(runs, directive, base);
          break;
        }
        const before = runs.length;
        runs.push(...(await parseInlineChildren(child, base, ctx, state)));
        if (/\bPAGE\b/i.test(instr) && !/NUMPAGES/i.test(instr)) {
          for (let index = before; index < runs.length; index += 1) runs[index].field = 'PAGE';
        } else if (/NUMPAGES/i.test(instr)) {
          for (let index = before; index < runs.length; index += 1) runs[index].field = 'NUMPAGES';
        } else if (/SECTIONPAGES/i.test(instr)) {
          for (let index = before; index < runs.length; index += 1) runs[index].field = 'SECTIONPAGES';
        }
        break;
      }
      case 'bookmarkStart': {
        // 功能：记录书签名（w:name），供目录项 anchor 跳转定位到该段落；时间：2026-10-30。
        const name = attrValue(child, 'w:name');
        if (name) (state.bookmarks ??= []).push(name);
        break;
      }
      case 'bookmarkEnd':
      case 'proofErr':
        break;
      case 'commentRangeStart': {
        // 功能：打开批注范围，后续 run 携带批注 id；时间：2026-10-31。
        const cid = attrValue(child, 'w:id');
        if (cid) (state.activeComments ??= new Set()).add(cid);
        break;
      }
      case 'commentRangeEnd': {
        const cidEnd = attrValue(child, 'w:id');
        if (cidEnd) state.activeComments?.delete(cidEnd);
        break;
      }
      case 'commentReference':
        break;
      case 'del': {
        // 功能：删除修订——delText 保留为带 delete 标记的 run（审阅可拒绝以恢复）；时间：2026-10-31。
        ctx.hasTrackedChanges = true;
        const prevDeleted = state.deletedRevision;
        state.deletedRevision = { author: getAttr(child, 'w:author') ?? undefined, date: getAttr(child, 'w:date') ?? undefined };
        const inner = await parseInlineChildren(child, base, ctx, state);
        state.deletedRevision = prevDeleted ?? null;
        runs.push(...inner);
        break;
      }
      case 'br': {
        // 功能：分页符强制断页、分栏符切到下一栏，其余按普通换行；时间：2026-10-30；作者：lq。
        const brType = getAttr(child, 'w:type');
        runs.push({
          text: '',
          style: base,
          pageBreak: brType === 'page',
          columnBreak: brType === 'column',
          lineBreak: brType !== 'page' && brType !== 'column',
        });
        break;
      }
      default:
        break;
    }
  }
  return runs;
}


async function parseParagraph(
  element: Element,
  ctx: ParseContext,
  baseStyle?: { run: RunStyle; para: ParagraphStyle },
): Promise<{ paragraph: DocParagraph; pageBreakAfter: boolean }> {
  const pPr = firstChild(element, 'w:pPr');
  const styleId = attrValue(firstChild(pPr, 'w:pStyle'), 'w:val');
  const styleFromId = styleId ? ctx.styles.styles.get(styleId) : undefined;
  const styleBase = styleFromId ?? ctx.styles.defaultParagraph ?? baseStyle ?? {
    run: ctx.styles.docDefaults.run,
    para: ctx.styles.docDefaults.para,
  };
  let paragraphStyle = applyParagraphStyle(mergeParagraph(styleBase.para, baseStyle?.para), pPr, ctx);
  // 功能：文档自带 TOC 样式若缺少制表位，回退到内置默认（右对齐+点前导符）；时间：2026-10-30。
  if (styleId && /^TOC\s*\d+$/i.test(styleId) && !paragraphStyle.tabStops?.length) {
    const level = parseInt(styleId.replace(/[^\d]/g, ''), 10);
    if (level >= 1 && level <= 9) {
      paragraphStyle.tabStops = [{ posPx: 400, kind: 'right', leader: 'dot' }];
      paragraphStyle.indentLeftPx ??= level * 18;
    }
  }
  const numbering = parseProjectedParagraphNumbering(pPr);
  const numId = numbering.numId ?? styleFromId?.numId;
  const ilvl = numbering.ilvl ?? styleFromId?.ilvl ?? 0;
  let list: ListInfo | undefined;
  if (numId) {
    list = listInfoFor(ctx, numId, ilvl);
    const level = ctx.numbering.get(numId)?.find((item) => item.ilvl === ilvl);
    if (level) {
      if (paragraphStyle.indentLeftPx === undefined && level.indentLeftPx !== undefined) {
        paragraphStyle.indentLeftPx = level.indentLeftPx;
      }
      if (
        paragraphStyle.firstLineIndentPx === undefined &&
        level.firstLineIndentPx !== undefined
      ) {
        paragraphStyle.firstLineIndentPx = level.firstLineIndentPx;
      }
    }
  }

  // 功能：继承段落属性中的文字格式（Word 将段落级加粗、字号等写入 w:pPr/w:rPr）；时间：2026-10-27；作者：lq。
  // 旧逻辑只合并样式表和字符级 w:rPr，导致从原生 Word 打开的段落格式在编辑区丢失。
  const paragraphRun = applyRunStyle({}, firstChild(pPr, 'w:rPr'), ctx);
  const runBase = mergeRun(
    mergeRun(mergeRun(ctx.styles.docDefaults.run, styleBase.run), baseStyle?.run),
    paragraphRun,
  );
  const state: ParseInlineState = { fieldStack: ctx.fieldStack, runCount: 0, markers: [], bookmarks: [] };
  let runs = await parseInlineChildren(element, runBase, ctx, state);
  // 功能：只有段末的分页符才算"此段之后分页"；段中/段首的分页符保留为强制断点，由分页器处理；时间：2026-10-30；作者：lq。
  const lastRun = runs[runs.length - 1];
  const pageBreakAfter = lastRun?.pageBreak === true;
  if (pageBreakAfter) runs = runs.slice(0, -1);
  // 功能：段首分栏符提升为段落属性，多栏节渲染时据此切栏；时间：2026-10-30。
  const leadingColumnBreak = runs[0]?.columnBreak === true;
  if (leadingColumnBreak) {
    runs = runs.slice(1);
    paragraphStyle.columnBreakBefore = true;
  }
  if (runs.length === 0) runs = [{ text: '', style: runBase }];
  // 功能：逐字母模式合并——空格 run 带负 letterSpacing，合并到前一字母 run，去掉可见空格；
  // 时间：2026-10-30。
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i];
    const prev = runs[i - 1];
    if (
      run.text === ' ' &&
      run.style.letterSpacingPx !== undefined &&
      prev &&
      prev.text &&
      prev.text.length === 1 &&
      !prev.image &&
      !prev.field &&
      !prev.mergeField &&
      !prev.anchor &&
      !prev.href &&
      !prev.style.letterSpacingPx
    ) {
      prev.style.letterSpacingPx = run.style.letterSpacingPx;
      runs.splice(i, 1);
      i--;
    }
  }

  const paragraph: DocParagraph = {
    type: 'paragraph',
    id: uid('p'),
    runs,
    style: paragraphStyle,
    list,
    styleId,
    bookmarks: state.bookmarks?.length ? [...state.bookmarks] : undefined,
    // 功能：保留本段落内的域字符标记，保存时写回复杂域结构（目录等）；时间：2026-10-30。
    fieldMarkers: state.markers?.length ? state.markers : undefined,
  };
  // 功能：保留段尾分节以及各节独立的页眉页脚与编号；时间：2026-10-24；作者：lq。
  const section = firstChild(pPr, 'w:sectPr');
  if (section) {
    const page = parsePageSetup(section, ctx);
    // 功能：解析首页/奇偶页页眉页脚变体（仅当分节引用了对应类型）；时间：2026-10-31。
    const hasVariant = (kind: string, variant: string): boolean =>
      Array.from(section.childNodes).some(
        (node) => node instanceof Element && localName(node) === `${kind}Reference` && attrValue(node, 'w:type') === variant,
      );
    paragraph.sectionEnd = {
      page,
      header: await parseHeaderFooter(section, ctx, 'header', page.margins.header),
      footer: await parseHeaderFooter(section, ctx, 'footer', page.margins.footer),
      headerFirst: hasVariant('header', 'first')
        ? await parseHeaderFooter(section, ctx, 'header', page.margins.header, 'first')
        : undefined,
      headerEven: hasVariant('header', 'even')
        ? await parseHeaderFooter(section, ctx, 'header', page.margins.header, 'even')
        : undefined,
      footerFirst: hasVariant('footer', 'first')
        ? await parseHeaderFooter(section, ctx, 'footer', page.margins.footer, 'first')
        : undefined,
      footerEven: hasVariant('footer', 'even')
        ? await parseHeaderFooter(section, ctx, 'footer', page.margins.footer, 'even')
        : undefined,
    };
    }
  return { paragraph, pageBreakAfter };
}

async function parseTable(element: Element, ctx: ParseContext): Promise<DocTable> {
  const tblPr = firstChild(element, 'w:tblPr');
  const grid = firstChild(element, 'w:tblGrid');
  let columnWidths = Array.from(grid?.childNodes ?? [])
    .filter((node): node is Element => node instanceof Element && localName(node) === 'gridCol')
    .map((col) => twipsToPx(numericAttr(col, 'w:w', 0) ?? 0))
    .filter((width) => width > 0);

  const styleId = attrValue(firstChild(tblPr, 'w:tblStyle'), 'w:val');
  const borders = parseTableBorders(firstChild(tblPr, 'w:tblBorders'));
  const hasAnyBorder = Object.values(borders).some((border) => border && border.style !== 'none');
  if (!hasAnyBorder && styleId && /grid/i.test(styleId)) {
    const single: TableBorderSide = { style: 'single', widthPx: 0.75, color: '#000000' };
    borders.top = single;
    borders.bottom = single;
    borders.left = single;
    borders.right = single;
    borders.insideH = single;
    borders.insideV = single;
  }

  const alignmentValue = attrValue(firstChild(tblPr, 'w:jc'), 'w:val');
  const alignmentMap: Record<string, DocTable['alignment']> = {
    center: 'center',
    right: 'right',
    end: 'right',
    left: 'left',
    start: 'left',
  };
  const tableWidth = firstChild(tblPr, 'w:tblW');
  const widthType = attrValue(tableWidth, 'w:type');
  const widthValue = numericAttr(tableWidth, 'w:w');
  const widthPct =
    widthType === 'pct' && widthValue !== undefined ? widthValue / 50 : undefined;

  interface RawCell {
    cell: DocTableCell;
    vMerge?: 'restart' | 'continue';
    colSpan: number;
  }

  const rawRows: Array<{ cells: RawCell[]; heightPx?: number; isHeader: boolean }> = [];
  let colCount = columnWidths.length;

  for (const rowElement of Array.from(element.childNodes)) {
    if (!(rowElement instanceof Element) || localName(rowElement) !== 'tr') continue;
    const trPr = firstChild(rowElement, 'w:trPr');
    const height = numericAttr(firstChild(trPr, 'w:trHeight'), 'w:val');
    const isHeader = toggleValue(firstChild(trPr, 'w:tblHeader'));
    const rawCells: RawCell[] = [];
    for (const cellElement of Array.from(rowElement.childNodes)) {
      if (!(cellElement instanceof Element) || localName(cellElement) !== 'tc') continue;
      const tcPr = firstChild(cellElement, 'w:tcPr');
      const gridSpan = numericAttr(firstChild(tcPr, 'w:gridSpan'), 'w:val') ?? 1;
      const vMergeElement = firstChild(tcPr, 'w:vMerge');
      const vMergeValue = vMergeElement ? attrValue(vMergeElement, 'w:val') : undefined;
      const vMerge: RawCell['vMerge'] = vMergeElement
        ? vMergeValue === 'restart'
          ? 'restart'
          : 'continue'
        : undefined;
      const cellWidthTwips = numericAttr(firstChild(tcPr, 'w:tcW'), 'w:w');
      const cellWidth = cellWidthTwips ? twipsToPx(cellWidthTwips) : undefined;
      const backgroundRaw = attrValue(firstChild(tcPr, 'w:shd'), 'w:fill');
      const backgroundColor =
        backgroundRaw && /^[0-9a-fA-F]{6}$/.test(backgroundRaw)
          ? `#${backgroundRaw.toUpperCase()}`
          : undefined;
      const vAlignValue = attrValue(firstChild(tcPr, 'w:vAlign'), 'w:val');
      const verticalAlign =
        vAlignValue === 'center' ? 'middle' : vAlignValue === 'bottom' ? 'bottom' : 'top';
      const cellBorders = parseTableBorders(firstChild(tcPr, 'w:tcBorders'));
      const mergedBorders: TableBorders = { ...borders, ...cellBorders };
      const blocks =
        vMerge === 'continue' ? [] : await parseBlockContainer(cellElement, ctx);
      if (vMerge !== 'continue' && blocks.length === 0) {
        blocks.push({
          type: 'paragraph',
          id: uid('p'),
          runs: [{ text: '', style: {} }],
          style: {},
        });
      }
      rawCells.push({
        colSpan: gridSpan,
        vMerge,
        cell: {
          id: uid('cell'),
          blocks,
          colSpan: gridSpan,
          rowSpan: 1,
          widthPx: cellWidth,
          backgroundColor,
          borders: mergedBorders,
          verticalAlign,
        },
      });
      colCount = Math.max(
        colCount,
        rawCells.reduce((sum, cell) => sum + cell.colSpan, 0),
      );
    }
    rawRows.push({
      cells: rawCells,
      heightPx: height ? twipsToPx(height) : undefined,
      isHeader,
    });
  }

  if (columnWidths.length === 0) {
    const available = Math.max(320, ctx.pageWidthPx - 192);
    columnWidths = Array.from({ length: Math.max(1, colCount) }, () => available / Math.max(1, colCount));
  } else if (columnWidths.length < colCount) {
    // 功能：补齐缺少的 tblGrid 列宽，避免合并单元格后的后续列回退到固定 100px 而错位；时间：2026-10-27；作者：lq。
    const average = columnWidths.reduce((sum, width) => sum + width, 0) / columnWidths.length;
    columnWidths = [...columnWidths, ...Array.from({ length: colCount - columnWidths.length }, () => average || 80)];
  }

  const columnTops: Array<DocTableCell | undefined> = [];
  const rows: DocTableRow[] = [];
  for (const rawRow of rawRows) {
    const cells: DocTableCell[] = [];
    let column = 0;
    for (const raw of rawRow.cells) {
      const span = Math.max(1, raw.colSpan);
      if (raw.vMerge === 'continue') {
        const top = columnTops[column];
        if (top) {
          top.rowSpan += 1;
          for (let offset = 0; offset < span; offset += 1) columnTops[column + offset] = top;
        }
      } else {
        raw.cell.rowSpan = 1;
        cells.push(raw.cell);
        for (let offset = 0; offset < span; offset += 1) columnTops[column + offset] = raw.cell;
      }
      column += span;
    }
    // 功能：保留完全被纵向合并覆盖的行，防止保存重开后表格坍缩；时间：2026-10-19；作者：lq。
    {
      rows.push({
        id: uid('row'),
        cells,
        heightPx: rawRow.heightPx,
        isHeader: rawRow.isHeader,
      });
    }
  }

  return {
    type: 'table',
    id: uid('tbl'),
    rows,
    columnWidthsPx: columnWidths,
    widthPct,
    alignment: alignmentValue ? alignmentMap[alignmentValue] ?? 'left' : 'left',
    borders,
    styleId,
  };
}

async function parseBlockContainer(element: Element, ctx: ParseContext): Promise<DocBlock[]> {
  const blocks: DocBlock[] = [];
  for (const child of Array.from(element.childNodes)) {
    if (!(child instanceof Element)) continue;
    switch (localName(child)) {
      case 'p': {
        const { paragraph, pageBreakAfter } = await parseParagraph(child, ctx);
        // 功能：纯分页标记不额外占用一行，避免长表格重开后插入空白页；时间：2026-10-24；作者：lq。
        const breakOnly = pageBreakAfter && !paragraph.sectionEnd && paragraph.runs.every((run) => !run.text && !run.image && !run.tab && !run.lineBreak);
        if (!breakOnly) blocks.push(paragraph);
        if (pageBreakAfter) blocks.push({ type: 'pageBreak', id: uid('break') });
        break;
      }
      case 'tbl':
        blocks.push(await parseTable(child, ctx));
        break;
      case 'sdt': {
        ctx.hasContentControls = true;
        const content = firstChild(child, 'w:sdtContent');
        if (!content) break;
        const inner = await parseBlockContainer(content, ctx);
        // 功能：块级内容控件只包一段时挂到该段上，渲染成可交互控件；时间：2026-10-30。
        const control = parseSdtControl(child);
        if (control && inner.length === 1 && inner[0].type === 'paragraph') {
          // 功能：块级控件挂到该段首个 run 上，渲染与写回和行内控件走同一条路径；时间：2026-10-30。
          const paragraph = inner[0];
          const runs = paragraph.runs.length ? paragraph.runs : [{ text: '', style: {} }];
          runs[0] = { ...runs[0], control };
          if (!(runs[0].text ?? '').trim()) control.placeholder = true;
          inner[0] = { ...paragraph, runs };
        }
        blocks.push(...inner);
        break;
      }
      case 'customXml': {
        blocks.push(...(await parseBlockContainer(child, ctx)));
        break;
      }
      case 'altChunk':
      case 'sectPr':
        break;
      default:
        break;
    }
  }
  return blocks;
}

function parsePageSetup(sectPr: Element | null, ctx: ParseContext): PageSetup {
  const pageSize = firstChild(sectPr, 'w:pgSz');
  const widthTwips = numericAttr(pageSize, 'w:w', 11_906) ?? 11_906;
  const heightTwips = numericAttr(pageSize, 'w:h', 16_838) ?? 16_838;
  const orientation = attrValue(pageSize, 'w:orient') === 'landscape' ? 'landscape' : 'portrait';
  const marginsElement = firstChild(sectPr, 'w:pgMar');
  const margins = {
    top: twipsToPx(numericAttr(marginsElement, 'w:top', 1440) ?? 1440),
    right: twipsToPx(numericAttr(marginsElement, 'w:right', 1440) ?? 1440),
    bottom: twipsToPx(numericAttr(marginsElement, 'w:bottom', 1440) ?? 1440),
    left: twipsToPx(numericAttr(marginsElement, 'w:left', 1440) ?? 1440),
    header: twipsToPx(numericAttr(marginsElement, 'w:header', 720) ?? 720),
    footer: twipsToPx(numericAttr(marginsElement, 'w:footer', 720) ?? 720),
    gutter: twipsToPx(numericAttr(marginsElement, 'w:gutter', 0) ?? 0),
  };
  const sizeMap: Array<{ name: string; w: number; h: number }> = [
    { name: 'A4', w: 11_906, h: 16_838 },
    { name: 'A3', w: 16_838, h: 23_811 },
    { name: 'A5', w: 8_391, h: 11_906 },
    { name: 'Letter', w: 12_240, h: 15_840 },
    { name: 'Legal', w: 12_240, h: 20_160 },
    { name: 'B5', w: 9_976, h: 14_170 },
  ];
  // 功能：正确区分横向与纵向纸张尺寸，避免所有自定义尺寸都被误标为同一类；时间：2026-10-27；作者：lq。
  const swapped = widthTwips > heightTwips;
  const matched = sizeMap.find(
    (size) =>
      (Math.abs(size.w - widthTwips) < 80 && Math.abs(size.h - heightTwips) < 80) ||
      (Math.abs(size.w - heightTwips) < 80 && Math.abs(size.h - widthTwips) < 80),
  );
  const columns = numericAttr(firstChild(sectPr, 'w:cols'), 'w:num');
  // 功能：解析节分隔类型，continuous 表示同页继续，nextPage 强制新页；时间：2026-10-30。
  const sectType = attrValue(firstChild(sectPr, 'w:type'), 'w:val');
  // 功能：解析文档网格行距 w:docGrid w:linePitch（twips），段落行高按此网格对齐；时间：2026-10-30。
  const docGrid = firstChild(sectPr, 'w:docGrid');
  const linePitchTwips = docGrid ? numericAttr(docGrid, 'w:linePitch') : undefined;
  // 功能：首页页眉页脚不同（w:titlePg）；时间：2026-10-31。
  const titlePage = firstChild(sectPr, 'w:titlePg') ? true : undefined;
  // 功能：解析每栏宽度与间距，支持不等宽多栏版式；时间：2026-10-30。
  const colsElement = firstChild(sectPr, 'w:cols');
  const columnEqualWidth = colsElement ? attrValue(colsElement, 'w:equalWidth') === 'true' : undefined;
  const colWidthsPx: number[] = [];
  const colSpacePx: number[] = [];
  if (colsElement) {
    for (const col of Array.from(colsElement.childNodes)) {
      if (!(col instanceof Element)) continue;
      if (localName(col) === 'col') {
        const w = numericAttr(col, 'w:w');
        const s = numericAttr(col, 'w:space');
        if (w !== undefined) colWidthsPx.push(twipsToPx(w));
        if (s !== undefined) colSpacePx.push(twipsToPx(s));
      }
    }
  }
  return {
    widthPx: twipsToPx(widthTwips),
    heightPx: twipsToPx(heightTwips),
    orientation,
    sizeName: matched?.name ?? (swapped ? '自定义' : '自定义'),
    margins,
    columns: columns && columns > 1 ? columns : undefined,
    columnWidthsPx: colWidthsPx.length ? colWidthsPx : undefined,
    columnSpacePx: colSpacePx.length ? colSpacePx : undefined,
    columnEqualWidth,
    type: sectType === 'continuous' ? 'continuous' : 'nextPage',
    linePitchPx: linePitchTwips !== undefined ? twipsToPx(linePitchTwips) : undefined,
    titlePage,
    // 功能：恢复 DOCX 起始页码；时间：2026-10-24；作者：lq。
    pageNumberStart: numericAttr(firstChild(sectPr, 'w:pgNumType'), 'w:start'),
  };
}

/**
 * 功能：FNV-1a 字符串哈希——页眉页脚内容去重的缓存键；时间：2026-10-31。
 */
function hashXml(xml: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < xml.length; index += 1) {
    hash ^= xml.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return `${(hash >>> 0).toString(36)}-${xml.length}`;
}

/** 功能：深克隆块数组并重新生成所有 id（段落/表格/行/单元格/图片），保证模型内 id 唯一；时间：2026-10-31。 */
function cloneBlocksWithFreshIds(blocks: DocBlock[]): DocBlock[] {
  return blocks.map((block) => cloneBlockWithFreshIds(block));
}

function cloneBlockWithFreshIds(block: DocBlock): DocBlock {
  if (block.type === 'paragraph') {
    return {
      ...block,
      id: uid('p'),
      runs: block.runs.map((run) => (run.image ? { ...run, image: { ...run.image, id: uid('img') } } : run)),
      bookmarks: block.bookmarks ? [...block.bookmarks] : undefined,
      fieldMarkers: block.fieldMarkers ? block.fieldMarkers.map((marker) => ({ ...marker })) : undefined,
    };
  }
  if (block.type === 'table') {
    return {
      ...block,
      id: uid('tbl'),
      rows: block.rows.map((row) => ({
        ...row,
        id: uid('row'),
        cells: row.cells.map((cell) => ({ ...cell, id: uid('cell'), blocks: cloneBlocksWithFreshIds(cell.blocks) })),
      })),
    };
  }
  return { ...block, id: uid('break') };
}

/**
 * 功能：解析 word/comments.xml 批注列表——id/作者/日期/正文聚合；时间：2026-10-31。
 */
async function parseComments(pkg: OoxmlPackage, rels: Map<string, Relationship>): Promise<DocComment[] | undefined> {
  const rel = rels.get('comments-part') ?? [...rels.values()].find((item) => item.type.endsWith('/comments'));
  if (!rel?.resolved) return undefined;
  try {
    const xml = await pkg.text(rel.resolved);
    if (!xml.trim()) return undefined;
    const doc = parseXmlDom(xml);
    const comments: DocComment[] = [];
    for (const node of Array.from(doc.getElementsByTagName('w:comment'))) {
      const id = node.getAttribute('w:id');
      if (id === null || id === undefined || id === '') continue;
      const blocks = await parseBlockContainer(node, {
        pkg,
        rels,
        styles: { docDefaults: { run: {}, para: {} }, styles: new Map(), names: new Map(), types: new Map() },
        numbering: new Map(),
        counters: new Map(),
        themeColors: {},
        themeFonts: { major: '', minor: '' },
        defaultFont: '',
        defaultSizePt: 10.5,
        imageUrls: new Map(),
        usedImages: new Map(),
        pageWidthPx: 794,
        blockCounter: 0,
        hasTrackedChanges: false,
        hasContentControls: false,
        fieldStack: [],
      });
      const text = blocks
        .map((block) => (block.type === 'paragraph' ? block.runs.map((run) => run.text ?? '').join('') : ''))
        .join('\n')
        .trim();
      comments.push({
        id,
        author: node.getAttribute('w:author') ?? '',
        ...(node.getAttribute('w:date') ? { date: node.getAttribute('w:date') as string } : {}),
        ...(node.getAttribute('w:resolved') === '1' ? { resolved: true } : {}),
        text,
        ...(blocks.length ? { blocks } : {}),
      });
    }
    return comments.length ? comments : undefined;
  } catch {
    return undefined;
  }
}

async function parseHeaderFooter(
  sectPr: Element | null,
  ctx: ParseContext,
  kind: 'header' | 'footer',
  distancePx: number,
  /** 功能：引用变体——default/first/even（首页与奇偶页页眉页脚）；时间：2026-10-31。 */
  variant: 'default' | 'first' | 'even' = 'default',
): Promise<HeaderFooter | undefined> {
  if (!sectPr) return undefined;
  const references = Array.from(sectPr.childNodes).filter(
    (node): node is Element =>
      node instanceof Element && localName(node) === `${kind}Reference`,
  );
  if (references.length === 0) return undefined;
  const preferred =
    references.find((reference) => attrValue(reference, 'w:type') === variant) ??
    (variant === 'default' ? references[0] : undefined);
  if (!preferred) return undefined;
  const relId = preferred.getAttribute('r:id') ?? getAttrLocal(preferred, 'id');
  if (!relId) return undefined;
  const rel = ctx.rels.get(relId);
  if (!rel?.resolved) return undefined;
  const xml = await ctx.pkg.text(rel.resolved);
  // 功能：相同内容的页眉页脚只解析一次，使用时克隆并重新分配 id；时间：2026-10-31。
  const cacheKey = `${kind}:${hashXml(xml)}`;
  const cached = ctx.headerFooterCache?.get(cacheKey);
  if (cached) {
    return { ...cached, blocks: cloneBlocksWithFreshIds(cached.blocks), distancePx, sourcePart: rel.resolved };
  }
  const doc = parseXmlDom(xml);
  const root = firstDescendant(doc, kind === 'header' ? 'w:hdr' : 'w:ftr') ?? doc.documentElement;
  // 功能：页眉页脚图片关系按自身部件解析；时间：2026-10-24；作者：lq。
  const blocks = await parseBlockContainer(root, { ...ctx, rels: await relationshipMap(ctx.pkg, rel.resolved) });
  const pageNumbers = blocks.some((block) =>
    block.type === 'paragraph' ? block.runs.some((run) => run.field === 'PAGE') : false,
  );
    ctx.headerFooterCache?.set(cacheKey, { blocks, distancePx, pageNumbers, sourcePart: rel.resolved });
  return { blocks, distancePx, pageNumbers, sourcePart: rel.resolved };
}

export interface ParseDocxOptions {
  fileName?: string;
}

export async function parseDocxPackage(
  pkg: OoxmlPackage,
  options: ParseDocxOptions = {},
): Promise<DocModel> {
  const documentXml = await pkg.text('word/document.xml');
  const stylesXml = pkg.has('word/styles.xml') ? await pkg.text('word/styles.xml') : '';
  const numberingXml = pkg.has('word/numbering.xml') ? await pkg.text('word/numbering.xml') : '';
  const themeColors = await readThemeColors(pkg, 'word/theme/theme1.xml');
  const themeFonts = await readThemeFonts(pkg, 'word/theme/theme1.xml');
  const rels = await relationshipMap(pkg, 'word/document.xml');

  const provisional: ParseContext = {
    pkg,
    rels,
    styles: {
      docDefaults: { run: {}, para: {} },
      styles: new Map(),
      names: new Map(),
      types: new Map(),
    },
    numbering: parseNumbering(numberingXml),
    counters: new Map(),
    themeColors,
    themeFonts,
    defaultFont: themeFonts.minor || '宋体',
    defaultSizePt: 10.5,
    imageUrls: new Map(),
    usedImages: new Map(),
    pageWidthPx: 793.7,
    blockCounter: 0,
    hasTrackedChanges: false,
    hasContentControls: false,
    fieldStack: [],
    headerFooterCache: new Map(),
  };

  provisional.styles = parseStyleSheet(stylesXml, provisional);
  const defaultRun = provisional.styles.docDefaults.run;
  const defaultFont =
    defaultRun.fontFamily ??
    themeFonts.minor ??
    themeFonts.major ??
    provisional.defaultFont;
  provisional.defaultFont = defaultFont;
  provisional.defaultSizePt = defaultRun.fontSizePt ?? 10.5;

  const doc = parseXmlDom(documentXml);
  const body = firstDescendant(doc, 'w:body') ?? doc.documentElement;
  const sectionProperties =
    Array.from(body.childNodes).find(
      (node): node is Element => node instanceof Element && localName(node) === 'sectPr',
    ) ??
    firstDescendant(doc, 'w:sectPr');
  // 功能：兼容分节符位于正文段落中的页眉页脚引用，避免只读取末节导致首节页脚丢失；时间：2026-10-27；作者：lq。
  const sectionForPart = (kind: 'header' | 'footer') => {
    if (sectionProperties && firstChild(sectionProperties, `w:${kind}Reference`)) return sectionProperties;
    const candidates = Array.from(body.getElementsByTagName('w:sectPr'));
    return candidates.find((section) => firstChild(section, `w:${kind}Reference`)) ?? sectionProperties;
  };
  const page = parsePageSetup(sectionProperties, provisional);
  provisional.pageWidthPx = page.widthPx;

  const blocks = await parseBlockContainer(body, provisional);

  const header = await parseHeaderFooter(
    sectionForPart('header'),
    provisional,
    'header',
    page.margins.header,
  );
  const footer = await parseHeaderFooter(
    sectionForPart('footer'),
    provisional,
    'footer',
    page.margins.footer,
  );
  // 功能：批注列表（word/comments.xml）；时间：2026-10-31。
  const commentsPart = [...provisional.rels.entries()].find(([, item]) => item.type.endsWith('/comments'));
  const comments = commentsPart ? await parseComments(pkg, provisional.rels) : undefined;

  // 功能：文档级 sectPr 的首页/奇偶页页眉页脚变体（正文段无 sectPr 时唯一入口）；时间：2026-10-31。
  const hasDocVariant = (kind: 'header' | 'footer', variant: string): boolean => {
    const sect = sectionForPart(kind);
    return !!sect && Array.from(sect.childNodes).some(
      (node) => node instanceof Element && localName(node) === `${kind}Reference` && attrValue(node, 'w:type') === variant,
    );
  };
  const headerFirst = hasDocVariant('header', 'first')
    ? await parseHeaderFooter(sectionForPart('header'), provisional, 'header', page.margins.header, 'first')
    : undefined;
  const headerEven = hasDocVariant('header', 'even')
    ? await parseHeaderFooter(sectionForPart('header'), provisional, 'header', page.margins.header, 'even')
    : undefined;
  const footerFirst = hasDocVariant('footer', 'first')
    ? await parseHeaderFooter(sectionForPart('footer'), provisional, 'footer', page.margins.footer, 'first')
    : undefined;
  const footerEven = hasDocVariant('footer', 'even')
    ? await parseHeaderFooter(sectionForPart('footer'), provisional, 'footer', page.margins.footer, 'even')
    : undefined;

  let words = 0;
  let characters = 0;
  const visitParagraph = (paragraph: DocParagraph) => {
    const text = paragraph.runs.map((run) => run.text ?? '').join('');
    words += countWords(text);
    characters += text.replace(/\s/g, '').length;
  };
  const visitBlocks = (items: DocBlock[]) => {
    for (const block of items) {
      if (block.type === 'paragraph') visitParagraph(block);
      else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) visitBlocks(cell.blocks);
      }
    }
  };
  visitBlocks(blocks);
  if (header) visitBlocks(header.blocks);
  if (footer) visitBlocks(footer.blocks);

  const meta = await parseDocMeta(pkg);
  // 功能：把导入文档的段落样式带进模型，样式库可直接套用与识别；时间：2026-10-29；作者：lq。
  const documentStyles = collectDocumentStyles(blocks, header, footer, provisional.styles);

  // 功能：读回模板里保存的邮件合并面板状态，供打开时恢复面板；同时读取默认制表位间距；时间：2026-10-29；作者：lq。
  let defaultTabStopPx: number | undefined;
  let evenAndOddHeaders: boolean | undefined;
  if (pkg.has('word/settings.xml')) {
    try {
      const settingsDocument = parseXmlDom(await pkg.text('word/settings.xml'));
      const docVars = settingsDocument.getElementsByTagName('w:docVar');
      for (const docVar of Array.from(docVars)) {
        if (docVar.getAttribute('w:name') === 'mailMergePanelState') {
          const value = docVar.getAttribute('w:val');
          if (value) (meta as Record<string, unknown>).mailMergeState = value;
          break;
        }
      }
      // 功能：默认制表位（w:defaultTabStop，twips），无显式制表位时 tab 按它的整数倍跳转；时间：2026-10-30。
      const tabStop = settingsDocument.getElementsByTagName('w:defaultTabStop')[0];
      const tabStopTwips = tabStop ? numericAttr(tabStop, 'w:val') : undefined;
      if (tabStopTwips && tabStopTwips > 0) defaultTabStopPx = twipsToPx(tabStopTwips);
      // 功能：奇偶页页眉页脚不同（w:evenAndOddHeaders）；时间：2026-10-31。
      if (settingsDocument.getElementsByTagName('w:evenAndOddHeaders').length > 0) evenAndOddHeaders = true;
    } catch {
      // 忽略读取失败，不影响正文解析
    }
  }

  return {
    version: 1,
    page,
    blocks,
    header,
    footer,
    defaultFont,
    defaultFontSizePt: provisional.defaultSizePt,
    themeColors,
    themeFonts,
    images: Object.fromEntries(provisional.usedImages),
    meta,
    hasTrackedChanges: provisional.hasTrackedChanges,
    hasComments: pkg.has('word/comments.xml'),
    hasContentControls: provisional.hasContentControls,
    stats: { words, characters },
    ...(documentStyles.length > 0 ? { styles: documentStyles } : {}),
    ...(defaultTabStopPx !== undefined ? { defaultTabStopPx } : {}),
    ...(evenAndOddHeaders ? { evenAndOddHeaders } : {}),
    ...(headerFirst ? { headerFirst } : {}),
    ...(headerEven ? { headerEven } : {}),
    ...(footerFirst ? { footerFirst } : {}),
    ...(footerEven ? { footerEven } : {}),
    ...(comments ? { comments } : {}),
  };
}

async function parseDocMeta(pkg: OoxmlPackage): Promise<DocModel['meta']> {
  if (!pkg.has('docProps/core.xml')) return {};
  try {
    const doc = parseXmlDom(await pkg.text('docProps/core.xml'));
    const val = (names: string[]): string | undefined => {
      for (const name of names) {
        const element =
          doc.getElementsByTagName(name)[0] ?? doc.getElementsByTagName(`dc:${name}`)[0];
        if (element?.textContent) return element.textContent;
      }
      return undefined;
    };
    return {
      title: val(['title', 'dc:title']),
      author: val(['creator', 'dc:creator']),
      subject: val(['subject', 'dc:subject']),
      created: val(['created', 'dcterms:created']),
      modified: val(['modified', 'dcterms:modified']),
      lastModifiedBy: val(['lastModifiedBy', 'cp:lastModifiedBy']),
    };
  } catch {
    return {};
  }
}

export async function parseDocx(bytes: Uint8Array, options: ParseDocxOptions = {}): Promise<DocModel> {
  const pkg = await loadOoxmlPackage(bytes);
  try {
    return await parseDocxPackage(pkg, options);
  } finally {
    pkg.dispose();
  }
}
