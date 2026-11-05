import {
  type DocBlock,
  type DocImage,
  type DocModel,
  type DocParagraph,
  type DocRun,
  type DocStyleDefinition,
  type DocTable,
  type DocTableCell,
  type DocTableRow,
  type DocGrid,
  type HeaderFooter,
  type ImageAnchor,
  type ListInfo,
  type PageSetup,
  type ParagraphStyle,
  type RunStyle,
  type TabAlignment,
  type TabLeader,
  type TabStop,
  type TableBorderSide,
  type TableBorders,
  type TableCellMargin,
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
  readThemeColors,
  readThemeFonts,
  resolveThemeSlot,
  resolveWordShading,
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
// 功能：按字符数换算缩进需要字号，与段落对话框保持同一套汉字宽度；时间：2026-11-03；作者：lq。
import { charWidthPx } from './typography';
// 功能：把原始正文与可编辑块关联，以支持保留未修改 XML 的局部保存；时间：2026-11-03；作者：lq。
import { captureDocumentSource } from './source-preservation';

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
  /** 功能：表格样式的边框和默认单元格边距，沿 basedOn 链继承；时间：2026-11-03；作者：lq。 */
  table?: { borders: TableBorders; margin?: TableCellMargin };
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
  /** 功能：保留表格属性节点供继承解析，不依赖 styleId 的英文名称；时间：2026-11-03；作者：lq。 */
  tablePr?: Element | null;
}

interface StyleSheet {
  docDefaults: { run: RunStyle; para: ParagraphStyle };
  styles: Map<string, ResolvedStyle>;
  defaultParagraph?: ResolvedStyle;
  /** 功能：没有显式 tblStyle 时使用文档默认表格样式；时间：2026-11-03；作者：lq。 */
  defaultTable?: ResolvedStyle;
  /** 功能：styleId → 显示名 / 类型，用于把文档里实际用到的样式还原到样式库；时间：2026-10-29；作者：lq。 */
  names: Map<string, string>;
  types: Map<string, string>;
}

interface ParseContext {
  /** 功能：文件字体表中的主名到备用名映射，随格式继承到文本 run；时间：2026-11-05；作者：lq。 */
  fontAlternates: Map<string, string>;
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
  /** 功能：版心宽度（页面宽减左右页边距），用于表格自适应收敛；时间：2026-11-03；作者：lq。 */
  contentWidthPx?: number;
  blockCounter: number;
  hasTrackedChanges: boolean;
  hasContentControls: boolean;
  /** 功能：目录域状态；非空表示当前段落仍处在 TOC 域的缓存结果里，应整段丢弃；时间：2026-10-30；作者：lq。 */
  toc?: { levels: number; depth: number };
  /** 功能：按正文直接子元素记录对应模型块，未知元素保留空组；时间：2026-11-03；作者：lq。 */
  sourceGroups?: string[][];
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
  const font =
    attrValue(fonts, 'w:eastAsia') ??
    attrValue(fonts, 'w:ascii') ??
    attrValue(fonts, 'w:hAnsi') ??
    attrValue(fonts, 'w:cs');
  // 功能：只在字体族发生明确覆盖时重设备用名，避免将旧字体的备用名继承给新字体；时间：2026-11-05；作者：lq。
  if (font) {
    style.fontFamily = font;
    style.fontFallback = ctx.fontAlternates.get(font);
    style.fontFallbackFor = style.fontFallback ? font : undefined;
  }

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
    style.underline = value ? !['none', '0', 'false'].includes(value.toLowerCase()) : true;
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
  // 功能：文字底纹同样支持主题色（w:themeFill），不再只认 w:fill；时间：2026-11-01；作者：lq。
  const shadingFill = resolveWordShading(firstChild(rPr, 'w:shd'), ctx.themeColors);
  if (shadingFill) style.highlight = shadingFill;

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

function applyParagraphStyle(base: ParagraphStyle, pPr: Element | null, ctx: ParseContext, fontSizePt?: number): ParagraphStyle {
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
    // 功能：中文文档常用「缩进 N 字符」（w:firstLineChars，单位 1/100 字符），只有它没有 twips 时旧实现会整段丢失首行缩进；时间：2026-11-03；作者：lq。
    const leftChars = numericAttr(indent, 'w:leftChars');
    const firstLineChars = numericAttr(indent, 'w:firstLineChars');
    const hangingChars = numericAttr(indent, 'w:hangingChars');
    const charPx = charWidthPx(fontSizePt && fontSizePt > 0 ? fontSizePt : 10.5);
    if (left !== undefined) style.indentLeftPx = twipsToPx(left);
    else if (leftChars !== undefined) style.indentLeftPx = (leftChars / 100) * charPx;
    if (right !== undefined) style.indentRightPx = twipsToPx(right);
    // 功能：字符单位优先于 twips，与 ECMA-376 中 *Chars 覆盖同名字段的规则一致；时间：2026-11-03；作者：lq。
    if (firstLineChars !== undefined) style.firstLineIndentPx = (firstLineChars / 100) * charPx;
    else if (hangingChars !== undefined) style.firstLineIndentPx = -(hangingChars / 100) * charPx;
    else if (firstLine !== undefined) style.firstLineIndentPx = twipsToPx(firstLine);
    else if (hanging !== undefined) style.firstLineIndentPx = -twipsToPx(hanging);
  }

  const spacing = firstChild(pPr, 'w:spacing');
  if (spacing) {
    const before = numericAttr(spacing, 'w:before');
    const after = numericAttr(spacing, 'w:after');
    const line = numericAttr(spacing, 'w:line');
    const rule = attrValue(spacing, 'w:lineRule') ?? 'auto';
    if (before !== undefined) style.spaceBeforePx = twipsToPx(before);
    if (after !== undefined) style.spaceAfterPx = twipsToPx(after);
    // 功能：显式行距覆盖继承的其他行距模式，并保留 atLeast=0 的 Word 语义；时间：2026-11-03；作者：lq。
    if (line !== undefined) {
      delete style.lineHeight;
      delete style.lineHeightAtLeastPx;
      delete style.lineHeightExactPx;
      if (rule === 'auto') style.lineHeight = line / 240;
      else if (rule === 'atLeast') style.lineHeightAtLeastPx = twipsToPx(line);
      else style.lineHeightExactPx = twipsToPx(line);
    }
  }

  const keepNext = firstChild(pPr, 'w:keepNext');
  if (keepNext) style.keepNext = toggleValue(keepNext);
  // 功能：解析整段不拆页与孤行控制，分页时据此决定段落能否跨页切开；时间：2026-11-03；作者：lq。
  const keepLines = firstChild(pPr, 'w:keepLines');
  if (keepLines) style.keepLines = toggleValue(keepLines);
  const widowControl = firstChild(pPr, 'w:widowControl');
  if (widowControl) style.widowControl = toggleValue(widowControl);
  const pageBreakBefore = firstChild(pPr, 'w:pageBreakBefore');
  if (pageBreakBefore) style.pageBreakBefore = toggleValue(pageBreakBefore);
  // 功能：snapToGrid 控制段落是否吸附文档网格；显式 false 时不使用 linePitch 定行高；时间：2026-11-03；作者：lq。
  const snapToGridNode = firstChild(pPr, 'w:snapToGrid');
  if (snapToGridNode) {
    style.snapToGrid = toggleValue(snapToGridNode);
  }
  const outline = numericAttr(firstChild(pPr, 'w:outlineLvl'), 'w:val');
  if (outline !== undefined) style.outlineLevel = outline;

  // 功能：段落底纹支持主题色填充；时间：2026-11-01；作者：lq。
  const paragraphShading = resolveWordShading(firstChild(pPr, 'w:shd'), ctx.themeColors);
  if (paragraphShading) style.backgroundColor = paragraphShading;
  // 功能：解析段落边框，恢复 Word 中常见的横线、签名线和分隔线；时间：2026-10-27；作者：lq。
  const paragraphBorders = firstChild(pPr, 'w:pBdr');
  if (paragraphBorders) style.borders = parseTableBorders(paragraphBorders, ctx.themeColors);
  // 功能：解析制表位，标尺与正文 Tab 定位共用；时间：2026-10-30；作者：lq。
  const tabs = firstChild(pPr, 'w:tabs');
  if (tabs) {
    const stops = parseTabStops(tabs);
    if (stops.length > 0) style.tabStops = stops;
  }
  return style;
}

const TAB_ALIGNMENT_MAP: Record<string, TabAlignment> = {
  left: 'left',
  start: 'left',
  center: 'center',
  right: 'right',
  end: 'right',
  decimal: 'decimal',
  num: 'left',
};

const TAB_LEADER_MAP: Record<string, TabLeader> = {
  none: 'none',
  dot: 'dot',
  hyphen: 'hyphen',
  underscore: 'underscore',
  heavy: 'heavy',
  middleDot: 'middleDot',
};

/** 功能：读取 w:tabs 下的制表位，跳过 clear / bar 这两类纯排版标记；时间：2026-10-30；作者：lq。 */
function parseTabStops(tabs: Element): TabStop[] {
  const stops: TabStop[] = [];
  for (const child of Array.from(tabs.childNodes)) {
    if (!(child instanceof Element) || localName(child) !== 'tab') continue;
    const value = (attrValue(child, 'w:val') ?? 'left').toLowerCase();
    if (value === 'clear' || value === 'bar') continue;
    const pos = numericAttr(child, 'w:pos');
    if (pos === undefined) continue;
    const leader = attrValue(child, 'w:leader');
    stops.push({
      posPx: Math.max(0, twipsToPx(pos)),
      alignment: TAB_ALIGNMENT_MAP[value] ?? 'left',
      ...(leader && TAB_LEADER_MAP[leader] ? { leader: TAB_LEADER_MAP[leader] } : {}),
    });
  }
  return stops.sort((a, b) => a.posPx - b.posPx);
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
  let defaultTableId: string | undefined;

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
      rawStyles.set(id, { id, type, name: styleName, basedOn, isDefault, run, para, ...num, tablePr: firstChild(styleElement, 'w:tblPr') });
      if (type === 'paragraph' && isDefault) defaultParagraphId = id;
      if (type === 'table' && isDefault) defaultTableId = id;
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
      table: raw.type === 'table' ? {
        borders: { ...base.table?.borders, ...parseTableBorders(firstChild(raw.tablePr, 'w:tblBorders'), ctx.themeColors) },
        margin: parseTableCellMargin(firstChild(raw.tablePr, 'w:tblCellMar'), base.table?.margin) ?? base.table?.margin,
      } : undefined,
    };
    resolved.set(id, result);
    resolving.delete(id);
    return result;
  };

  rawStyles.forEach((_style, id) => resolve(id));
  const defaultParagraph = defaultParagraphId
    ? resolved.get(defaultParagraphId)
    : resolved.get('Normal');
  return { docDefaults, styles: resolved, defaultParagraph, defaultTable: defaultTableId ? resolved.get(defaultTableId) : undefined, names, types };
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

function parseTableBorder(
  element: Element | null,
  theme: Record<string, string> = {},
): TableBorderSide | undefined {
  if (!element) return undefined;
  const value = attrValue(element, 'w:val') ?? 'single';
  if (['nil', 'none'].includes(value)) {
    return { style: 'none', widthPx: 0, color: 'transparent' };
  }
  const size = numericAttr(element, 'w:sz', 4) ?? 4;
  const colorRaw = attrValue(element, 'w:color');
  // 功能：框线颜色支持主题色（w:themeColor），主题配色表格的边框不再一律变黑；时间：2026-11-01；作者：lq。
  const color =
    colorRaw && colorRaw.toLowerCase() !== 'auto' && /^[0-9a-fA-F]{6}$/.test(colorRaw)
      ? `#${colorRaw.toUpperCase()}`
      : resolveThemeSlot(
          attrValue(element, 'w:themeColor'),
          theme,
          attrValue(element, 'w:themeTint'),
          attrValue(element, 'w:themeShade'),
        ) ?? '#000000';
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

function parseTableBorders(
  element: Element | null,
  theme: Record<string, string> = {},
): TableBorders {
  if (!element) return {};
  // 功能：缺省边不生成 undefined 覆盖值；显式 nil 则保留 none，确保逐边继承语义；时间：2026-11-03；作者：lq。
  const borders: TableBorders = {};
  for (const side of ['top', 'right', 'bottom', 'left', 'insideH', 'insideV'] as const) {
    const value = parseTableBorder(firstChild(element, `w:${side}`), theme);
    if (value !== undefined) borders[side] = value;
  }
  return borders;
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

/**
 * 功能：读取 wp:positionH / wp:positionV 的基准与偏移；时间：2026-10-31；作者：lq。
 *
 * OOXML 把偏移写成子元素 wp:posOffset 的文本（EMU），对齐写成 wp:align，
 * 二者只能取其一；这里统一转成 px，交给渲染层做定位。
 */
function readAnchorAxis(
  anchor: Element,
  tag: 'wp:positionH' | 'wp:positionV',
): { relativeFrom?: string; offsetPx?: number; align?: string } {
  const node = firstChild(anchor, tag);
  if (!node) return {};
  const offsetNode = firstChild(node, 'wp:posOffset');
  const rawOffset = offsetNode ? Number((offsetNode.textContent ?? '').trim()) : Number.NaN;
  const alignNode = firstChild(node, 'wp:align');
  return {
    relativeFrom: attrValue(node, 'relativeFrom'),
    offsetPx: Number.isFinite(rawOffset) ? emuToPx(rawOffset) : undefined,
    align: alignNode ? (alignNode.textContent ?? '').trim().toLowerCase() : undefined,
  };
}

/**
 * 功能：把 wp:anchor 解析成结构化锚定信息，保留环绕类型与四周距离；时间：2026-10-31；作者：lq。
 *
 * 旧实现只记录「左/右/上下」方向，丢掉了 distT/distB 等环绕距离，
 * 渲染时又硬编码 8px 下边距，导致浮动图片把下一段文字挤出正常行宽。
 */
function readImageAnchor(drawing: Element): ImageAnchor | undefined {
  const anchor = firstDescendant(drawing, 'wp:anchor');
  if (!anchor) return undefined;
  const distPx = (name: string) => emuToPx(numericAttr(anchor, name, 0) ?? 0);
  const horizontal = readAnchorAxis(anchor, 'wp:positionH');
  const vertical = readAnchorAxis(anchor, 'wp:positionV');

  let wrapType: ImageAnchor['wrapType'] = 'none';
  if (firstDescendant(anchor, 'wp:wrapTopAndBottom')) wrapType = 'topAndBottom';
  else if (firstDescendant(anchor, 'wp:wrapSquare')) wrapType = 'square';
  else if (firstDescendant(anchor, 'wp:wrapTight')) wrapType = 'tight';
  else if (firstDescendant(anchor, 'wp:wrapThrough')) wrapType = 'through';

  const wrapNode =
    firstDescendant(anchor, 'wp:wrapSquare') ??
    firstDescendant(anchor, 'wp:wrapTight') ??
    firstDescendant(anchor, 'wp:wrapThrough');

  return {
    hRelative: horizontal.relativeFrom,
    hOffsetPx: horizontal.offsetPx,
    hAlign: horizontal.align as ImageAnchor['hAlign'],
    vRelative: vertical.relativeFrom,
    vOffsetPx: vertical.offsetPx,
    vAlign: vertical.align as ImageAnchor['vAlign'],
    wrapType,
    wrapText: attrValue(wrapNode, 'wrapText'),
    behindDoc: attrValue(anchor, 'behindDoc') === '1' || attrValue(anchor, 'behindDoc') === 'true',
    distTopPx: distPx('distT'),
    distBottomPx: distPx('distB'),
    distLeftPx: distPx('distL'),
    distRightPx: distPx('distR'),
  };
}

/**
 * 功能：由锚定信息推导环绕方向，方向以 positionH 为准而不是 wrapText；时间：2026-10-31；作者：lq。
 *
 * wrapText 只描述「允许文字出现在图片哪一侧」，真正决定图片贴在左还是右的是 positionH，
 * 二者混用会让右对齐的下划线图片跑到左边，文字绕排方向也随之反转。
 */
function wrapKindFromAnchor(
  anchor: ImageAnchor | undefined,
  pageWidthPx: number,
  imageWidthPx: number,
): DocImage['wrap'] {
  if (!anchor) return 'inline';
  if (anchor.wrapType === 'topAndBottom') return 'top-bottom';
  const alignRight = anchor.hAlign === 'right' || anchor.hAlign === 'outside';
  const alignLeft = anchor.hAlign === 'left' || anchor.hAlign === 'inside';
  if (alignRight) return 'square-right';
  if (alignLeft) return 'square-left';
  const offset = anchor.hOffsetPx ?? 0;
  // 功能：无对齐声明时用偏移量推断贴边方向，超过半栏宽度视为右对齐；时间：2026-10-31；作者：lq。
  return offset > 0 && offset + imageWidthPx / 2 > pageWidthPx / 2 ? 'square-right' : 'square-left';
}

async function parseDrawing(drawing: Element, ctx: ParseContext): Promise<DocRun[]> {
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
  const anchor = readImageAnchor(drawing);
  // 功能：读取图片阴影等效果的外扩占位，避免后续段落提前；时间：2026-11-04；作者：lq。
  const effect = firstDescendant(drawing, 'wp:effectExtent');
  const effectExtent = { left: emuToPx(numericAttr(effect, 'l', 0) ?? 0), top: emuToPx(numericAttr(effect, 't', 0) ?? 0),
    right: emuToPx(numericAttr(effect, 'r', 0) ?? 0), bottom: emuToPx(numericAttr(effect, 'b', 0) ?? 0) };
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
    wrap: wrapKindFromAnchor(anchor, ctx.pageWidthPx, widthPx),
    ...(anchor ? { anchor } : {}),
    ...(Object.values(effectExtent).some((value) => value !== 0) ? { effectExtent } : {}),
  };
  ctx.usedImages.set(rid, { src, part });
  return [{ text: '', style: {}, image }];
}

async function parseVmlPicture(pict: Element, ctx: ParseContext): Promise<DocRun[]> {
  const imageData = firstDescendant(pict, 'v:imagedata');
  const rid = attrValue(imageData, 'r:id') ?? attrValue(imageData, 'r:embed');
  if (!rid) return [];
  const rel = ctx.rels.get(rid);
  const part = rel?.resolved;
  if (!part) return [];
  const src = await loadImageUrl(ctx, part);
  if (!src) return [];
  const shape = firstDescendant(pict, 'v:shape');
  // 功能：只把 VML 图片类型转成图片 run，避免把线条/装饰形状误判成可拖动的大图片框；时间：2026-10-31；作者：lq。
  const shapeType = attrValue(shape, 'type');
  if (shapeType && !/(?:t75|picture)/i.test(shapeType)) return [];
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
  fieldStack: Array<{ instr: string; phase: 'instr' | 'result'; mergeEmitted?: boolean }>;
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
        const raw = child.textContent ?? '';
        // 功能：按 OOXML 语义处理 w:t 空白——未声明 xml:space="preserve" 时首尾空白应被丢弃；时间：2026-10-31；作者：lq。
        // 旧实现无条件保留首尾空格，行宽比 Word 宽，末尾多出的空格会把最后一个字挤到下一行。
        const preserved = child.getAttribute('xml:space') === 'preserve';
        const trimmed = preserved ? raw : raw.replace(/^[ \t\r\n]+/, '').replace(/[ \t\r\n]+$/, '');
        // 功能：严格遵循 Word 语义，w:t 内的普通 CR/LF 只是空白；真正的段内换行由 w:br/w:cr 分支生成；时间：2026-10-31；作者：lq。
        const normalizedText = trimmed.replace(/\r\n?|\n/g, ' ');
        if (normalizedText) {
          const run: DocRun = { text: normalizedText, style };
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
      case 'ptab':
        runs.push({ text: '', style, tab: true });
        break;
      case 'br': {
        const type = attrValue(child, 'w:type');
        runs.push({ text: '', style, field: type === 'page' ? '__PAGE_BREAK__' : undefined, lineBreak: type !== 'page' });
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
      case 'fldChar': {
        const type = attrValue(child, 'w:fldCharType');
        if (type === 'begin') state.fieldStack.push({ instr: '', phase: 'instr' });
        else if (type === 'separate' && topField()) topField()!.phase = 'result';
        else if (type === 'end') {
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
          }
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
      case 'r':
        runs.push(...(await parseRunElement(child, base, ctx, state)));
        break;
      case 'hyperlink':
      case 'ins':
      case 'smartTag':
      case 'customXml':
      case 'sdt': {
        const container =
          localName(child) === 'sdt' ? firstChild(child, 'w:sdtContent') ?? child : child;
        if (localName(child) === 'ins') ctx.hasTrackedChanges = true;
        if (localName(child) === 'sdt') ctx.hasContentControls = true;
        runs.push(...(await parseInlineChildren(container, base, ctx, state)));
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
        const field = /\bPAGE\b/i.test(instr) && !/NUMPAGES/i.test(instr)
          ? 'PAGE'
          : /NUMPAGES/i.test(instr) ? 'NUMPAGES' : /SECTIONPAGES/i.test(instr) ? 'SECTIONPAGES' : undefined;
        if (field) {
          for (let index = before; index < runs.length; index += 1) {
            if (runs[index].field !== '__PAGE_BREAK__') runs[index].field = field;
          }
        }
        break;
      }
      case 'del':
        ctx.hasTrackedChanges = true;
        break;
      case 'bookmarkStart':
      case 'bookmarkEnd':
      case 'commentRangeStart':
      case 'commentRangeEnd':
      case 'proofErr':
        break;
      case 'br': {
        const type = attrValue(child, 'w:type');
        runs.push({ text: '', style: base, field: type === 'page' ? '__PAGE_BREAK__' : undefined, lineBreak: type !== 'page' });
        break;
      }
      // 功能：兼容 WPS/旧版 Word 将控制节点直接挂在段落或超链接下的写法；时间：2026-10-31；作者：lq。
      case 'cr':
        runs.push({ text: '', style: base, lineBreak: true });
        break;
      case 'tab':
      case 'ptab':
        runs.push({ text: '', style: base, tab: true });
        break;
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
): Promise<DocBlock[]> {
  const pPr = firstChild(element, 'w:pPr');
  const styleId = attrValue(firstChild(pPr, 'w:pStyle'), 'w:val');
  const styleFromId = styleId ? ctx.styles.styles.get(styleId) : undefined;
  const styleBase = styleFromId ?? ctx.styles.defaultParagraph ?? baseStyle ?? {
    run: ctx.styles.docDefaults.run,
    para: ctx.styles.docDefaults.para,
  };
  // 功能：先解析段落标记字号，再换算字符缩进；避免 16pt 正文的三字缩进被按文档默认 10pt 缩小；时间：2026-11-05；作者：lq。
  const paragraphRun = applyRunStyle({}, firstChild(pPr, 'w:rPr'), ctx);
  const runBase = mergeRun(
    mergeRun(mergeRun(ctx.styles.docDefaults.run, styleBase.run), baseStyle?.run),
    paragraphRun,
  );
  let paragraphStyle = applyParagraphStyle(
    mergeParagraph(styleBase.para, baseStyle?.para),
    pPr,
    ctx,
    // 功能：按段落样式的字号换算字符缩进，标题等大字号段落的缩进量才准确；时间：2026-11-03；作者：lq。
    runBase.fontSizePt,
  );
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
  const state: ParseInlineState = { fieldStack: [] };
  const runs = await parseInlineChildren(element, runBase, ctx, state);

  const paragraph: DocParagraph = {
    type: 'paragraph',
    id: uid('p'),
    runs,
    style: paragraphStyle,
    list,
    styleId,
  };
  // 功能：保留段尾分节以及各节独立的页眉页脚与编号；时间：2026-10-24；作者：lq。
  const section = firstChild(pPr, 'w:sectPr');
  if (section) {
    const page = parsePageSetup(section, ctx);
    paragraph.sectionEnd = {
      page,
      header: await parseHeaderFooter(section, ctx, 'header', page.margins.header),
      footer: await parseHeaderFooter(section, ctx, 'footer', page.margins.footer),
    };
  }
  // 硬分页按 run 顺序拆块；lastRenderedPageBreak 不生成 run，因此不会变成硬分页。
  const parts: DocRun[][] = [[]];
  for (const run of runs) {
    if (run.field === '__PAGE_BREAK__') parts.push([]);
    else parts[parts.length - 1].push(run);
  }
  if (parts.length === 1) {
    if (paragraph.runs.length === 0) paragraph.runs = [{ text: '', style: runBase }];
    return [paragraph];
  }

  // 纯分页不补空段；分节必须由最后部分的段落承载，不能提前到分页之前。
  const paragraphParts = parts.flatMap((part, index) =>
    part.length > 0 || (index === parts.length - 1 && paragraph.sectionEnd) ? [index] : [],
  );
  const firstPart = paragraphParts[0];
  const lastPart = paragraphParts.at(-1);
  const blocks: DocBlock[] = [];
  for (let index = 0; index < parts.length; index += 1) {
    if (paragraphParts.includes(index)) {
      const first = index === firstPart;
      const last = index === lastPart;
      blocks.push({
        ...paragraph,
        id: first ? paragraph.id : uid('p'),
        runs: parts[index].length > 0 ? parts[index] : [{ text: '', style: runBase }],
        style: {
          ...paragraphStyle,
          ...(!first ? { spaceBeforePx: 0, pageBreakBefore: false, firstLineIndentPx: 0 } : {}),
          ...(!last ? { spaceAfterPx: 0 } : {}),
        },
        list: first ? list : undefined,
        // 续段已展开样式，不能再由样式 ID 恢复段前分页或编号。
        styleId: first ? styleId : undefined,
        sectionEnd: last ? paragraph.sectionEnd : undefined,
      });
    }
    if (index < parts.length - 1) blocks.push({ type: 'pageBreak', id: uid('break') });
  }
  return blocks;
}

/**
 * 功能：解析表格级单元格内边距（w:tblCellMar）；时间：2026-11-03；作者：lq。
 *
 * Word 默认上下为 0、左右 108 twips（5.4pt，即 7.2 CSS px）。未显式声明时返回 undefined，
 * 由渲染层套用 Word 默认值。
 */
function parseTableCellMargin(node: Element | null | undefined, inherited?: TableCellMargin): TableCellMargin | undefined {
  if (!node) return undefined;
  const read = (name: string): number | undefined => {
    const side = firstChild(node, name) ?? firstChild(node, name === 'w:left' ? 'w:start' : name === 'w:right' ? 'w:end' : name);
    if (!side) return undefined;
    const width = numericAttr(side, 'w:w');
    if (width === undefined) return undefined;
    return attrValue(side, 'w:type') === 'nil' ? 0 : twipsToPx(width);
  };
  const top = read('w:top');
  const right = read('w:right');
  const bottom = read('w:bottom');
  const left = read('w:left');
  if (top === undefined && right === undefined && bottom === undefined && left === undefined) {
    return undefined;
  }
  return {
    top: top ?? inherited?.top ?? 0,
    right: right ?? inherited?.right ?? twipsToPx(108),
    bottom: bottom ?? inherited?.bottom ?? 0,
    left: left ?? inherited?.left ?? twipsToPx(108),
  };
}

async function parseTable(element: Element, ctx: ParseContext): Promise<DocTable> {
  const tblPr = firstChild(element, 'w:tblPr');
  const grid = firstChild(element, 'w:tblGrid');
  let columnWidths = Array.from(grid?.childNodes ?? [])
    .filter((node): node is Element => node instanceof Element && localName(node) === 'gridCol')
    .map((col) => twipsToPx(numericAttr(col, 'w:w', 0) ?? 0))
    .filter((width) => width > 0);

  const styleId = attrValue(firstChild(tblPr, 'w:tblStyle'), 'w:val');
  // 功能：表格样式可能使用 a4 等内部 ID，边框从样式继承后再应用直接格式；时间：2026-11-03；作者：lq。
  const tableStyle = styleId ? ctx.styles.styles.get(styleId) : ctx.styles.defaultTable;
  const borders = { ...tableStyle?.table?.borders, ...parseTableBorders(firstChild(tblPr, 'w:tblBorders'), ctx.themeColors) };
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
  // 功能：解析列宽布局模式与表格缩进；时间：2026-11-03；作者：lq。
  // 功能：保留布局模式供显示层使用，不在导入阶段按版心篡改列宽；时间：2026-11-03；作者：lq。
  const layoutType = attrValue(firstChild(tblPr, 'w:tblLayout'), 'w:type');
  const layout: DocTable['layout'] = layoutType === 'fixed' ? 'fixed' : 'autofit';
  const tableIndentTwips = numericAttr(firstChild(tblPr, 'w:tblInd'), 'w:w');
  const indentPx = tableIndentTwips !== undefined ? twipsToPx(tableIndentTwips) : undefined;
  const cellMarginPx = parseTableCellMargin(firstChild(tblPr, 'w:tblCellMar'), tableStyle?.table?.margin) ?? tableStyle?.table?.margin;

  interface RawCell {
    cell: DocTableCell;
    vMerge?: 'restart' | 'continue';
    colSpan: number;
  }

  const rawRows: Array<{ cells: RawCell[]; heightPx?: number; heightRule?: 'atLeast' | 'exact'; isHeader: boolean }> = [];
  let colCount = columnWidths.length;

  for (const rowElement of Array.from(element.childNodes)) {
    if (!(rowElement instanceof Element) || localName(rowElement) !== 'tr') continue;
    const trPr = firstChild(rowElement, 'w:trPr');
    const heightNode = firstChild(trPr, 'w:trHeight');
    const height = numericAttr(heightNode, 'w:val');
    // 功能：行高规则 w:trHeight@w:hRule——exact 为固定行高，缺省按 atLeast；时间：2026-11-03；作者：lq。
    // 此前统一按 minmax(下限, auto) 渲染，exact 的行会被内容撑高，与 Word 不一致。
    const heightRuleValue = attrValue(heightNode, 'w:hRule');
    const heightRule: 'atLeast' | 'exact' = heightRuleValue === 'exact' ? 'exact' : 'atLeast';
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
      // 功能：单元格底纹支持主题色填充（Office 主题色表头常见写法）；时间：2026-11-01；作者：lq。
      const backgroundColor = resolveWordShading(firstChild(tcPr, 'w:shd'), ctx.themeColors);
      const vAlignValue = attrValue(firstChild(tcPr, 'w:vAlign'), 'w:val');
      const verticalAlign =
        vAlignValue === 'center' ? 'middle' : vAlignValue === 'bottom' ? 'bottom' : 'top';
      const cellBorders = parseTableBorders(firstChild(tcPr, 'w:tcBorders'), ctx.themeColors);
      // 功能：单元格仅存自身覆盖值；内框与外框由渲染器根据行列位置从表格继承；时间：2026-11-03；作者：lq。
      const mergedBorders: TableBorders = cellBorders;
      // 功能：单元格自定义内边距（w:tcMar）覆盖表格级 tblCellMar；时间：2026-11-03；作者：lq。
      const cellMargin = parseTableCellMargin(firstChild(tcPr, 'w:tcMar'), cellMarginPx);
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
          marginPx: cellMargin,
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
      heightRule,
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
        heightRule: rawRow.heightRule,
        isHeader: rawRow.isHeader,
      });
    }
  }

  // 功能：导入保留原始 tblGrid；显示适配不能改写源列宽，尤其是 fixed 表格；时间：2026-11-03；作者：lq。

  return {
    type: 'table',
    id: uid('tbl'),
    rows,
    columnWidthsPx: columnWidths,
    widthPct,
    alignment: alignmentValue ? alignmentMap[alignmentValue] ?? 'left' : 'left',
    borders,
    styleId,
    layout,
    indentPx,
    cellMarginPx,
  };
}

/**
 * 功能：扫描段落内的域字符与域指令，用于识别跨段落的 TOC 域边界；时间：2026-10-30；作者：lq。
 *
 * Word 的目录域从第一个条目的段落开始、在最后一个条目段落结束，中间还嵌着 PAGEREF 等子域，
 * 因此必须统计 begin/end 的净深度，而不能只看有没有出现 `w:fldChar`。
 */
function paragraphFieldScan(paragraph: Element): { begins: number; ends: number; instr: string } {
  let begins = 0;
  let ends = 0;
  let instr = '';
  const visit = (node: Element) => {
    for (const child of Array.from(node.childNodes)) {
      if (!(child instanceof Element)) continue;
      const name = localName(child);
      if (name === 'instrText') instr += child.textContent ?? '';
      else if (name === 'fldChar') {
        const type = attrValue(child, 'w:fldCharType');
        if (type === 'begin') begins += 1;
        else if (type === 'end') ends += 1;
      }
      visit(child);
    }
  };
  visit(paragraph);
  return { begins, ends, instr };
}

/** 功能：从域指令中识别 TOC 域并取出标题级别上限（`\o "1-3"`），非 TOC 域返回 undefined；时间：2026-10-30；作者：lq。 */
function tocLevelsFromInstruction(instr: string): number | undefined {
  if (!/^\s*TOC\b/i.test(instr)) return undefined;
  const range = /\\o\s+"(\d)\s*-\s*(\d)"/i.exec(instr);
  if (range) return Math.min(9, Math.max(1, Number(range[2]) || 3));
  const single = /\\o\s+"(\d)"/i.exec(instr);
  if (single) return Math.min(9, Math.max(1, Number(single[1]) || 3));
  return 3;
}

async function parseBlockContainer(element: Element, ctx: ParseContext): Promise<DocBlock[]> {
  const blocks: DocBlock[] = [];
  for (const child of Array.from(element.childNodes)) {
    if (!(child instanceof Element)) continue;
    // 功能：每个源顶层节点可展开为多个块（分页、内容控件），保存时按组恢复；时间：2026-11-03；作者：lq。
    const groupStart = blocks.length;
    switch (localName(child)) {
      case 'p': {
        // 功能：把 TOC 域的缓存结果换成可动态重算的目录段落，避免打开后仍显示过期条目；时间：2026-10-30；作者：lq。
        const field = paragraphFieldScan(child);
        if (ctx.toc) {
          ctx.toc.depth += field.begins - field.ends;
          if (ctx.toc.depth <= 0) ctx.toc = undefined;
          break;
        }
        const tocLevels = tocLevelsFromInstruction(field.instr);
        if (tocLevels !== undefined) {
          const depth = field.begins - field.ends;
          if (depth > 0) ctx.toc = { levels: tocLevels, depth };
          blocks.push({
            type: 'paragraph',
            id: uid('p'),
            runs: [{ text: '', style: {} }],
            // 功能：目录段落带一个右对齐点线制表位，写回 DOCX 时页码才能对齐右边界；时间：2026-10-30；作者：lq。
            style: {
              tabStops: [
                { posPx: Math.max(120, ctx.pageWidthPx - 192), alignment: 'right', leader: 'dot' },
              ],
            },
            toc: { levels: tocLevels },
          });
          break;
        }
        blocks.push(...(await parseParagraph(child, ctx)));
        break;
      }
      case 'tbl':
        blocks.push(await parseTable(child, ctx));
        break;
      case 'sdt': {
        ctx.hasContentControls = true;
        const content = firstChild(child, 'w:sdtContent');
        if (content) blocks.push(...(await parseBlockContainer(content, ctx)));
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
    if (localName(element) === 'body') ctx.sourceGroups?.push(blocks.slice(groupStart).map((block) => block.id));
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
  // 功能：读取文档网格，公文类文档靠 linePitch 决定每页行数；时间：2026-11-02；作者：lq。
  const docGridNode = firstChild(sectPr, 'w:docGrid');
  // 功能：w:type 缺省或为 default 时按「只限定行网格」处理，避免写出非法属性值；时间：2026-11-02；作者：lq。
  const rawGridType = attrValue(docGridNode, 'w:type') ?? 'lines';
  const gridType: DocGrid['type'] =
    rawGridType === 'none' || rawGridType === 'linesAndChars' || rawGridType === 'snapToChars'
      ? rawGridType
      : 'lines';
  const linePitch = numericAttr(docGridNode, 'w:linePitch');
  const charSpace = numericAttr(docGridNode, 'w:charSpace');
  const docGrid =
    docGridNode && gridType !== 'none' && linePitch
      ? {
          type: gridType,
          linePitchPx: twipsToPx(linePitch),
          charSpacePx: charSpace ? (charSpace / 4096) * (96 / 72) : undefined,
        }
      : undefined;
  return {
    widthPx: twipsToPx(widthTwips),
    heightPx: twipsToPx(heightTwips),
    orientation,
    sizeName: matched?.name ?? (swapped ? '自定义' : '自定义'),
    margins,
    columns: columns && columns > 1 ? columns : undefined,
    // 功能：恢复 DOCX 起始页码；时间：2026-10-24；作者：lq。
    pageNumberStart: numericAttr(firstChild(sectPr, 'w:pgNumType'), 'w:start'),
    docGrid,
  };
}

async function parseHeaderFooter(
  sectPr: Element | null,
  ctx: ParseContext,
  kind: 'header' | 'footer',
  distancePx: number,
): Promise<HeaderFooter | undefined> {
  if (!sectPr) return undefined;
  const references = Array.from(sectPr.childNodes).filter(
    (node): node is Element =>
      node instanceof Element && localName(node) === `${kind}Reference`,
  );
  if (references.length === 0) return undefined;
  const preferred =
    references.find((reference) => attrValue(reference, 'w:type') === 'default') ?? references[0];
  const relId = preferred.getAttribute('r:id') ?? getAttrLocal(preferred, 'id');
  if (!relId) return undefined;
  const rel = ctx.rels.get(relId);
  if (!rel?.resolved) return undefined;
  const xml = await ctx.pkg.text(rel.resolved);
  const doc = parseXmlDom(xml);
  const root = firstDescendant(doc, kind === 'header' ? 'w:hdr' : 'w:ftr') ?? doc.documentElement;
  // 功能：页眉页脚图片关系按自身部件解析；时间：2026-10-24；作者：lq。
  const partContext = { ...ctx, rels: await relationshipMap(ctx.pkg, rel.resolved) };
  const blocks = await parseBlockContainer(root, partContext);
  // 功能：单独解析 DrawingML 文本框，其文字不是图片，不能因没有 a:blip 而丢掉页码等域；时间：2026-11-05；作者：lq。
  const textBoxes: NonNullable<HeaderFooter['textBoxes']> = [];
  for (const content of Array.from(root.getElementsByTagNameNS('*', 'txbxContent'))) {
    let drawing: Element | null = content.parentElement;
    while (drawing && localName(drawing) !== 'drawing') drawing = drawing.parentElement;
    if (!drawing) continue;
    const extent = firstDescendant(drawing, 'wp:extent');
    const bodyPr = firstDescendant(drawing, 'wps:bodyPr');
    textBoxes.push({ id: uid('hf-box'), blocks: await parseBlockContainer(content, partContext),
      widthPx: emuToPx(numericAttr(extent, 'cx', 0) ?? 0), heightPx: emuToPx(numericAttr(extent, 'cy', 0) ?? 0),
      anchor: readImageAnchor(drawing), autoFit: !!firstDescendant(bodyPr, 'a:spAutoFit'),
      noWrap: attrValue(bodyPr, 'wrap') === 'none',
      insetsPx: { left: emuToPx(numericAttr(bodyPr, 'lIns', 0) ?? 0), top: emuToPx(numericAttr(bodyPr, 'tIns', 0) ?? 0),
        right: emuToPx(numericAttr(bodyPr, 'rIns', 0) ?? 0), bottom: emuToPx(numericAttr(bodyPr, 'bIns', 0) ?? 0) },
      sourceXml: new XMLSerializer().serializeToString(drawing) });
  }
  const pageNumbers = [...blocks, ...textBoxes.flatMap((box) => box.blocks)].some((block) =>
    block.type === 'paragraph' ? block.runs.some((run) => run.field === 'PAGE') : false,
  );
  return { blocks, distancePx, pageNumbers, sourcePart: rel.resolved, ...(textBoxes.length ? { textBoxes } : {}) };
}

export interface ParseDocxOptions {
  fileName?: string;
  /** 功能：只对确认未安装的旧 GB 字体应用兼容替代，不覆盖已安装字体；时间：2026-11-05；作者：lq。 */
  knownFonts?: string[];
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

  // 功能：读取源 OOXML 的 w:altName，备用名仅参与显示，不改写文件原有主字体；时间：2026-11-05；作者：lq。
  const fontAlternates = new Map<string, string>();
  if (pkg.has('word/fontTable.xml')) {
    const fontTable = parseXmlDom(await pkg.text('word/fontTable.xml'));
    for (const font of Array.from(fontTable.getElementsByTagNameNS('*', 'font'))) {
      const name = attrValue(font, 'w:name');
      const alternate = attrValue(firstChild(font, 'w:altName'), 'w:val');
      if (name && alternate && name !== alternate) fontAlternates.set(name, alternate);
    }
  }
  // 功能：Word 16.114/macOS 对缺失 GB2312 字体采用微软雅黑；保留主名，仅在有安装清单且原字体缺失时添加显示备用名；时间：2026-11-05；作者：lq。
  if (options.knownFonts) {
    const installed = new Set(options.knownFonts.map((name) => name.normalize('NFKC').toLowerCase().trim()));
    for (const name of ['仿宋_GB2312', 'FangSong_GB2312', '楷体_GB2312', 'KaiTi_GB2312']) {
      if (!installed.has(name.normalize('NFKC').toLowerCase()) && !fontAlternates.has(name)) fontAlternates.set(name, '微软雅黑');
    }
  }

  const provisional: ParseContext = {
    fontAlternates,
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
  // 功能：记录版心宽度，供表格自适应收敛使用（Word 不允许表格越过版心右边界）；时间：2026-11-03；作者：lq。
  provisional.contentWidthPx = Math.max(
    120,
    page.widthPx - page.margins.left - page.margins.right,
  );

  provisional.sourceGroups = [];
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

  // 功能：读回模板里保存的邮件合并面板状态，供打开时恢复面板；时间：2026-10-29；作者：lq。
  let defaultTabStopPx: number | undefined;
  if (pkg.has('word/settings.xml')) {
    try {
      const settingsDocument = parseXmlDom(await pkg.text('word/settings.xml'));
      // 功能：读取文档默认制表位间距，Tab 在无自定义制表位时按它等距推进；时间：2026-10-30；作者：lq。
      const tabStopElement = settingsDocument.getElementsByTagName('w:defaultTabStop')[0];
      const tabStopTwips = tabStopElement ? Number(tabStopElement.getAttribute('w:val')) : NaN;
      if (Number.isFinite(tabStopTwips) && tabStopTwips > 0) {
        defaultTabStopPx = twipsToPx(tabStopTwips);
      }
      const docVars = settingsDocument.getElementsByTagName('w:docVar');
      for (const docVar of Array.from(docVars)) {
        if (docVar.getAttribute('w:name') === 'mailMergePanelState') {
          const value = docVar.getAttribute('w:val');
          if (value) (meta as Record<string, unknown>).mailMergeState = value;
          break;
        }
      }
    } catch {
      // 忽略读取失败，不影响正文解析
    }
  }

  const model: DocModel = {
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
    ...(defaultTabStopPx ? { defaultTabStopPx } : {}),
  };
  model.source = captureDocumentSource(model, documentXml, provisional.sourceGroups);
  return model;
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
