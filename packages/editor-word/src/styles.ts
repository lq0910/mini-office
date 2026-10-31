/**
 * 功能：Word 样式库——内置预设样式、样式套用、样式识别与「我的样式集」持久化；时间：2026-10-29；作者：lq。
 *
 * 为什么需要这个模块：编辑器的段落格式是扁平字段（对齐、缩进、行距、字号…），
 * 用户每次排一份文档都要重复调整同一组参数。样式库把「一组格式」命名保存，
 * 一次点击即可批量套用，并把样式名写回 DOCX 的 `w:pStyle`，Word / WPS 打开后同样可识别。
 *
 * 三条设计约束：
 * 1. 纯函数：套用、识别、解析都不碰 DOM 与 React，便于单元测试与在导出流程复用。
 * 2. 样式只描述「应当是什么」，不保存继承结果；应用时再与文档默认字体合并，
 *    因此文档换默认字体后，正文样式会跟着变，标题样式仍保留自己的字号与粗细。
 * 3. 内置预设沿用 Word 的 styleId（Heading1 / Title / Quote…），保证与 Word 互通；
 *    自定义样式使用 `MoStyle_` 前缀，避免与 Word 内置样式冲突。
 */
import type {
  DocModel,
  DocParagraph,
  DocRun,
  DocStyleDefinition,
  ParagraphStyle,
  RunStyle,
} from '@mini-office/document-model';
import { charWidthPx } from './typography';

/** 1pt = 4/3 px，与 typography 中的换算保持一致。 */
const pt = (value: number): number => Math.round(value * (4 / 3) * 100) / 100;

/** 功能：以正文字号折算「n 个字符」的缩进像素值；时间：2026-10-29；作者：lq。 */
function chars(count: number, fontSizePt = 10.5): number {
  return Math.round(charWidthPx(fontSizePt) * count * 100) / 100;
}

/**
 * 功能：内置预设样式，顺序即样式库默认展示顺序（前 8 个常驻，其余收在「更多」里）；时间：2026-10-29；作者：lq。
 *
 * 版式取值参考 Word 中文默认模板：标题使用「字号 + 加粗 + 段前段后」建立层级，
 * 不锁定字体，避免把用户选定的正文字体强行换成黑体。
 */
export const BUILTIN_STYLES: ReadonlyArray<DocStyleDefinition> = [
  {
    id: 'Normal',
    name: '正文',
    builtinName: 'Normal',
    origin: 'preset',
    paragraph: {
      alignment: 'left',
      lineHeight: 1.5,
      spaceBeforePx: 0,
      spaceAfterPx: 0,
      firstLineIndentPx: 0,
      indentLeftPx: 0,
      indentRightPx: 0,
    },
    // 字号与字体留空，应用时回落到文档默认值，保证「正文」永远是文档基准。
    run: {},
  },
  {
    id: 'Heading1',
    name: '标题 1',
    builtinName: 'heading 1',
    origin: 'preset',
    paragraph: {
      alignment: 'left',
      outlineLevel: 0,
      keepNext: true,
      lineHeight: 1.3,
      spaceBeforePx: pt(12),
      spaceAfterPx: pt(6),
      firstLineIndentPx: 0,
      indentLeftPx: 0,
    },
    run: { fontSizePt: 18, bold: true },
  },
  {
    id: 'Heading2',
    name: '标题 2',
    builtinName: 'heading 2',
    origin: 'preset',
    paragraph: {
      alignment: 'left',
      outlineLevel: 1,
      keepNext: true,
      lineHeight: 1.35,
      spaceBeforePx: pt(10),
      spaceAfterPx: pt(4),
      firstLineIndentPx: 0,
      indentLeftPx: 0,
    },
    run: { fontSizePt: 15, bold: true },
  },
  {
    id: 'Heading3',
    name: '标题 3',
    builtinName: 'heading 3',
    origin: 'preset',
    paragraph: {
      alignment: 'left',
      outlineLevel: 2,
      keepNext: true,
      lineHeight: 1.4,
      spaceBeforePx: pt(8),
      spaceAfterPx: pt(4),
      firstLineIndentPx: 0,
      indentLeftPx: 0,
    },
    run: { fontSizePt: 13, bold: true },
  },
  {
    id: 'Heading4',
    name: '标题 4',
    builtinName: 'heading 4',
    origin: 'preset',
    paragraph: {
      alignment: 'left',
      outlineLevel: 3,
      keepNext: true,
      lineHeight: 1.45,
      spaceBeforePx: pt(6),
      spaceAfterPx: pt(3),
      firstLineIndentPx: 0,
      indentLeftPx: 0,
    },
    run: { fontSizePt: 11, bold: true },
  },
  {
    id: 'Title',
    name: '标题',
    builtinName: 'Title',
    origin: 'preset',
    paragraph: {
      alignment: 'center',
      lineHeight: 1.25,
      spaceBeforePx: pt(12),
      spaceAfterPx: pt(12),
      firstLineIndentPx: 0,
      indentLeftPx: 0,
    },
    run: { fontSizePt: 24, bold: true },
  },
  {
    id: 'Subtitle',
    name: '副标题',
    builtinName: 'Subtitle',
    origin: 'preset',
    paragraph: {
      alignment: 'center',
      lineHeight: 1.35,
      spaceBeforePx: 0,
      spaceAfterPx: pt(12),
      firstLineIndentPx: 0,
      indentLeftPx: 0,
    },
    run: { fontSizePt: 14, color: '#475467' },
  },
  {
    id: 'Quote',
    name: '引用',
    builtinName: 'Quote',
    origin: 'preset',
    paragraph: {
      alignment: 'left',
      lineHeight: 1.5,
      spaceBeforePx: pt(6),
      spaceAfterPx: pt(6),
      indentLeftPx: chars(2),
      indentRightPx: chars(2),
      firstLineIndentPx: 0,
      borders: { left: { style: 'single', widthPx: 3, color: '#D0D5DD' } },
    },
    run: { fontFamily: '楷体', fontSizePt: 10.5, color: '#475467' },
  },
  {
    id: 'Caption',
    name: '题注',
    builtinName: 'caption',
    origin: 'preset',
    paragraph: {
      alignment: 'center',
      lineHeight: 1.3,
      spaceBeforePx: pt(3),
      spaceAfterPx: pt(6),
      firstLineIndentPx: 0,
      indentLeftPx: 0,
    },
    run: { fontSizePt: 9, color: '#667085' },
  },
  {
    id: 'MoBodyIndent',
    name: '正文首行缩进',
    origin: 'preset',
    paragraph: {
      alignment: 'justify',
      lineHeight: 1.5,
      spaceBeforePx: 0,
      spaceAfterPx: 0,
      firstLineIndentPx: chars(2),
      indentLeftPx: 0,
    },
    run: {},
  },
  {
    id: 'NoSpacing',
    name: '无间隔',
    builtinName: 'No Spacing',
    origin: 'preset',
    paragraph: {
      alignment: 'left',
      lineHeight: 1,
      spaceBeforePx: 0,
      spaceAfterPx: 0,
      firstLineIndentPx: 0,
      indentLeftPx: 0,
    },
    run: {},
  },
  {
    id: 'IntenseQuote',
    name: '强调引用',
    builtinName: 'Intense Quote',
    origin: 'preset',
    paragraph: {
      alignment: 'center',
      lineHeight: 1.5,
      spaceBeforePx: pt(10),
      spaceAfterPx: pt(10),
      indentLeftPx: chars(2),
      indentRightPx: chars(2),
      firstLineIndentPx: 0,
      borders: {
        top: { style: 'single', widthPx: 1.5, color: '#98A2B3' },
        bottom: { style: 'single', widthPx: 1.5, color: '#98A2B3' },
      },
    },
    run: { fontSizePt: 12, bold: true, italic: true, color: '#344054' },
  },
  {
    id: 'ListParagraph',
    name: '列表段落',
    builtinName: 'List Paragraph',
    origin: 'preset',
    paragraph: {
      alignment: 'left',
      lineHeight: 1.5,
      spaceBeforePx: 0,
      spaceAfterPx: 0,
      indentLeftPx: chars(2),
      firstLineIndentPx: -chars(2),
    },
    run: {},
  },
  {
    id: 'MoListBullet',
    name: '列表项目符号',
    origin: 'preset',
    paragraph: {
      alignment: 'left',
      lineHeight: 1.5,
      spaceBeforePx: 0,
      spaceAfterPx: 0,
      indentLeftPx: chars(2),
      firstLineIndentPx: -chars(2),
    },
    run: {},
    list: { ordered: false, level: 0 },
  },
  {
    id: 'MoListNumber',
    name: '列表编号',
    origin: 'preset',
    paragraph: {
      alignment: 'left',
      lineHeight: 1.5,
      spaceBeforePx: 0,
      spaceAfterPx: 0,
      indentLeftPx: chars(2),
      firstLineIndentPx: -chars(2),
    },
    run: {},
    list: { ordered: true, level: 0 },
  },
  {
    id: 'MoMono',
    name: '等宽文本',
    builtinName: 'Macro Text',
    origin: 'preset',
    paragraph: {
      alignment: 'left',
      lineHeight: 1.4,
      spaceBeforePx: 0,
      spaceAfterPx: 0,
      firstLineIndentPx: 0,
      indentLeftPx: 0,
    },
    run: { fontFamily: 'Courier New', fontSizePt: 9 },
  },
  // 功能：内置 TOC1~TOC9 目录样式（右对齐制表位 + 点前导符），与 Word/WPS 默认一致；时间：2026-10-30；作者：hou0。
  ...(() => {
    const tocStyles: DocStyleDefinition[] = [];
    for (let i = 1; i <= 9; i++) {
      tocStyles.push({
        id: `TOC${i}`,
        name: `目录 ${i}`,
        builtinName: `TOC ${i}`,
        origin: 'preset',
        paragraph: {
          alignment: 'left',
          lineHeight: 1.2,
          spaceBeforePx: 0,
          spaceAfterPx: pt(2),
          firstLineIndentPx: 0,
          indentLeftPx: i * 18,
          tabStops: [{ posPx: 400, kind: 'right', leader: 'dot' }],
        },
        run: {},
      });
    }
    return tocStyles;
  })(),
];

/** 常驻展示的预设数量，其余收在「显示更多样式」里，与 WPS 的折叠策略一致。 */
export const BASIC_STYLE_COUNT = 8;

const BUILTIN_IDS = new Set(BUILTIN_STYLES.map((style) => style.id));

/** 功能：判断样式是否为内置预设；时间：2026-10-29；作者：lq。 */
export function isBuiltinStyle(id: string): boolean {
  return BUILTIN_IDS.has(id);
}

/** 功能：生成不与现有样式冲突的自定义样式 ID，写入 DOCX 时作为 w:styleId；时间：2026-10-29；作者：lq。 */
export function newStyleId(taken: Iterable<string> = []): string {
  const used = new Set(taken);
  for (let attempt = 0; attempt < 50; attempt += 1) {
    const candidate = `MoStyle_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
    if (!used.has(candidate)) return candidate;
  }
  return `MoStyle_${Math.random().toString(36).slice(2, 10)}`;
}

/**
 * 功能：合并内置预设与文档自身样式，文档内定义优先（同名 ID 覆盖预设）；时间：2026-10-29；作者：lq。
 * 这样打开一份 Word 文档时，其「标题 1」会以文档里的真实格式展示，而不是我们的默认值。
 */
export function effectiveStyles(doc: Pick<DocModel, 'styles'> | undefined): DocStyleDefinition[] {
  const docStyles = doc?.styles ?? [];
  const overrides = new Map(docStyles.map((style) => [style.id, style]));
  const presets = BUILTIN_STYLES.map((style) => {
    const override = overrides.get(style.id);
    if (!override) return style;
    // 文档里的内置样式名是英文（heading 1），界面上仍用中文名，格式则沿用文档真实定义。
    const englishName = (style.builtinName ?? '').toLowerCase();
    const localized = override.name.toLowerCase() === englishName ? style.name : override.name;
    return { ...override, name: localized };
  });
  const extras = docStyles.filter((style) => !BUILTIN_IDS.has(style.id));
  return [...presets, ...extras];
}

/** 功能：样式库中「本文档新增」的样式（自定义 + 从 DOCX 导入），不含内置预设；时间：2026-10-29；作者：lq。 */
export function documentStyles(doc: Pick<DocModel, 'styles'> | undefined): DocStyleDefinition[] {
  return (doc?.styles ?? []).filter((style) => !BUILTIN_IDS.has(style.id));
}

/** 功能：把样式的字体设置与文档默认值合并，得到可直接写入 run 的完整字体格式；时间：2026-10-29；作者：lq。 */
export function resolveStyleRun(
  style: Pick<DocStyleDefinition, 'run'>,
  doc: Pick<DocModel, 'defaultFont' | 'defaultFontSizePt'>,
): RunStyle {
  return {
    ...style.run,
    fontFamily: style.run.fontFamily ?? doc.defaultFont,
    fontSizePt: style.run.fontSizePt ?? doc.defaultFontSizePt,
  };
}

/**
 * 功能：把段落套用为指定样式，返回新的段落对象；时间：2026-10-29；作者：lq。
 *
 * 语义说明：样式是一次「完整替换」而不是「叠加」——段落级的对齐、缩进、间距、行距、边框
 * 都会被样式重写；字体则按 resolveStyleRun 结果整体写入每个 run，仅保留用户层面的
 * 下划线、删除线、突出显示与上下标，避免一次套用把用户标好的重点抹掉。
 */
export function applyStyleToParagraph(
  paragraph: DocParagraph,
  style: DocStyleDefinition,
  options: { doc: Pick<DocModel, 'defaultFont' | 'defaultFontSizePt'>; marker?: string },
): DocParagraph {
  const resolved = resolveStyleRun(style, options.doc);
  const runs = paragraph.runs.map((run) => ({ ...run, style: mergeRunStyle(run.style, resolved) }));
  const list = style.list
    ? {
        level: style.list.level,
        ordered: style.list.ordered,
        marker: options.marker ?? (style.list.ordered ? '1.' : '•'),
      }
    : undefined;
  return {
    ...paragraph,
    style: { ...style.paragraph },
    runs,
    list,
    // 正文是文档默认样式，写 pStyle="Normal" 只增加噪音，保持与 Word 一致的「默认样式不显式声明」。
    styleId: style.id === 'Normal' ? undefined : style.id,
  };
}

/** 功能：样式套用时的字体合并规则，未在样式中声明的强调属性（下划线/高亮等）保持不变；时间：2026-10-29；作者：lq。 */
function mergeRunStyle(base: RunStyle, resolved: RunStyle): RunStyle {
  const next: RunStyle = { ...base };
  next.fontFamily = resolved.fontFamily;
  next.fontSizePt = resolved.fontSizePt;
  // 用 undefined 显式清掉旧值：从「标题 1」回到「正文」时必须去掉加粗，否则格式不会变。
  next.bold = resolved.bold ? true : undefined;
  next.italic = resolved.italic ? true : undefined;
  next.color = resolved.color;
  next.allCaps = resolved.allCaps;
  next.smallCaps = resolved.smallCaps;
  return next;
}

/** 功能：清除段落格式，回到文档默认字体与无缩进、无间距的状态；时间：2026-10-29；作者：lq。 */
export function clearParagraphFormatting(
  paragraph: DocParagraph,
  doc: Pick<DocModel, 'defaultFont' | 'defaultFontSizePt'>,
): DocParagraph {
  return {
    ...paragraph,
    style: {
      alignment: 'left',
      lineHeight: 1.5,
      spaceBeforePx: 0,
      spaceAfterPx: 0,
      firstLineIndentPx: 0,
      indentLeftPx: 0,
      indentRightPx: 0,
    },
    runs: paragraph.runs.map((run) => ({
      ...run,
      style: { fontFamily: doc.defaultFont, fontSizePt: doc.defaultFontSizePt },
    })),
    list: undefined,
    styleId: undefined,
  };
}

function near(a: number | undefined, b: number | undefined, tolerance: number): boolean {
  return Math.abs((a ?? 0) - (b ?? 0)) <= tolerance;
}

function lineValue(style: ParagraphStyle): number {
  return style.lineHeightExactPx ?? style.lineHeightAtLeastPx ?? style.lineHeight ?? 1.5;
}

/** 功能：判断段落现有格式是否与某个样式一致，用于工具栏显示当前样式名；时间：2026-10-29；作者：lq。 */
export function matchesStyle(
  paragraph: DocParagraph,
  style: DocStyleDefinition,
  doc: Pick<DocModel, 'defaultFont' | 'defaultFontSizePt'>,
): boolean {
  const actual = paragraph.style;
  const expected = style.paragraph;
  if ((actual.alignment ?? 'left') !== (expected.alignment ?? 'left')) return false;
  if ((actual.outlineLevel ?? -1) !== (expected.outlineLevel ?? -1)) return false;
  if (!near(actual.firstLineIndentPx, expected.firstLineIndentPx, 2)) return false;
  if (!near(actual.indentLeftPx, expected.indentLeftPx, 2)) return false;
  if (!near(actual.spaceBeforePx, expected.spaceBeforePx, 2)) return false;
  if (!near(actual.spaceAfterPx, expected.spaceAfterPx, 2)) return false;
  const exact = !!(expected.lineHeightExactPx || expected.lineHeightAtLeastPx);
  if (Math.abs(lineValue(actual) - lineValue(expected)) > (exact ? 2 : 0.06)) return false;
  if (!!paragraph.list !== !!style.list) return false;

  const run = paragraph.runs.find((item) => item.text)?.style ?? {};
  const resolved = resolveStyleRun(style, doc);
  const fontSize = run.fontSizePt ?? doc.defaultFontSizePt;
  if (Math.abs(fontSize - (resolved.fontSizePt ?? doc.defaultFontSizePt)) > 0.2) return false;
  if (!!run.bold !== !!resolved.bold) return false;
  if (!!run.italic !== !!resolved.italic) return false;
  return true;
}

/**
 * 功能：识别段落当前使用的样式 ID；时间：2026-10-29；作者：lq。
 * 优先信任 pStyle（保存/导入后可精确还原），否则按排版特征匹配，
 * 使「用户手动调成 18 磅加粗」的段落也能显示为「标题 1」，避免工具栏显示空白。
 */
export function detectStyleId(
  paragraph: DocParagraph,
  styles: ReadonlyArray<DocStyleDefinition>,
  doc: Pick<DocModel, 'defaultFont' | 'defaultFontSizePt'>,
): string | undefined {
  if (paragraph.styleId) {
    const matched = styles.find((style) => style.id === paragraph.styleId);
    if (matched) return matched.id;
  }
  return styles.find((style) => matchesStyle(paragraph, style, doc))?.id;
}

/** 功能：按 ID 取样式，未命中返回 undefined；时间：2026-10-29；作者：lq。 */
export function findStyle(
  styles: ReadonlyArray<DocStyleDefinition>,
  id: string | undefined,
): DocStyleDefinition | undefined {
  if (!id) return undefined;
  return styles.find((style) => style.id === id);
}

/** 功能：以段落当前格式为蓝本创建自定义样式；时间：2026-10-29；作者：lq。 */
export function createStyleFromParagraph(
  name: string,
  paragraph: DocParagraph,
  doc: Pick<DocModel, 'defaultFont' | 'defaultFontSizePt' | 'styles'>,
): DocStyleDefinition {
  const sourceRun = paragraph.runs.find((run) => run.text)?.style ?? paragraph.runs[0]?.style ?? {};
  const resolved = resolveStyleRun({ run: sourceRun }, doc);
  return {
    id: newStyleId((doc.styles ?? []).map((style) => style.id)),
    name: name.trim(),
    origin: 'custom',
    basedOn: 'Normal',
    paragraph: { ...paragraph.style },
    run: {
      fontFamily: resolved.fontFamily,
      fontSizePt: resolved.fontSizePt,
      bold: resolved.bold,
      italic: resolved.italic,
      color: resolved.color,
      letterSpacingPx: resolved.letterSpacingPx,
    },
    list: paragraph.list ? { ordered: paragraph.list.ordered, level: paragraph.list.level } : undefined,
  };
}

/* -------------------------------------------------------------------------- */
/* 我的样式集（跨文档复用）                                                     */
/* -------------------------------------------------------------------------- */

export interface WordStyleSet {
  id: string;
  name: string;
  styles: DocStyleDefinition[];
  savedAt: string;
}

export const STYLE_SET_STORAGE_KEY = 'mini-office.word.style-sets.v1';

const MAX_STYLE_SETS = 30;

function defaultStorage(): Storage | undefined {
  try {
    return typeof window !== 'undefined' ? window.localStorage : undefined;
  } catch {
    return undefined;
  }
}

function sanitizeDefinition(value: unknown): DocStyleDefinition | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const record = value as Record<string, unknown>;
  if (typeof record.id !== 'string' || !record.id) return undefined;
  if (typeof record.name !== 'string' || !record.name) return undefined;
  const paragraph = record.paragraph && typeof record.paragraph === 'object' ? (record.paragraph as ParagraphStyle) : {};
  const run = record.run && typeof record.run === 'object' ? (record.run as RunStyle) : {};
  return {
    id: record.id,
    name: record.name,
    builtinName: typeof record.builtinName === 'string' ? record.builtinName : undefined,
    basedOn: typeof record.basedOn === 'string' ? record.basedOn : undefined,
    paragraph,
    run,
    list:
      record.list && typeof record.list === 'object'
        ? {
            ordered: !!(record.list as { ordered?: boolean }).ordered,
            level: Number((record.list as { level?: number }).level) || 0,
          }
        : undefined,
    origin: 'custom',
  };
}

/** 功能：读取本地保存的样式集，损坏数据静默丢弃而不是抛出；时间：2026-10-29；作者：lq。 */
export function readStyleSets(storage: Storage | undefined = defaultStorage()): WordStyleSet[] {
  if (!storage) return [];
  try {
    const raw = storage.getItem(STYLE_SET_STORAGE_KEY);
    if (!raw) return [];
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed
      .map((item) => {
        if (!item || typeof item !== 'object') return undefined;
        const record = item as Record<string, unknown>;
        if (typeof record.id !== 'string' || typeof record.name !== 'string') return undefined;
        const styles = Array.isArray(record.styles)
          ? record.styles.map(sanitizeDefinition).filter((style): style is DocStyleDefinition => !!style)
          : [];
        return {
          id: record.id,
          name: record.name,
          styles,
          savedAt: typeof record.savedAt === 'string' ? record.savedAt : new Date().toISOString(),
        } satisfies WordStyleSet;
      })
      .filter((set): set is WordStyleSet => !!set)
      .slice(0, MAX_STYLE_SETS);
  } catch {
    return [];
  }
}

function writeStyleSets(sets: WordStyleSet[], storage: Storage | undefined): WordStyleSet[] {
  if (!storage) return sets;
  try {
    storage.setItem(STYLE_SET_STORAGE_KEY, JSON.stringify(sets.slice(0, MAX_STYLE_SETS)));
  } catch {
    // 存储配额满或被禁用时保持内存结果，不影响当前文档编辑。
  }
  return sets;
}

/** 功能：保存样式集；同名样式集覆盖而不是重复堆积；时间：2026-10-29；作者：lq。 */
export function addStyleSet(
  name: string,
  styles: DocStyleDefinition[],
  storage: Storage | undefined = defaultStorage(),
): WordStyleSet[] {
  const trimmed = name.trim();
  if (!trimmed || styles.length === 0) return readStyleSets(storage);
  const existing = readStyleSets(storage).filter((set) => set.name !== trimmed);
  const next: WordStyleSet = {
    id: `set_${Date.now().toString(36)}`,
    name: trimmed,
    styles: styles.map((style) => ({ ...style, origin: 'custom' as const })),
    savedAt: new Date().toISOString(),
  };
  return writeStyleSets([next, ...existing], storage);
}

/** 功能：删除样式集；时间：2026-10-29；作者：lq。 */
export function removeStyleSet(
  id: string,
  storage: Storage | undefined = defaultStorage(),
): WordStyleSet[] {
  return writeStyleSets(
    readStyleSets(storage).filter((set) => set.id !== id),
    storage,
  );
}

/**
 * 功能：把样式集导入文档，返回新的文档样式列表；时间：2026-10-29；作者：lq。
 * 同名样式集里的样式 ID 可能与文档已有样式冲突，冲突时重新生成 ID 而不是覆盖用户已有样式。
 */
export function importStyleSet(
  set: WordStyleSet,
  current: DocStyleDefinition[] | undefined,
): DocStyleDefinition[] {
  const existing = current ?? [];
  const taken = new Set(existing.map((style) => style.id));
  const additions: DocStyleDefinition[] = [];
  for (const style of set.styles) {
    if (isBuiltinStyle(style.id)) continue;
    const id = taken.has(style.id) ? newStyleId(taken) : style.id;
    taken.add(id);
    additions.push({ ...style, id, origin: 'custom' });
  }
  return [...existing, ...additions];
}

/** 功能：收集文档中实际使用到的样式定义，保存时按需写入 styles.xml；时间：2026-10-29；作者：lq。 */
export function collectUsedStyleDefinitions(
  doc: Pick<DocModel, 'blocks' | 'styles' | 'header' | 'footer'>,
): DocStyleDefinition[] {
  const used = new Set<string>();
  const visit = (blocks: DocModel['blocks']) => {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        if (block.styleId) used.add(block.styleId);
      } else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
      }
    }
  };
  visit(doc.blocks);
  if (doc.header) visit(doc.header.blocks);
  if (doc.footer) visit(doc.footer.blocks);
  const available = effectiveStyles(doc);
  return Array.from(used)
    .map((id) => available.find((style) => style.id === id))
    .filter((style): style is DocStyleDefinition => !!style);
}

/** 功能：判断两个 run 的字体格式是否等价，供样式识别与测试使用；时间：2026-10-29；作者：lq。 */
export function sameRunStyle(a: RunStyle | undefined, b: RunStyle | undefined): boolean {
  const left = a ?? {};
  const right = b ?? {};
  return (
    left.fontFamily === right.fontFamily &&
    (left.fontSizePt ?? 0) === (right.fontSizePt ?? 0) &&
    !!left.bold === !!right.bold &&
    !!left.italic === !!right.italic &&
    left.color === right.color
  );
}

/** 功能：段落的纯文本长度（不含图片与域），用于创建样式时判断是否有可取样文本；时间：2026-10-29；作者：lq。 */
export function paragraphCharCount(runs: DocRun[]): number {
  return runs.reduce((sum, run) => sum + (run.text?.length ?? 0), 0);
}
