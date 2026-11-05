/**
 * 功能：字体单行高度度量与 Word 行距换算，供段落渲染与分页共用；时间：2026-11-02；作者：lq。
 *
 * 为什么需要它：
 * 1. Word 的「1.5 倍行距」是「单行高度 × 1.5」，而单行高度由字体自身的 ascent + descent 决定，
 *    通常约为字号的 1.15–1.35 倍；CSS 的 `line-height: 1.5` 却是「字号 × 1.5」。
 *    两者相差 15%–35%，行数一多就会累积成整页偏差。
 * 2. 因此这里统一用 Canvas 的 `fontBoundingBoxAscent/Descent` 取真实度量，
 *    再按 Word 规则换算成具体像素值，渲染时不再使用无单位倍数。
 * 3. 度量结果按「字体栈 + 字号」缓存；字体加载完成后调用 `resetFontMetrics()` 重新测量。
 */
import { fontStackFor } from '@mini-office/design-system';
import { PX_PER_PT } from './typography';
import type { ParagraphLayout } from './pagination';
// 功能：使用实际 run 字体、字号与备用名计算段落行高；时间：2026-11-05；作者：lq。
import type { ParagraphStyle, RunStyle } from '@mini-office/document-model';
// 功能：常用中文字体按 Word 的设计行高布局，字体字形仍由本机实际字体绘制；时间：2026-11-05；作者：lq。
import { wordFontLineFactor } from './word-font-metrics';

/** 无法测量时的兜底单行高度系数：中文常用字体约 1.30–1.35，西文约 1.15。 */
const CJK_FALLBACK_RATIO = 1.32;
const LATIN_FALLBACK_RATIO = 1.15;
// 功能：放大字体度量后换算回实际字号，减少 Canvas 在小字号下将升降部取整引起的逐行累积误差；时间：2026-11-05；作者：lq。
const METRIC_FONT_SIZE_PX = 4096;

const CJK_FONT_HINTS = [
  '宋体', '黑体', '仿宋', '楷体', '微软雅黑', '等线', '思源', '方正', '华文',
  'songti', 'heiti', 'fangsong', 'kaiti', 'yahei', 'pingfang', 'noto sans cjk', 'noto serif cjk',
  'simsun', 'simhei', 'dengxian', 'mingliu', 'ms song',
];

const cache = new Map<string, number>();
let canvas: HTMLCanvasElement | null | undefined;

/** 功能：判断字体族是否为中文字体，决定兜底系数的取值；时间：2026-11-02；作者：lq。 */
export function isCjkFont(family: string | undefined): boolean {
  if (!family) return true;
  const value = family.toLowerCase();
  return CJK_FONT_HINTS.some((hint) => value.includes(hint.toLowerCase()));
}

function measureContext(): CanvasRenderingContext2D | null {
  if (canvas !== undefined) return canvas?.getContext('2d') ?? null;
  if (typeof document === 'undefined') {
    canvas = null;
    return null;
  }
  // 功能：jsdom 未实现 getContext，直接探测会刷出无意义的错误日志；测试环境统一走兜底系数；时间：2026-11-02；作者：lq。
  if (typeof navigator !== 'undefined' && /jsdom/i.test(navigator.userAgent)) {
    canvas = null;
    return null;
  }
  canvas = document.createElement('canvas');
  try {
    return canvas.getContext('2d');
  } catch {
    canvas = null;
    return null;
  }
}

/**
 * 功能：返回某字体在指定字号下的单行高度（CSS px）；时间：2026-11-02；作者：lq。
 *
 * 优先取 Canvas 实测的 fontBoundingBox；WebView 不支持时按中英文系数兜底，
 * 保证在没有测量的环境（如单元测试）里也能得到稳定、可比较的结果。
 */
export function singleLineHeightPx(family: string | undefined, fontSizePt: number, alternateFamily?: string): number {
  const sizePt = Number.isFinite(fontSizePt) && fontSizePt > 0 ? fontSizePt : 10.5;
  const stack = fontStackFor(family, alternateFamily) || 'sans-serif';
  const key = `${stack}|${sizePt}`;
  const cached = cache.get(key);
  if (cached !== undefined) return cached;

  const sizePx = sizePt * PX_PER_PT;
  // 功能：Word 中文单倍行高含额外行距，直接使用 Canvas 字形边界会让大字号和网格段落少占一格；时间：2026-11-05；作者：lq。
  const wordFactor = wordFontLineFactor(family, alternateFamily);
  if (wordFactor !== undefined) {
    const height = sizePx * wordFactor;
    cache.set(key, height);
    return height;
  }
  let height = 0;
  const ctx = measureContext();
  if (ctx) {
    try {
      // 功能：只测量字体边界，不绘制大图；缓存仍按实际字体和字号隔离，字体就绪后可统一失效；时间：2026-11-05；作者：lq。
      ctx.font = `${METRIC_FONT_SIZE_PX}px ${stack}`;
      const metrics = ctx.measureText('汉字Hg');
      const ascent = metrics.fontBoundingBoxAscent;
      const descent = metrics.fontBoundingBoxDescent;
      if (Number.isFinite(ascent) && Number.isFinite(descent) && ascent + descent > 0) {
        height = (ascent + descent) * sizePx / METRIC_FONT_SIZE_PX;
      }
    } catch {
      height = 0;
    }
  }
  if (!height) {
    height = sizePx * (isCjkFont(family) ? CJK_FALLBACK_RATIO : LATIN_FALLBACK_RATIO);
  }
  cache.set(key, height);
  return height;
}

/**
 * 功能：按 Word 规则把段落行距换算成具体像素值；时间：2026-11-02；作者：lq。
 *
 * - 启用文档网格时，「单倍」取网格行距（w:docGrid/@linePitch），而不是字体度量；
 * - 固定值（exact）：直接使用设定值；
 * - 最小值（atLeast）：取「设定值」与「单行高度」的较大者；
 * - 倍数（auto）：单行高度 × 倍数（Word 的 240 分之一行）。
 */
export function lineHeightPx(
  style: { lineHeight?: number; lineHeightExactPx?: number; lineHeightAtLeastPx?: number } | undefined,
  family: string | undefined,
  fontSizePt: number,
  gridLinePitchPx?: number,
  // 功能：度量与实际 CSS 使用同一份文档备用字体信息；时间：2026-11-05；作者：lq。
  alternateFamily?: string,
): number {
  if (style?.lineHeightExactPx) return style.lineHeightExactPx;
  // 功能：依据 GenOffice 的 Word 网格探针，单行高度向上吸附网格，大字号可占多个网格；时间：2026-11-03；作者：lq。
  const natural = singleLineHeightPx(family, fontSizePt, alternateFamily);
  const pitch = gridLinePitchPx && gridLinePitchPx > 0 ? gridLinePitchPx : 0;
  const single = pitch ? Math.ceil((natural - 0.001) / pitch) * pitch : natural;
  // 功能：最小值为 0 时不吸附网格；其他最小值不得低于网格吸附后的自然行高；时间：2026-11-03；作者：lq。
  if (style?.lineHeightAtLeastPx !== undefined) return style.lineHeightAtLeastPx === 0 ? natural : Math.max(style.lineHeightAtLeastPx, single);
  const multiple = typeof style?.lineHeight === 'number' && style.lineHeight > 0 ? style.lineHeight : 1;
  return pitch ? Math.max(pitch * multiple, single) : single * multiple;
}

/** 功能：逐 run 度量段落，最大字号与最高字体可来自不同 run；文档默认值只补缺失格式，不再覆盖显式的小字号；时间：2026-11-05；作者：lq。 */
export function paragraphLineMetrics(
  runs: ReadonlyArray<{ style?: RunStyle }>, style: ParagraphStyle,
  defaultFont: string, defaultFontSizePt: number, gridLinePitchPx?: number,
): { family: string; sizePt: number; alternateFamily?: string; heightPx: number; naturalHeightPx: number } {
  const candidates = runs.length ? runs : [{ style: {} }];
  let selected: ReturnType<typeof paragraphLineMetrics> | undefined;
  let heightPx = 0;
  for (const run of candidates) {
    const family = run.style?.fontFamily ?? defaultFont;
    const value = run.style?.fontSizePt;
    const sizePt = value && Number.isFinite(value) && value > 0 ? value : defaultFontSizePt;
    const alternateFamily = fallbackFontFamily(run.style);
    const naturalHeightPx = singleLineHeightPx(family, sizePt, alternateFamily);
    heightPx = Math.max(heightPx, lineHeightPx(style, family, sizePt, gridLinePitchPx, alternateFamily));
    if (!selected || naturalHeightPx > selected.naturalHeightPx) {
      selected = { family, sizePt, alternateFamily, heightPx, naturalHeightPx };
    }
  }
  return { ...selected!, heightPx };
}

/** 功能：只在主字体仍与导入时相同时使用该字体的备用名，防止编辑字体后残留旧映射；时间：2026-11-05；作者：lq。 */
export function fallbackFontFamily(style?: RunStyle): string | undefined {
  return style?.fontFallbackFor === style?.fontFamily ? style?.fontFallback : undefined;
}

/**
 * 功能：段落内用于行高的字号——取各 run 的最大字号，与 Word「按行内最大字号定行高」一致；时间：2026-11-02；作者：lq。
 */
export function paragraphFontSizePt(
  runs: ReadonlyArray<{ style?: { fontSizePt?: number } }> | undefined,
  fallbackPt: number,
): number {
  let max = fallbackPt;
  if (!runs) return max;
  for (const run of runs) {
    const size = run.style?.fontSizePt;
    if (typeof size === 'number' && size > max) max = size;
  }
  return max;
}

/** 功能：字体加载或系统字体变化后清空度量缓存；时间：2026-11-02；作者：lq。 */
export function resetFontMetrics(): void {
  cache.clear();
}

/**
 * 功能：读取段落的行盒信息，供分页按行切分；时间：2026-11-03；作者：lq。
 *
 * 判定依据：段落渲染时已把行高写成确定像素值（见 `paragraphTextStyle`），
 * 因此「实测高度 ÷ 行高」就是行数。若两者不能整除，说明行内有更高的内联内容
 * （大图、独立行高的内联块），行盒不再均匀，此时返回 undefined 让分页回退到整段移动，
 * 宁可少切一刀，也不把行切错。
 */
export function measureParagraphLines(element: HTMLElement | null | undefined): ParagraphLayout | undefined {
  if (!element || typeof getComputedStyle !== 'function') return undefined;
  let lineHeight = 0;
  let height = 0;
  try {
    lineHeight = Number.parseFloat(getComputedStyle(element).lineHeight);
    height = element.getBoundingClientRect().height;
  } catch {
    return undefined;
  }
  if (!Number.isFinite(lineHeight) || lineHeight <= 0) return undefined;
  if (!Number.isFinite(height) || height <= 0) return undefined;
  const count = Math.round(height / lineHeight);
  // 单行段落无需切分；不足一行时也不切。
  if (count < 2) return undefined;
  if (Math.abs(height - count * lineHeight) > 1.5) return undefined;
  return {
    lineTops: Array.from({ length: count }, (_, index) => index * lineHeight),
    height: count * lineHeight,
  };
}
