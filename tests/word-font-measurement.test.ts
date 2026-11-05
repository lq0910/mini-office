/** 功能：模拟浏览器字体边界取整，回归小字号行高精度、倍数行距及字体加载后的缓存失效；时间：2026-11-05；作者：lq。 */
import { afterEach, describe, expect, it, vi } from 'vitest';
// 功能：回归只有行网格参与段落行高，字符网格属性仍可独立保留；时间：2026-11-05；作者：lq。
import { documentGridLinePitch } from '../packages/editor-word/src/metrics';

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

/** 功能：保留 Canvas 的整像素边界特征，使用独立字体比例构造可复现的浏览器度量；时间：2026-11-05；作者：lq。 */
async function browserMetrics() {
  vi.resetModules();
  vi.stubGlobal('navigator', { userAgent: 'Chromium' });
  let ascentRatio = 0.952;
  const context = {
    font: '',
    measureText: vi.fn(() => {
      const size = Number.parseFloat(context.font);
      return { fontBoundingBoxAscent: Math.round(size * ascentRatio), fontBoundingBoxDescent: Math.round(size * 0.269) };
    }),
  };
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(context as unknown as CanvasRenderingContext2D);
  const metrics = await import('../packages/editor-word/src/metrics');
  return { ...metrics, context, changeFont: () => { ascentRatio = 1.1; } };
}

describe('浏览器字体边界精度', () => {
  // 功能：snapToChars 不得把可能存在的 linePitch 当作正文垂直网格；时间：2026-11-05；作者：lq。
  it('仅 lines 与 linesAndChars 提供垂直行距', () => {
    expect(documentGridLinePitch({ type: 'lines', linePitchPx: 24 })).toBe(24);
    expect(documentGridLinePitch({ type: 'linesAndChars', linePitchPx: 24 })).toBe(24);
    expect(documentGridLinePitch({ type: 'snapToChars', linePitchPx: 24 })).toBeUndefined();
    expect(documentGridLinePitch({ type: 'lines' })).toBeUndefined();
  });

  it('11pt 的单倍和倍数行距不会把取整误差累计到每一行', async () => {
    const metrics = await browserMetrics();
    const expected = 11 * 4 / 3 * (0.952 + 0.269);
    const actual = metrics.singleLineHeightPx('Calibri', 11);
    expect(Math.abs(actual - expected)).toBeLessThan(0.01);
    expect(Math.abs(metrics.lineHeightPx({ lineHeight: 1.15 }, 'Calibri', 11) - expected * 1.15)).toBeLessThan(0.012);
    // 功能：一百行的误差保持在一个 CSS 像素以内，旧版 18px 取整会累积约九像素；时间：2026-11-05；作者：lq。
    expect(Math.abs(actual - expected) * 100).toBeLessThan(1);
  });

  it('字体缓存可在字体加载完成后重测，其他字号按相同比例换算', async () => {
    const metrics = await browserMetrics();
    const initial = metrics.singleLineHeightPx('Calibri', 11);
    expect(metrics.singleLineHeightPx('Calibri', 22)).toBeCloseTo(initial * 2, 8);
    metrics.changeFont();
    expect(metrics.singleLineHeightPx('Calibri', 11)).toBe(initial);
    metrics.resetFontMetrics();
    expect(metrics.singleLineHeightPx('Calibri', 11)).toBeGreaterThan(initial);
  });

  it('浏览器不提供字体边界时仍使用兜底，固定行距不受测量影响', async () => {
    const metrics = await browserMetrics();
    metrics.context.measureText.mockReturnValue({ fontBoundingBoxAscent: Number.NaN, fontBoundingBoxDescent: Number.NaN });
    expect(metrics.singleLineHeightPx('Calibri', 12)).toBeCloseTo(16 * 1.15, 5);
    expect(metrics.lineHeightPx({ lineHeightExactPx: 20 }, 'Calibri', 12)).toBe(20);
  });
});
