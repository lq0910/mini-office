/**
 * 功能：逐字体自然行高实测——用隐藏探针元素按「同一字体栈 + 字号」测一次 line-height:normal
 * 的实际行盒高度（即渲染引擎为该字体使用的自然行高，来自字体真实度量），
 * 替换全局系数近似，使行距与 Word 的逐字体表现一致；时间：2026-10-31；作者：hou0。
 */

let probeEl: HTMLDivElement | null = null;
const cache = new Map<string, number>();

/**
 * 功能：测量字体栈在指定字号下的自然行高（CSS px）。返回 0 表示环境无法测量
 * （测试/SSR），调用方应回退到系数近似；时间：2026-10-31；作者：hou0。
 */
export function measureNaturalLineHeightPx(fontStack: string, fontSizePt: number): number {
  if (typeof document === 'undefined' || !fontStack) return 0;
  const key = `${fontStack}|${fontSizePt}`;
  const hit = cache.get(key);
  if (hit !== undefined) return hit;
  if (!probeEl) {
    probeEl = document.createElement('div');
    probeEl.setAttribute('aria-hidden', 'true');
    probeEl.style.cssText = 'position:absolute;left:-100000px;top:0;visibility:hidden;white-space:pre;line-height:normal;';
    document.body.appendChild(probeEl);
  }
  probeEl.style.fontFamily = fontStack;
  probeEl.style.fontSize = `${fontSizePt}pt`;
  // 功能：混排中英文与数字——行盒高度取行内实际用到字体的最大度量，与正文行为一致；时间：2026-10-31；作者：hou0。
  probeEl.textContent = '测Mxg1.5';
  const height = probeEl.getBoundingClientRect().height;
  const value = Number.isFinite(height) && height > 0 ? height : 0;
  cache.set(key, value);
  return value;
}

/** 功能：字体加载完成后清空缓存，让下一次测量拿到真实字体度量；时间：2026-10-31；作者：hou0。 */
export function clearFontMetricsCache(): void {
  cache.clear();
}
