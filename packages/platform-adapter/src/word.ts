/** 功能：Word 桌面字体与位图剪贴板原生接口；时间：2026-11-11；作者：lq。 */
import { invoke } from '@tauri-apps/api/core';

/** 功能：与系统字体原始表对应的度量结构，避免平台层依赖 Word 核心；时间：2026-11-11；作者：lq。 */
export interface SystemWordFontMetrics {
  unitsPerEm: number; hheaAscender: number; hheaDescender: number; hheaLineGap: number;
  typoAscender: number | null; typoDescender: number | null; typoLineGap: number | null;
  winAscent: number | null; winDescent: number | null; useTypoMetrics: boolean;
  xAvgCharWidth: number | null; codePageRange1: number | null;
}
/** 功能：按字体名称读取系统索引的度量，不接受任意路径；时间：2026-11-11；作者：lq。 */
export const coreWordFontMetrics = (family: string) => invoke<SystemWordFontMetrics | null>('word_font_metrics', { family });

/** 功能：使用真实解码结果生成 PNG 和可往返的格式 HTML，异常转为明确失败，剪切不会误删；时间：2026-11-11；作者：lq。 */
export async function copyCoreWordImage(dataUrl: string, metadata: string): Promise<boolean> {
  try {
    const image = new Image();
    image.src = dataUrl;
    await image.decode();
    if (!image.naturalWidth || !image.naturalHeight || image.naturalWidth * image.naturalHeight > 32 * 1024 * 1024) return false;
    const canvas = document.createElement('canvas');
    canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
    const context = canvas.getContext('2d');
    if (!context) return false;
    context.drawImage(image, 0, 0);
    const pngUrl = canvas.toDataURL('image/png');
    const img = document.createElement('img');
    img.src = pngUrl; img.setAttribute('data-image-meta', metadata);
    const info = JSON.parse(metadata);
    if (Number.isFinite(info.imageWidthPx) && info.imageWidthPx > 0) img.width = info.imageWidthPx;
    if (Number.isFinite(info.imageHeightPx) && info.imageHeightPx > 0) img.height = info.imageHeightPx;
    const png = Array.from(atob(pngUrl.slice(pngUrl.indexOf(',') + 1)), char => char.charCodeAt(0));
    return await invoke<boolean>('word_copy_image', { png, html: img.outerHTML });
  } catch { return false; }
}

