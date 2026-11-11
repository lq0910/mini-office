/** 功能：原生核心快照的分页预览、打印和 PDF 共用同一实际版式，不使用旧模型/旧分页器；时间：2026-11-11；作者：lq。 */
import { WordCoreSurface } from './core-surface';
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import type { WordPageFrame } from './core-layout';
import '../../word-core/src/styles/document.css';
import '../../word-core/src/styles/fonts.css';

export interface CoreWordReport { pages: HTMLElement[]; frames: readonly WordPageFrame[]; surface: WordCoreSurface; dispose(): void; }
/** 功能：等待字体、图片和有界后续排版稳定，取消不会保留隐藏会话或输出页；时间：2026-11-11；作者：lq。 */
export async function mountCoreWordReport(bytes: Uint8Array, isCancelled: () => boolean = () => false): Promise<CoreWordReport> {
  const host = document.createElement('div'); host.dataset.coreOutputMeasure = 'true';
  host.style.cssText = 'position:fixed;left:-100000px;top:0;width:2000px;background:white;color:black;pointer-events:none;';
  document.body.append(host); let surface: WordCoreSurface | undefined; let disposed = false;
  const check = () => { if (isCancelled()) throw new Error('已取消输出。'); };
  const dispose = () => { if (disposed) return; disposed = true; surface?.dispose(); host.remove(); };
  try {
    check(); surface = await WordCoreSurface.mount(host, bytes, { readonly: true }); check();
    let timeout: ReturnType<typeof setTimeout> | undefined;
    try { await Promise.race([
      Promise.all([document.fonts?.ready, ...Array.from(surface.session.editor.view.dom.querySelectorAll('img')).map(image => image.decode())]),
      new Promise((_, reject) => { timeout = setTimeout(() => reject(new Error('字体或图片加载超时，请检查文档资源。')), 15000); }),
    ]); } finally { clearTimeout(timeout); }
    let signature = '', stable = 0;
    for (let pass = 0; pass < 12; pass++) {
      check(); await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
      if (surface.layoutError) throw surface.layoutError;
      surface.layout.flush(); const state = surface.layout.current();
      const next = JSON.stringify(state?.frames.map((frame, index) => [frame, state.cuts[index].start, state.cuts[index].end]));
      stable = next === signature ? stable + 1 : 0; signature = next; if (stable >= 2) break;
    }
    check(); if (stable < 2 || surface.layoutError) throw new Error('文档版式尚未稳定，已停止输出。');
    const state = surface.layout.current(); if (!state?.frames.length) throw new Error('没有已完成排版的页面。');
    if (state.frames.length > 500) throw new Error('输出超过 500 页，请缩小范围。');
    const pages = cloneCoreWordPages(surface); const captures = document.createElement('div'); captures.dataset.coreOutputCaptures = 'true'; captures.append(...pages); host.append(captures);
    return { pages, frames: state.frames, surface, dispose };
  } catch (error) { dispose(); throw error; }
}
/** 功能：克隆同源自然流和已提交装饰，按真实纸张裁出每页；页间空白、重复表头和跨页行不再重新推算；时间：2026-11-11；作者：lq。 */
export function cloneCoreWordPages(surface: WordCoreSurface): HTMLElement[] {
  const state = surface.layout.current(), source = surface.root.querySelector<HTMLElement>('.page-wrap');
  if (!state?.frames.length || !source) throw new Error('Word 页面尚未排版。');
  return state.frames.map((frame, index) => {
    const page = document.createElement('div'); page.className = 'mo-core-output-page mo-word-core'; page.dataset.coreOutputPage = String(index + 1);
    page.style.cssText = `position:relative;width:${frame.widthPx}px;height:${frame.heightPx}px;overflow:hidden;background:white;color:black;box-sizing:border-box;`;
    const style = document.createElement('style'); style.textContent = surface.session.styleCss() + '\n.mo-core-output-page .doc-page{box-shadow:none!important;border-radius:0!important;outline:none!important}.mo-core-output-page .page-wrap{background:transparent!important;box-shadow:none!important;--docs-paper-crop-mark:transparent}.mo-core-output-page .doc-ref-field{background:transparent!important}.mo-core-output-page .page-break-before::before,.mo-core-output-page .page-break-before::after{visibility:hidden!important}.mo-core-output-page .ProseMirror-selectednode{outline:none!important}.mo-core-output-page .selectedCell::after,.mo-core-output-page .page-gap::after,.mo-core-output-page .page-wrap::before,.mo-core-output-page .page-wrap::after{display:none!important}.mo-core-output-page .doc-page *{break-before:auto!important;break-after:auto!important;break-inside:auto!important}';
    const scroll = document.createElement('div'); scroll.className = 'editor-scroll'; scroll.style.cssText = 'position:absolute;inset:0;padding:0;overflow:visible;background:transparent;';
    const clone = source.cloneNode(true) as HTMLElement;
    // 功能：保存屏幕已提交的分页几何，打印媒体规则不能隐藏页间占位或把不同节正文改成同一宽度；时间：2026-11-11；作者：lq。
    const geometrySelector = '.doc-page,.page-gap,.page-float-host,.doc-field-pagebreak';
    const originals = source.querySelectorAll<HTMLElement>(geometrySelector);
    clone.querySelectorAll<HTMLElement>(geometrySelector).forEach((element, index) => {
      const computed = getComputedStyle(originals[index]);
      for (const property of ['display', 'width', 'height', 'min-height', 'margin-top', 'margin-bottom', 'margin-left', 'margin-right']) element.style.setProperty(property, computed.getPropertyValue(property), 'important');
    });
    const canvases = source.querySelectorAll('canvas'); clone.querySelectorAll('canvas').forEach((canvas, index) => {
      const image = document.createElement('img'); image.src = canvases[index].toDataURL('image/png'); image.style.cssText = canvas.style.cssText; image.width = canvases[index].width; image.height = canvases[index].height; canvas.replaceWith(image);
    });
    clone.style.position = 'absolute'; clone.style.top = `${-frame.top}px`; clone.style.left = `${-frame.left}px`; clone.style.margin = '0'; clone.style.zoom = '1'; clone.style.transform = 'none';
    clone.querySelectorAll<HTMLElement>('[contenteditable]').forEach(element => { element.setAttribute('contenteditable', 'false'); element.removeAttribute('tabindex'); });
    clone.querySelectorAll('.mo-no-print,.doc-table-handle,.doc-table-resize,.doc-image-resize-handle,.doc-resize-handle,.img-resize-handle,.doc-image-anchor-marker,.doc-move-handle,.doc-protected-label,.ProseMirror-gapcursor,.page-hf-tags,.search-hit,.search-hit-active').forEach(element => {
      if (element.classList.contains('search-hit') || element.classList.contains('search-hit-active')) element.classList.remove('search-hit', 'search-hit-active'); else element.remove();
    });
    scroll.append(clone); page.append(style, scroll); return page;
  });
}
/** 功能：输出捕获当前不可变版本，只读文档可预览/导出原始内容，保存确认与 dirty 状态完全独立；时间：2026-11-11；作者：lq。 */
export async function coreWordOutputBytes(core: WordCoreSession): Promise<Uint8Array> {
  if (core.editor.view.composing) throw new Error('中文输入尚未确认，请结束输入后输出。');
  return core.editor.isEditable ? (await core.prepareSave()).bytes : core.parsed.internal.originalBytes.slice();
}
