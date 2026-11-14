/** 功能：原生核心快照的分页预览、打印和 PDF 共用同一实际版式，不使用旧模型/旧分页器；时间：2026-11-11；作者：lq。 */
import { WordCoreSurface } from './core-surface';
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import type { WordPageFrame } from './core-layout';
import '../../word-core/src/styles/document.css';
import '../../word-core/src/styles/fonts.css';

export interface CoreWordReport { pages: HTMLElement[]; frames: readonly WordPageFrame[]; surface: WordCoreSurface; captures: HTMLElement; style: HTMLStyleElement; dispose(): void; }
/** 功能：等待字体、图片和有界后续排版稳定，取消不会保留隐藏会话或输出页；时间：2026-11-11；作者：lq。 */
export async function mountCoreWordReport(input: Uint8Array | WordCoreSession, isCancelled: () => boolean = () => false): Promise<CoreWordReport> {
  const host = document.createElement('div'); host.dataset.coreOutputMeasure = 'true';
  host.style.cssText = 'position:fixed;left:-100000px;top:0;width:2000px;background:white;color:black;pointer-events:none;';
  document.body.append(host); let surface: WordCoreSurface | undefined; let disposed = false; let ownsSurface = false;
  const check = () => { if (isCancelled()) throw new Error('已取消输出。'); };
  const dispose = () => { if (disposed) return; disposed = true; if (ownsSurface) surface?.dispose(); host.remove(); };
  let releaseRendering: (() => void) | undefined;
  try {
    // 功能：复用正文或读取只读来源前先等待完整加载，预览/PDF 不得截取首屏半份文档；时间：2026-11-13；作者：lq。
    if (!(input instanceof Uint8Array)) await input.whenContentReady();
    // 功能：已有正文直接冻结当前排版，只为合并报告/未挂载字节创建独立只读宿主；时间：2026-11-11；作者：lq。
    check(); surface = input instanceof Uint8Array ? undefined : WordCoreSurface.current(input);
    // 功能：刚切换标签或隐藏正文尚无可用纸面时，采用独立只读测量，不捕获零宽度半成品；时间：2026-11-11；作者：lq。
    if (surface && (!surface.layout.current()?.frames.length || surface.session.editor.view.dom.getBoundingClientRect().width <= 0)) surface = undefined;
    if (!surface) { ownsSurface = true; surface = await WordCoreSurface.mount(host, input instanceof Uint8Array ? input : await coreWordOutputBytes(input), { readonly: true }); }
    await surface.session.whenContentReady();
    // 功能：捕获页面全过程使用完整正文，不能把视窗外占位带入打印和 PDF；时间：2026-11-14；作者：lq。
    releaseRendering = surface.acquireFullRendering();
    check(); const revision = surface.session.state().revision;
    if (ownsSurface) {
    let timeout: ReturnType<typeof setTimeout> | undefined;
    try { await Promise.race([
      // 功能：Safari 的 ProseMirror 光标分隔图片没有 src，不属于文档媒体，不能交给 decode 阻止打印/PDF；时间：2026-11-11；作者：lq。
      Promise.all([document.fonts?.ready, ...Array.from(surface.session.editor.view.dom.querySelectorAll<HTMLImageElement>('img:not(.ProseMirror-separator)')).map(image => image.decode())]),
      new Promise((_, reject) => { timeout = setTimeout(() => reject(new Error('字体或图片加载超时，请检查文档资源。')), 15000); }),
    ]); } finally { clearTimeout(timeout); }
    let signature = '', stable = 0;
    for (let pass = 0; pass < 12; pass++) {
      check(); await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
      if (surface.layoutError) throw surface.layoutError;
      // 功能：只读重新挂载需要等待装饰收敛；已完成的屏幕排版只刷新一次，不循环整篇重排；时间：2026-11-11；作者：lq。
      if (ownsSurface || pass === 0) surface.layout.flush(); const state = surface.layout.current();
      const next = JSON.stringify(state?.frames.map((frame, index) => [frame, state.cuts[index].start, state.cuts[index].end]));
      stable = next === signature ? stable + 1 : 0; signature = next; if (stable >= 2) break;
    }
    check(); if (stable < 2 || surface.layoutError) throw new Error('文档版式尚未稳定，已停止输出。');
    } else {
      // 功能：当前正文已经完成资源加载，直接同步最终一轮并冻结几何，不再全局等待字体或重开正文；时间：2026-11-11；作者：lq。
      surface.layout.flush(); if (surface.layoutError) throw surface.layoutError;
    }
    const state = surface.layout.current(); if (!state?.frames.length) throw new Error('没有已完成排版的页面。');
    if (state.frames.length > 500) throw new Error('输出超过 500 页，请缩小范围。');
    const pages = await cloneCoreWordPages(surface, check); check();
    if (surface.session.state().revision !== revision) throw new Error('准备预览期间文档已更改，请重新打开预览。');
    // 功能：预览翻页不会反复移除/插入样式表而使整篇正文重新计算样式；每份报告仅挂一张共享样式表；时间：2026-11-11；作者：lq。
    const style = pages[0].querySelector('style')!; host.append(style);
    for (const page of pages) page.querySelectorAll('style').forEach(element => element.remove());
    const captures = document.createElement('div'); captures.dataset.coreOutputCaptures = 'true'; captures.append(...pages); host.append(captures);
    return { pages, frames: state.frames, surface, captures, style, dispose };
  } catch (error) { dispose(); throw error; }
  finally { releaseRendering?.(); }
}
/** 功能：克隆同源自然流和已提交装饰，按真实纸张裁出每页；页间空白、重复表头和跨页行不再重新推算；时间：2026-11-11；作者：lq。 */
export async function cloneCoreWordPages(surface: WordCoreSurface, check: () => void = () => {}): Promise<HTMLElement[]> {
  const state = surface.layout.current(), source = surface.root.querySelector<HTMLElement>('.page-wrap');
  if (!state?.frames.length || !source) throw new Error('Word 页面尚未排版。');
  // 功能：一次采集屏幕几何和样式；每页仅复制与该页相交的顶层块，避免页数×全文 DOM 的膨胀；时间：2026-11-11；作者：lq。
  const paper = surface.session.editor.view.dom, paperBox = paper.getBoundingClientRect();
  const factor = paperBox.width / paper.offsetWidth || 1;
  const blocks = Array.from(paper.children, element => {
    const box = element.getBoundingClientRect();
    return { element: element as HTMLElement, top: (box.top - paperBox.top) / factor, left: (box.left - paperBox.left) / factor, width: box.width / factor, height: box.height / factor };
  });
  const regions = Array.from(source.querySelectorAll<HTMLElement>('.mo-core-page-regions > *'));
  const css = surface.session.styleCss().replace(/\.doc-page\b/g, '.mo-core-output-page .doc-page');
  const pages: HTMLElement[] = [];
  for (const [index, frame] of state.frames.entries()) {
    check();
    const page = document.createElement('div'); page.className = 'mo-core-output-page mo-word-core'; page.dataset.coreOutputPage = String(index + 1);
    page.style.cssText = `position:relative;width:${frame.widthPx}px;height:${frame.heightPx}px;overflow:hidden;background:white;color:black;box-sizing:border-box;`;
    const style = document.createElement('style'); style.textContent = css + '\n.mo-core-output-page .doc-page{box-shadow:none!important;border-radius:0!important;outline:none!important}.mo-core-output-page .page-wrap{background:transparent!important;box-shadow:none!important;--docs-paper-crop-mark:transparent}.mo-core-output-page .doc-ref-field{background:transparent!important}.mo-core-output-page .page-break-before::before,.mo-core-output-page .page-break-before::after{visibility:hidden!important}.mo-core-output-page .ProseMirror-selectednode{outline:none!important}.mo-core-output-page .selectedCell::after,.mo-core-output-page .page-gap::after,.mo-core-output-page .page-wrap::before,.mo-core-output-page .page-wrap::after{display:none!important}.mo-core-output-page .doc-page *{break-before:auto!important;break-after:auto!important;break-inside:auto!important}';
    const scroll = document.createElement('div'); scroll.className = 'editor-scroll'; scroll.style.cssText = 'position:absolute;inset:0;padding:0;overflow:visible;background:transparent;';
    const clone = source.cloneNode(false) as HTMLElement;
    const body = paper.cloneNode(false) as HTMLElement;
    body.style.position = 'relative'; body.style.width = `${paperBox.width / factor}px`; body.style.height = `${paperBox.height / factor}px`;
    for (const block of blocks) {
      if (block.top + block.height < frame.top - 1 || block.top > frame.top + frame.heightPx + 1 || block.height <= 0) continue;
      const child = block.element.cloneNode(true) as HTMLElement;
      child.style.setProperty('position', 'absolute', 'important'); child.style.setProperty('top', `${block.top}px`, 'important');
      child.style.setProperty('left', `${block.left}px`, 'important'); child.style.setProperty('width', `${block.width}px`, 'important');
      child.style.setProperty('height', `${block.height}px`, 'important'); child.style.setProperty('margin', '0', 'important'); child.style.transform = 'none';
      child.style.boxSizing = 'border-box'; body.append(child);
      // 功能：只转换当前块里的绘图画布，避免每一页扫描和编码全文图表；时间：2026-11-11；作者：lq。
      const canvases = block.element.querySelectorAll('canvas'); child.querySelectorAll('canvas').forEach((canvas, canvasIndex) => {
        const image = document.createElement('img'); image.src = canvases[canvasIndex].toDataURL('image/png'); image.style.cssText = canvas.style.cssText;
        image.width = canvases[canvasIndex].width; image.height = canvases[canvasIndex].height; canvas.replaceWith(image);
      });
    }
    clone.append(body);
    const regionLayer = document.createElement('div'); regionLayer.className = 'mo-core-page-regions'; regionLayer.style.cssText = 'position:absolute;inset:0;pointer-events:none';
    for (const region of regions) {
      const regionPage = Number(region.dataset.hfPage);
      const top = parseFloat(region.style.top);
      if (region.dataset.hfPage !== undefined ? regionPage === index : top >= frame.top && top < frame.top + frame.heightPx) regionLayer.append(region.cloneNode(true));
    }
    clone.append(regionLayer);
    // 功能：保存屏幕已提交的分页几何，打印媒体规则不能隐藏页间占位或把不同节正文改成同一宽度；时间：2026-11-11；作者：lq。
    const geometrySelector = '.doc-page,.page-gap,.page-float-host,.doc-field-pagebreak';
    const originals = [paper, ...blocks.filter(block => block.top + block.height >= frame.top - 1 && block.top <= frame.top + frame.heightPx + 1 && block.height > 0).flatMap(block => [block.element, ...block.element.querySelectorAll<HTMLElement>(geometrySelector)])].filter(el => el.matches(geometrySelector));
    clone.querySelectorAll<HTMLElement>(geometrySelector).forEach((element, index) => {
      const original = originals[index]; if (!original) return;
      const computed = getComputedStyle(original);
      for (const property of ['display', 'width', 'height', 'min-height', 'margin-top', 'margin-bottom', 'margin-left', 'margin-right']) element.style.setProperty(property, computed.getPropertyValue(property), 'important');
    });
    clone.style.position = 'absolute'; clone.style.top = `${-frame.top}px`; clone.style.left = `${-frame.left}px`; clone.style.margin = '0'; clone.style.zoom = '1'; clone.style.transform = 'none';
    clone.querySelectorAll<HTMLElement>('[contenteditable]').forEach(element => { element.setAttribute('contenteditable', 'false'); element.removeAttribute('tabindex'); });
    // 功能：编辑器专用的 Safari 光标分隔节点不进入输出页面；时间：2026-11-11；作者：lq。
    clone.querySelectorAll('.mo-no-print,.doc-table-handle,.doc-table-resize,.doc-image-resize-handle,.doc-resize-handle,.img-resize-handle,.doc-image-anchor-marker,.doc-move-handle,.doc-protected-label,.ProseMirror-gapcursor,.ProseMirror-separator,.page-hf-tags,.search-hit,.search-hit-active').forEach(element => {
      if (element.classList.contains('search-hit') || element.classList.contains('search-hit-active')) element.classList.remove('search-hit', 'search-hit-active'); else element.remove();
    });
    scroll.append(clone); page.append(style, scroll); pages.push(page);
    // 功能：大报告分批制作页面，把事件循环交还滚轮/取消按钮；翻页直接移动已缓存的轻量页面；时间：2026-11-11；作者：lq。
    if (index % 4 === 3) await new Promise<void>(resolve => setTimeout(resolve, 0));
  }
  return pages;
}
/** 功能：输出捕获当前不可变版本，只读文档可预览/导出原始内容，保存确认与 dirty 状态完全独立；时间：2026-11-11；作者：lq。 */
export async function coreWordOutputBytes(core: WordCoreSession): Promise<Uint8Array> {
  if (core.editor.view.composing) throw new Error('中文输入尚未确认，请结束输入后输出。');
  return core.editor.isEditable ? (await core.prepareSave()).bytes : core.parsed.internal.originalBytes.slice();
}
