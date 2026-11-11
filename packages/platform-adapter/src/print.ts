import { invoke } from '@tauri-apps/api/core';
import { isTauriRuntime } from './types';

/**
 * 功能：统一 Word / Excel 的打印分页基础设施，避免各编辑器各写一套 @media print 造成页数丢失；
 * 时间：2026-10-29；作者：lq。
 *
 * 关键约束（都对应线上真实故障）：
 * 1. 打印时必须把 html/body 的 height:100% 与 overflow 复位，否则宿主窗口高度会把打印文档裁成一两页。
 * 2. 打印根节点必须参与普通流（不能 position:absolute），否则 WebKit 只按首屏高度分页。
 * 3. 页面盒尺寸必须与 @page size 使用同一物理长度，任何 0.x px 溢出都会让每页后面多出一张白纸。
 * 4. 最后一页不能有 break-after，否则末尾固定多一张空白页。
 */

/** CSS 像素与毫米的换算基准：96dpi。 */
const PX_PER_MM = 96 / 25.4;

/** 功能：常见纸张尺寸（毫米），用于把像素页面吸附到标准纸张，规避浮点溢出多出的空白页；时间：2026-10-29；作者：lq。 */
const STANDARD_PAPERS: Array<{ name: string; widthMm: number; heightMm: number }> = [
  { name: 'A3', widthMm: 297, heightMm: 420 },
  { name: 'A4', widthMm: 210, heightMm: 297 },
  { name: 'A5', widthMm: 148, heightMm: 210 },
  { name: 'B5', widthMm: 176, heightMm: 250 },
  { name: 'Letter', widthMm: 215.9, heightMm: 279.4 },
  { name: 'Legal', widthMm: 215.9, heightMm: 355.6 },
];

export interface PrintPageGeometry {
  widthPx: number;
  heightPx: number;
}

export interface ResolvedPaper {
  /** 形如 `size: 210mm 297mm;`。 */
  pageRule: string;
  /** 页面盒宽度，与 pageRule 是同一物理长度。 */
  width: string;
  height: string;
  /** 命中的标准纸张名称，未命中为标准外的自定义尺寸。 */
  name?: string;
  /** 纸张物理宽度（毫米），用于在打印时把布局宽度锁到纸张宽度。 */
  widthMm: number;
  /** 纸张物理高度（毫米）。 */
  heightMm: number;
  /**
   * 打印页面盒的实际高度：比纸张矮一点点，用来吸收浏览器“按可打印宽度等比缩放”带来的纵向溢出。
   * 打印机可打印区域的纵横比（A4 约 1.401）永远小于纸张本身（1.414），
   * 页面盒按纸张高度绘制时每页都会溢出一两毫米，从而在末尾多出一张空白纸。
   */
  sheetHeight: string;
}

/** 功能：页面盒相对纸张留出的纵向安全余量（毫米），抵消可打印区域纵横比小于纸张带来的溢出；时间：2026-10-29；作者：lq。 */
const SHEET_HEIGHT_SAFETY_MM = 6;

/**
 * 功能：把像素页面换算成打印纸张；命中标准纸张时改用毫米，保证与物理纸张严格相等；
 * 时间：2026-10-29；作者：lq。
 */
export function resolvePrintPaper(geometry: PrintPageGeometry): ResolvedPaper {
  const widthPx = Number.isFinite(geometry.widthPx) && geometry.widthPx > 0 ? geometry.widthPx : 793.7;
  const heightPx = Number.isFinite(geometry.heightPx) && geometry.heightPx > 0 ? geometry.heightPx : 1122.52;
  const widthMm = widthPx / PX_PER_MM;
  const heightMm = heightPx / PX_PER_MM;
  const landscape = widthPx > heightPx;
  for (const paper of STANDARD_PAPERS) {
    const long = Math.max(paper.widthMm, paper.heightMm);
    const short = Math.min(paper.widthMm, paper.heightMm);
    const targetWidth = landscape ? long : short;
    const targetHeight = landscape ? short : long;
    // 1.5mm 容差覆盖 A4/Letter 的 96dpi 取整误差（794×1123、816×1056 等）。
    if (Math.abs(widthMm - targetWidth) > 1.5 || Math.abs(heightMm - targetHeight) > 1.5) continue;
    return {
      pageRule: `size: ${targetWidth}mm ${targetHeight}mm;`,
      width: `${round(targetWidth * PX_PER_MM)}px`,
      height: `${round(targetHeight * PX_PER_MM)}px`,
      name: paper.name,
      widthMm: targetWidth,
      heightMm: targetHeight,
      sheetHeight: `${round((targetHeight - SHEET_HEIGHT_SAFETY_MM) * PX_PER_MM)}px`,
    };
  }
  return {
    pageRule: `size: ${widthPx}px ${heightPx}px;`,
    width: `${widthPx}px`,
    height: `${heightPx}px`,
    widthMm,
    heightMm,
    sheetHeight: `${round(heightPx - SHEET_HEIGHT_SAFETY_MM * PX_PER_MM)}px`,
  };
}

function round(value: number): number {
  return Math.round(value * 1000) / 1000;
}

/** 功能：打印根节点标记，用于隔离应用界面与打印内容；时间：2026-10-29；作者：lq。 */
export const PRINT_ROOT_ATTRIBUTE = 'data-mo-print-root';

export interface PrintPagesOptions extends PrintPageGeometry {
  /** 每个元素都会被渲染成独立的一页纸，元素数量即打印页数。 */
  pages: HTMLElement[];
  /**
   * 每页各自的纸张尺寸，用于「同一文档里既有纵向节又有横向节」的场景；
   * 缺省时整份文档统一使用 widthPx / heightPx（例如 Excel 单表打印）。
   */
  pageSizes?: PrintPageGeometry[];
  /** 功能：已按完整纸张裁切的原生页面保留全高，避免旧流程安全余量裁掉页脚；时间：2026-11-11；作者：lq。 */
  fullPaper?: boolean;
  /** 页面内部的附加样式（例如编辑器自身的缩放复位）。 */
  extraCss?: string;
  /** 打印任务标题，便于系统打印队列识别。 */
  title?: string;
}

/** 功能：把每页的纸张解算结果按尺寸分组，供命名 @page 规则使用；时间：2026-10-29；作者：lq。 */
interface PaperGroup {
  name: string;
  paper: ResolvedPaper;
}

/**
 * 功能：解析每页纸张并返回分组；多节混合方向时按尺寸去重成命名页组；
 * 时间：2026-10-29；作者：lq。
 */
function resolvePapers(options: PrintPagesOptions): { papers: ResolvedPaper[]; groups: PaperGroup[]; rootWidthMm: number } {
  const sizes = options.pageSizes && options.pageSizes.length === options.pages.length ? options.pageSizes : null;
  const papers = options.pages.map((_, index) => {
    const paper = resolvePrintPaper(sizes ? sizes[index] : options);
    return options.fullPaper ? { ...paper, sheetHeight: paper.height } : paper;
  });
  const groups: PaperGroup[] = [];
  for (const paper of papers) {
    const key = `${paper.width}|${paper.height}`;
    if (groups.some((group) => `${group.paper.width}|${group.paper.height}` === key)) continue;
    groups.push({ name: `mo-sheet-${groups.length + 1}`, paper });
  }
  const rootWidthMm = papers.reduce((max, paper) => Math.max(max, paper.widthMm), 0);
  return { papers, groups, rootWidthMm };
}

export interface PrintRootHandle {
  root: HTMLElement;
  cleanup: () => void;
}

/**
 * 功能：清理上一次打印遗留的根节点与样式；残留的 `body > *:not(root)` 规则会把
 * 下一次分页的测量容器变成 display:none，导致按错误行高分页；时间：2026-10-29；作者：lq。
 */
export function clearPrintRoots(): void {
  document.querySelectorAll(`[${PRINT_ROOT_ATTRIBUTE}]`).forEach((node) => node.remove());
  document.querySelectorAll('style[data-mo-print-style]').forEach((node) => node.remove());
}

/**
 * 功能：构建与 @page 尺寸严格一致的打印文档；时间：2026-10-29；作者：lq。
 *
 * 打印根节点在屏幕上始终 display:none，只有 @media print 才显示，
 * 因此即使 afterprint 没有触发也不会污染界面。
 */
export function createPrintRoot(options: PrintPagesOptions): PrintRootHandle {
  clearPrintRoots();
  const { papers, groups, rootWidthMm } = resolvePapers(options);
  const mixed = groups.length > 1;
  // 功能：混合纸张按各命名页宽排版，最大横向页宽不能触发 Chromium 对纵向页整体缩小；时间：2026-11-11；作者：lq。
  const layoutWidth = options.fullPaper && mixed ? 'auto' : `${rootWidthMm}mm`;
  const maximumWidth = options.fullPaper && mixed ? 'none' : `${rootWidthMm}mm`;
  const root = document.createElement('div');
  root.setAttribute(PRINT_ROOT_ATTRIBUTE, 'true');
  if (options.title) root.setAttribute('data-print-title', options.title);
  options.pages.forEach((page, index) => {
    const paper = papers[index];
    page.classList.add('mo-print-page');
    // 功能：多节混合方向时给每页挂命名页组，Chromium 会按命名 @page 分配纸张；时间：2026-10-29；作者：lq。
    if (mixed) {
      const group = groups.find((item) => item.paper.width === paper.width && item.paper.height === paper.height);
      if (group) page.classList.add(group.name);
    }
    // 功能：逐页写入尺寸，避免外部 flex/gap 影响页数；页边距仍由页面自身 padding 承担；时间：2026-10-29；作者：lq。
    page.style.setProperty('width', paper.width, 'important');
    page.style.setProperty('height', paper.sheetHeight, 'important');
    page.style.setProperty('margin', '0', 'important');
    root.appendChild(page);
  });

  // 功能：混合方向时逐个页组声明纸张，单一纸张时保持原来的单条 @page；时间：2026-10-29；作者：lq。
  const pageRules = mixed
    ? groups
      .map(
        (group) => `
      @page ${group.name} { ${group.paper.pageRule} margin: 0; }
      [${PRINT_ROOT_ATTRIBUTE}] .mo-print-page.${group.name} { page: ${group.name}; }`,
      )
      .join('')
    : `@page { ${papers[0].pageRule} margin: 0; }`;

  const style = document.createElement('style');
  style.setAttribute('data-mo-print-style', 'true');
  style.textContent = `
    [${PRINT_ROOT_ATTRIBUTE}] { display: none; }
    @media print {
      html, body {
        height: auto !important;
        min-height: 0 !important;
        max-height: none !important;
        overflow: visible !important;
        margin: 0 !important;
        padding: 0 !important;
        background: #ffffff !important;
        /* 功能：把布局宽度锁到纸张宽度，否则 WebKit 会按当前窗口宽度排版再整体缩放到纸张，导致内容变小且贴左；时间：2026-10-29；作者：lq。 */
        width: ${layoutWidth} !important;
        max-width: ${maximumWidth} !important;
      }
      body > *:not([${PRINT_ROOT_ATTRIBUTE}]) { display: none !important; }
      [${PRINT_ROOT_ATTRIBUTE}] {
        display: block !important;
        position: static !important;
        width: ${layoutWidth} !important;
        max-width: ${maximumWidth} !important;
        height: auto !important;
        margin: 0 !important;
        padding: 0 !important;
        background: #ffffff !important;
      }
      [${PRINT_ROOT_ATTRIBUTE}] .mo-print-page {
        display: block !important;
        position: relative !important;
        margin: ${options.fullPaper && mixed ? '0' : '0 auto'} !important;
        border: 0 !important;
        border-radius: 0 !important;
        box-shadow: none !important;
        background: #ffffff !important;
        overflow: hidden !important;
        break-inside: avoid !important;
        break-after: page !important;
        page-break-after: always !important;
      }
      /* 功能：仅最后一页取消分页，避免末尾多出一张空白纸；时间：2026-10-29；作者：lq。 */
      [${PRINT_ROOT_ATTRIBUTE}] .mo-print-page:last-child {
        break-after: auto !important;
        page-break-after: auto !important;
      }
      *, *::before, *::after {
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
      }
      [data-table-resize] { display: none !important; }
      ${pageRules}
      ${options.extraCss ?? ''}
    }
  `;

  document.head.appendChild(style);
  document.body.appendChild(root);
  const cleanup = () => {
    root.remove();
    style.remove();
    window.removeEventListener('afterprint', cleanup);
  };
  window.addEventListener('afterprint', cleanup);
  return { root, cleanup };
}

/** 功能：等待字体与页面内图片就绪，避免打印时缺字或缺图；时间：2026-10-29；作者：lq。 */
export async function waitForPrintAssets(root: HTMLElement, timeoutMs = 4000): Promise<void> {
  const images = Array.from(root.querySelectorAll('img'));
  const jobs: Array<Promise<unknown>> = images.map((image) =>
    image.decode ? image.decode().catch(() => undefined) : Promise.resolve(),
  );
  if (typeof document !== 'undefined' && document.fonts) jobs.push(document.fonts.ready.catch(() => undefined));
  await Promise.race([Promise.all(jobs), new Promise((resolve) => setTimeout(resolve, timeoutMs))]);
  await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
}

/**
 * 功能：把给定页面交给系统打印，页数严格等于传入页面数量；时间：2026-10-29；作者：lq。
 */
export async function printPages(options: PrintPagesOptions): Promise<boolean> {
  if (options.pages.length === 0) return false;
  const { root, cleanup } = createPrintRoot(options);
  let printed = false;
  try {
    await waitForPrintAssets(root);
    printed = await printCurrentView();
    return printed;
  } finally {
    if (!printed) cleanup();
    // 功能：系统打印面板在 WebKit 上可能不派发 afterprint，保留兜底清理；时间：2026-10-29；作者：lq。
    else setTimeout(cleanup, 30_000);
  }
}

export interface PrintPreviewOptions extends PrintPageGeometry {
  /** 预览容器，函数会先清空它。 */
  host: HTMLElement;
  /** 已经构建好的页面元素，数量即预览页数。 */
  pages: HTMLElement[];
  /** 预览区左右留白，默认 24px。 */
  gutter?: number;
  /** 是否在每页上方显示“第 n / N 页”，默认显示。 */
  labels?: boolean;
  /** 固定缩放比例（1 = 实际大小）；不传时按 fit 规则自适应。 */
  scale?: number;
  /** 功能：单页预览的适配方式——整页可见 / 适应页宽（可纵向滚动）/ 实际大小；时间：2026-11-01；作者：lq。 */
  fit?: PrintPreviewFit;
  /** 每页各自的纸张尺寸，用于多节混合方向的文档；缺省时所有页共用 widthPx / heightPx。 */
  pageSizes?: PrintPageGeometry[];
  /** 单页模式：只渲染当前页，配合 setPage 翻页；默认 false（整列平铺）。 */
  singlePage?: boolean;
  /** 功能：每页画框构建后回调，供调用方叠加边距参考线等交互层；时间：2026-11-01；作者：lq。 */
  overlay?: (context: PreviewOverlayContext) => void;
}

/** 功能：打印预览的缩放策略；时间：2026-11-01；作者：lq。 */
export type PrintPreviewFit = 'page' | 'width' | 'actual';

/** 功能：预览画框上下文，供调用方叠加边距参考线等交互层；时间：2026-11-01；作者：lq。 */
export interface PreviewOverlayContext {
  /** 已按 scale 缩放的页面画框。 */
  frame: HTMLElement;
  /** 页面元素本身（未缩放）。 */
  page: HTMLElement;
  /** 画框相对纸张的缩放比例。 */
  scale: number;
  /** 页码索引（0 起）。 */
  index: number;
}

/** 功能：页边距参考线的可拖动规格；时间：2026-11-01；作者：lq。 */
export interface PreviewMarginSpec {
  /** 功能：只读文档仍可显示参考线，不开放拖拽手柄；时间：2026-11-11；作者：lq。 */
  readonly?: boolean;
  margins: { left: number; right: number; top: number; bottom: number };
  pageWidthPx: number;
  pageHeightPx: number;
  /** 拖动结束后回调，单位为未缩放的页面像素。 */
  onChange: (edge: 'left' | 'right' | 'top' | 'bottom', valuePx: number) => void;
}

/**
 * 功能：在预览页上叠加四条可拖动的页边距参考线，直接拖动即改边距；
 * 时间：2026-11-01；作者：lq。
 *
 * 说明：拖动期间只移动参考线本身，松手才回调改值，避免每帧重排整份文档。
 */
export function attachPreviewMargins(context: PreviewOverlayContext, spec: PreviewMarginSpec): () => void {
  const { frame, scale } = context;
  const { margins, pageWidthPx, pageHeightPx, onChange } = spec;
  const factor = scale > 0 ? scale : 1;
  const layer = document.createElement('div');
  layer.dataset.previewMargins = 'true';
  layer.style.cssText = 'position:absolute;inset:0;pointer-events:none;';
  frame.appendChild(layer);

  const band = document.createElement('div');
  band.style.cssText = 'position:absolute;border:1px dashed rgba(52,120,246,0.55);pointer-events:none;';
  layer.appendChild(band);

  const box = {
    left: margins.left * factor,
    top: margins.top * factor,
    right: (pageWidthPx - margins.right) * factor,
    bottom: (pageHeightPx - margins.bottom) * factor,
  };
  const paint = () => {
    band.style.left = `${box.left}px`;
    band.style.top = `${box.top}px`;
    band.style.width = `${Math.max(0, box.right - box.left)}px`;
    band.style.height = `${Math.max(0, box.bottom - box.top)}px`;
    for (const edge of edges) {
      const strip = strips[edge];
      if (edge === 'left') strip.style.cssText = stripStyle(`left:${box.left - 4}px;top:${box.top}px;width:8px;height:${Math.max(0, box.bottom - box.top)}px;`, 'ew-resize');
      if (edge === 'right') strip.style.cssText = stripStyle(`left:${box.right - 4}px;top:${box.top}px;width:8px;height:${Math.max(0, box.bottom - box.top)}px;`, 'ew-resize');
      if (edge === 'top') strip.style.cssText = stripStyle(`left:${box.left}px;top:${box.top - 4}px;width:${Math.max(0, box.right - box.left)}px;height:8px;`, 'ns-resize');
      if (edge === 'bottom') strip.style.cssText = stripStyle(`left:${box.left}px;top:${box.bottom - 4}px;width:${Math.max(0, box.right - box.left)}px;height:8px;`, 'ns-resize');
    }
  };
  const stripStyle = (position: string, cursor: string) =>
    `position:absolute;${position}pointer-events:auto;cursor:${cursor};`;

  const edges = ['left', 'right', 'top', 'bottom'] as const;
  const strips = {} as Record<(typeof edges)[number], HTMLElement>;
  // 功能：取消和卸载只清理拖动，不能提交旧弹窗的页边距；时间：2026-11-11；作者：lq。
  let cancelDrag: (() => void) | undefined;
  for (const edge of edges) {
    const strip = document.createElement('span');
    strip.setAttribute('role', 'separator');
    strip.setAttribute('aria-label', `拖动调整${edge === 'left' ? '左' : edge === 'right' ? '右' : edge === 'top' ? '上' : '下'}页边距`);
    strip.title = strip.getAttribute('aria-label')!;
    strip.addEventListener('pointerdown', (event) => {
      if (event.button !== 0 || spec.readonly) return;
      cancelDrag?.();
      event.preventDefault();
      event.stopPropagation();
      const startX = event.clientX;
      const startY = event.clientY;
      const initial = { ...box };
      const move = (pointer: PointerEvent) => {
        if (pointer.pointerId !== event.pointerId) return;
        const dx = pointer.clientX - startX;
        const dy = pointer.clientY - startY;
        if (edge === 'left') box.left = Math.max(0, Math.min(initial.right - 40, initial.left + dx));
        else if (edge === 'right') box.right = Math.min(pageWidthPx * factor, Math.max(initial.left + 40, initial.right + dx));
        else if (edge === 'top') box.top = Math.max(0, Math.min(initial.bottom - 40, initial.top + dy));
        else box.bottom = Math.min(pageHeightPx * factor, Math.max(initial.top + 40, initial.bottom + dy));
        paint();
      };
      const cleanup = () => {
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', finish);
        window.removeEventListener('pointercancel', abort);
        cancelDrag = undefined;
      };
      const abort = () => { cleanup(); Object.assign(box, initial); paint(); };
      const finish = (pointer: PointerEvent) => {
        if (pointer.pointerId !== event.pointerId) return;
        cleanup();
        const value = edge === 'left' ? box.left : edge === 'right' ? pageWidthPx * factor - box.right : edge === 'top' ? box.top : pageHeightPx * factor - box.bottom;
        onChange(edge, Math.max(0, Math.round(value / factor)));
      };
      cancelDrag = abort;
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', finish);
      window.addEventListener('pointercancel', abort);
    });
    strips[edge] = strip;
    layer.appendChild(strip);
  }
  paint();
  if (spec.readonly) Object.values(strips).forEach(strip => { strip.style.pointerEvents = 'none'; strip.style.cursor = 'default'; });
  return () => { cancelDrag?.(); layer.remove(); };
}

/** 功能：打印预览句柄——清理、翻页与页数；时间：2026-11-01；作者：lq。 */
export interface PrintPreviewHandle {
  cleanup: () => void;
  /** 切到第 index 页（0 起）；整列平铺模式下无效。 */
  setPage: (index: number) => void;
  pageCount: number;
}

/**
 * 功能：在应用内渲染前置打印预览（Word / Excel 共用同一套视觉与缩放规则），
 * 预览页数与实际打印页数严格一致；时间：2026-10-29；作者：lq。
 *
 * 返回的清理函数会移除预览 DOM，但不销毁调用方传入的页面元素。
 */
export function renderPrintPreview(options: PrintPreviewOptions): PrintPreviewHandle {
  const { host, pages } = options;
  host.replaceChildren();
  const sizes = options.pageSizes && options.pageSizes.length === pages.length ? options.pageSizes : null;
  const geometryAt = (index: number): PrintPageGeometry => {
    const size = sizes ? sizes[index] : options;
    const width = Number.isFinite(size.widthPx) && size.widthPx > 0 ? size.widthPx : 793.7;
    const height = Number.isFinite(size.heightPx) && size.heightPx > 0 ? size.heightPx : 1122.52;
    return { widthPx: width, heightPx: height };
  };
  const papers = pages.map((_, index) => resolvePrintPaper(geometryAt(index)));
  // 功能：混合方向时用最大页宽/页高统一缩放，保证整列视觉比例一致；时间：2026-10-29；作者：lq。
  const maxWidth = papers.reduce((max, paper) => Math.max(max, Number.parseFloat(paper.width)), 0);
  const maxHeight = papers.reduce((max, paper) => Math.max(max, Number.parseFloat(paper.sheetHeight)), 0);
  const mixed = papers.some((paper) => paper.width !== papers[0].width || paper.height !== papers[0].height);
  const gutter = options.gutter ?? 24;
  const cleanup = () => host.replaceChildren();
  const clampIndex = (index: number) => Math.max(0, Math.min(pages.length - 1, Math.trunc(index) || 0));

  // 功能：单页模式——一次只渲染一页，按容器宽高双向适配，接近 Office/GenOffice 的打印预览；时间：2026-11-01；作者：lq。
  if (options.singlePage) {
    const stage = document.createElement('div');
    stage.style.cssText = 'display:flex;align-items:flex-start;justify-content:center;width:100%;min-height:100%;padding:10px;box-sizing:border-box;';
    host.appendChild(stage);
    const paint = (index: number) => {
      const paper = papers[index];
      if (!paper) return;
      const width = Number.parseFloat(paper.width);
      const height = Number.parseFloat(paper.sheetHeight);
      const availableWidth = Math.max(200, (host.clientWidth || 640) - gutter * 2);
      const availableHeight = Math.max(200, (host.clientHeight || 520) - gutter * 2);
      // 功能：默认整页可见；“适应页宽”放大到页宽、允许纵向滚动，细节更清楚；时间：2026-11-01；作者：lq。
      const fit: PrintPreviewFit = options.scale && options.scale > 0 ? 'actual' : options.fit ?? 'page';
      const scale =
        fit === 'actual'
          ? 1
          : fit === 'width'
            ? Math.min(1, availableWidth / maxWidth)
            : Math.min(1, availableWidth / maxWidth, availableHeight / maxHeight);
      const frame = document.createElement('div');
      frame.style.cssText = `position:relative;flex:0 0 auto;width:${width * scale}px;height:${height * scale}px;background:#ffffff;box-shadow:0 1px 3px rgba(15,23,42,.16),0 0 0 1px rgba(15,23,42,.10);overflow:hidden;`;
      const page = pages[index];
      page.style.position = 'absolute';
      page.style.left = '0';
      page.style.top = '0';
      page.style.margin = '0';
      page.style.transform = `scale(${scale})`;
      page.style.transformOrigin = 'top left';
      frame.appendChild(page);
      stage.replaceChildren(frame);
      options.overlay?.({ frame, page, scale, index });
    };
    paint(0);
    return {
      cleanup,
      pageCount: pages.length,
      setPage: (index: number) => paint(clampIndex(index)),
    };
  }

  // 功能：整列平铺模式——页数少时一屏看全，仍是 Excel 多页对照的默认视图；时间：2026-10-29；作者：lq。
  const available = Math.max(240, (host.clientWidth || 640) - gutter);
  const scale = options.scale && options.scale > 0 ? options.scale : Math.min(1, available / maxWidth);
  const list = document.createElement('div');
  list.style.cssText = 'display:flex;flex-direction:column;align-items:center;gap:16px;padding:14px 0 18px;';
  pages.forEach((page, index) => {
    const paper = papers[index];
    const width = Number.parseFloat(paper.width);
    const height = Number.parseFloat(paper.sheetHeight);
    const column = document.createElement('div');
    column.style.cssText = 'display:flex;flex-direction:column;align-items:center;gap:5px;';
    if (options.labels !== false) {
      const label = document.createElement('div');
      // 功能：混合方向时在页码后标出该页纸张与方向，用户能一眼看出哪些页是横向；时间：2026-10-29；作者：lq。
      const suffix = mixed
        ? ` · ${paper.name ?? '自定义'} ${paper.widthMm > paper.heightMm ? '横向' : '纵向'}`
        : '';
      label.textContent = `第 ${index + 1} / ${pages.length} 页${suffix}`;
      label.style.cssText = 'font-size:11px;line-height:14px;color:var(--mo-text-muted,#667085);';
      column.appendChild(label);
    }
    const frame = document.createElement('div');
    frame.style.cssText = `position:relative;width:${width * scale}px;height:${height * scale}px;background:#ffffff;box-shadow:0 2px 12px rgba(0,0,0,.32);overflow:hidden;`;
    page.style.position = 'absolute';
    page.style.left = '0';
    page.style.top = '0';
    page.style.margin = '0';
    page.style.transform = `scale(${scale})`;
    page.style.transformOrigin = 'top left';
    frame.appendChild(page);
    column.appendChild(frame);
    list.appendChild(column);
    options.overlay?.({ frame, page, scale, index });
  });
  host.appendChild(list);
  return { cleanup, pageCount: pages.length, setPage: () => undefined };
}

/** Print the current view. Tauri uses the native WKWebView/WebView2 print API. */
export async function printCurrentView(): Promise<boolean> {
  if (isTauriRuntime()) {
    try {
      await invoke('print_window');
      return true;
    } catch (error) {
      console.warn('Tauri 原生打印失败，回退到 window.print()', error);
    }
  }
  if (typeof window.print === 'function') {
    window.print();
    return true;
  }
  return false;
}
