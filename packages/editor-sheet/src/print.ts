import type { CellRange, SheetModel, SheetPageSetup, WorkbookModel } from '@mini-office/document-model';
import { defaultSheetPageSetup } from '@mini-office/document-model';
import { columnLabel } from '@mini-office/shared';
import { clearPrintRoots } from '@mini-office/platform-adapter';
import { cellAt, cellDisplay, cellStyle, usedRange } from './sheet-utils';

/**
 * 功能：Excel 打印分页；按真实行高与列宽切页，页数、页边距与合并单元格都与纸张严格对应；
 * 时间：2026-10-29；作者：lq。
 *
 * 与旧实现的差异（都对应线上反馈的缺陷）：
 * 1. 页边距只由页面盒 padding 承担，@page 不再叠加 margin，避免可打印区域被二次压缩后右侧列丢失。
 * 2. 列宽超出可打印宽度时按列分页，而不是让 overflow:hidden 把右侧列裁掉。
 * 3. 行高按渲染后的真实高度切分，换行文本不会被裁在页脚。
 * 4. 合并单元格跨页时按页裁剪并保留边框底纹，不再整块丢失。
 * 5. 最后一页不再带 break-after，末尾不会多出空白页。
 */

/** CSS 像素与毫米的换算基准：96dpi。 */
const PX_PER_MM = 96 / 25.4;
/** A4 纸张尺寸（像素），与 @page 使用同一物理长度。 */
const A4_WIDTH_PX = Math.round(210 * PX_PER_MM * 100) / 100;
const A4_HEIGHT_PX = Math.round(297 * PX_PER_MM * 100) / 100;
/** 页眉信息条高度（工作表名 + 页码）。 */
const TITLE_HEIGHT = 24;
/** 打印行列标题时的行号列宽与列标行高。 */
const HEADING_COLUMN_WIDTH = 40;
const HEADING_ROW_HEIGHT = 22;

export interface SheetPageSize {
  width: number;
  height: number;
}

/** 功能：返回当前页面设置的纸张像素尺寸，导出与打印共用；时间：2026-10-29；作者：lq。 */
export function sheetPageSize(setup: SheetPageSetup): SheetPageSize {
  return setup.orientation === 'portrait'
    ? { width: A4_WIDTH_PX, height: A4_HEIGHT_PX }
    : { width: A4_HEIGHT_PX, height: A4_WIDTH_PX };
}

interface ColumnChunk {
  start: number;
  end: number;
  /** 该列组的缩放后宽度。 */
  width: number;
}

interface RowChunk {
  start: number;
  end: number;
}

/** 功能：把英寸页边距换算成打印像素；时间：2026-10-29；作者：lq。 */
function inchToPx(value: number): number {
  return Math.round(value * 96);
}

/**
 * 功能：解析“打印标题”行区间（支持 `1:2`、`$1:$2`、`2`），并裁剪到打印范围内；
 * 时间：2026-11-01；作者：lq。
 *
 * 说明：非法输入返回 null，由调用方回退到“首行当表头”，避免用户输入错误导致整表不打印。
 */
export function parsePrintTitleRows(
  text: string | undefined,
  target: CellRange,
): { r1: number; r2: number } | null {
  if (!text) return null;
  const match = /^\s*\$?(\d{1,7})\s*(?::\s*\$?(\d{1,7})\s*)?$/.exec(text);
  if (!match) return null;
  const first = Number(match[1]) - 1;
  const last = match[2] ? Number(match[2]) - 1 : first;
  const clamped1 = Math.max(target.r1, Math.min(first, last));
  const clamped2 = Math.min(target.r2, Math.max(first, last));
  if (clamped2 < clamped1) return null;
  return { r1: clamped1, r2: clamped2 };
}

/** 功能：统计一段连续行的模型高度，供打印标题占用高度预算；时间：2026-11-01；作者：lq。 */
function rowsHeight(sheet: SheetModel, r1: number, r2: number): number {
  let total = 0;
  for (let row = r1; row <= r2; row += 1) total += sheet.rowHeights[row] ?? sheet.defaultRowHeight;
  return total;
}

/** 功能：生成工作表打印页面，返回的每个元素就是一页纸；时间：2026-10-29；作者：lq。 */
export function buildSheetPrintPages(
  workbook: WorkbookModel,
  sheet: SheetModel,
  container: HTMLElement,
): HTMLElement[] {
  const setup = sheet.pageSetup ?? defaultSheetPageSetup();
  // 功能：先清理上次打印残留的样式；残留规则会把分页测量容器隐藏成 0 高度，导致按模型行高分页而裁掉行；时间：2026-10-29；作者：lq。
  clearPrintRoots();
  const { width: pageWidth, height: pageHeight } = sheetPageSize(setup);
  const marginLeft = inchToPx(setup.margins.left);
  const marginRight = inchToPx(setup.margins.right);
  const marginTop = inchToPx(setup.margins.top);
  const marginBottom = inchToPx(setup.margins.bottom);
  const contentWidth = Math.max(60, pageWidth - marginLeft - marginRight);
  const contentHeight = Math.max(60, pageHeight - marginTop - marginBottom);

  const target = printableRange(sheet);
  const columnWidths: number[] = [];
  for (let col = target.c1; col <= target.c2; col += 1) {
    columnWidths.push(sheet.colWidths[col] ?? sheet.defaultColWidth);
  }
  const totalWidth = columnWidths.reduce((sum, width) => sum + width, 0) || 100;

  const headingWidth = setup.printHeadings ? HEADING_COLUMN_WIDTH : 0;
  const headingHeight = setup.printHeadings ? HEADING_ROW_HEIGHT : 0;
  const bodyWidth = Math.max(40, contentWidth - headingWidth);
  const wrapperHeight = Math.max(1, contentHeight - TITLE_HEIGHT);
  const tableHeight = Math.max(40, wrapperHeight - headingHeight);

  // 功能：调整为页宽/页高只按需缩小，绝不放大小表；未设置时按用户比例并允许横向分页；时间：2026-10-29；作者：lq。
  const fitCandidates: number[] = [];
  if (setup.fitToWidth) fitCandidates.push(bodyWidth / totalWidth);
  if (setup.fitToHeight) fitCandidates.push(tableHeight / Math.max(1, totalRowHeight(sheet, target)));
  const rawScale = fitCandidates.length > 0 ? Math.min(1, ...fitCandidates) : setup.scalePercent / 100;
  const scale = Math.max(0.1, Math.min(4, rawScale));

  // 功能：配置了“打印标题”就重复这些行；未配置时沿用“首行当表头”的历史行为，不改变已有文档的打印结果；时间：2026-11-01；作者：lq。
  const explicitTitles = parsePrintTitleRows(setup.printTitleRows, target);
  const titleRows = explicitTitles ?? (target.r2 > target.r1 ? { r1: target.r1, r2: target.r1 } : null);
  const headerRow = titleRows ? titleRows.r1 : null;
  const bodyStart = titleRows ? titleRows.r2 + 1 : target.r1;
  const headerHeight = titleRows ? Math.round(rowsHeight(sheet, titleRows.r1, titleRows.r2) * scale) : 0;

  container.style.cssText = `position:fixed;left:-200000px;top:0;width:${pageWidth}px;background:#fff;z-index:-1;`;
  document.body.appendChild(container);

  const columnChunks = splitColumns(columnWidths, scale, bodyWidth, target);
  const pageBodyHeight = Math.max(1, wrapperHeight - headingHeight - headerHeight);
  const pages: HTMLElement[] = [];
  // 功能：先按列分组再按行分页，页面顺序与 Excel 一致（先向下再向右）；时间：2026-10-29；作者：lq。
  const rowChunksByColumn = new Map<number, RowChunk[]>();
  for (const chunk of columnChunks) {
    rowChunksByColumn.set(chunk.start, measureRowChunks({
      workbook, sheet, target, chunk, scale, titleRows, bodyStart,
      columnWidths, printGridLines: setup.printGridLines, pageBodyHeight, container,
    }));
  }
  const totalPages = columnChunks.reduce((sum, chunk) => sum + (rowChunksByColumn.get(chunk.start)?.length ?? 0), 0);
  let pageIndex = 0;

  for (const chunk of columnChunks) {
    for (const rows of rowChunksByColumn.get(chunk.start) ?? []) {
      pages.push(buildPage({
        workbook, sheet, setup, target, chunk, rows, scale,
        pageWidth, pageHeight, contentWidth, contentHeight,
        margins: { top: marginTop, right: marginRight, bottom: marginBottom, left: marginLeft },
        columnWidths, headerRow, titleRows,
        pageNumber: (setup.firstPageNumber ?? 1) + pageIndex,
        totalPages, headingWidth, headingHeight,
      }));
      pageIndex += 1;
    }
  }

  container.replaceChildren(...pages);
  return pages;
}

/** 功能：计算打印范围，并把浮动图片占用的行列纳入；时间：2026-10-29；作者：lq。 */
function printableRange(sheet: SheetModel): CellRange {
  const target = usedRange(sheet);
  const rowHeight = (row: number) => sheet.rowHeights[row] ?? sheet.defaultRowHeight;
  const colWidth = (col: number) => sheet.colWidths[col] ?? sheet.defaultColWidth;
  for (const image of sheet.images ?? []) {
    let height = image.offsetY + image.height;
    let endRow = image.row;
    while (height > rowHeight(endRow) && endRow < 1_048_575) { height -= rowHeight(endRow); endRow += 1; }
    let width = image.offsetX + image.width;
    let endCol = image.col;
    while (width > colWidth(endCol) && endCol < 16_383) { width -= colWidth(endCol); endCol += 1; }
    target.r2 = Math.max(target.r2, endRow);
    target.c2 = Math.max(target.c2, endCol);
  }
  return target;
}

/** 功能：计算打印范围内所有行的模型总高度，用于“调整为页高”；时间：2026-10-29；作者：lq。 */
function totalRowHeight(sheet: SheetModel, target: CellRange): number {
  let total = 0;
  for (let row = target.r1; row <= target.r2; row += 1) total += sheet.rowHeights[row] ?? sheet.defaultRowHeight;
  return total;
}

/** 功能：按可打印宽度切分列组，保证右侧列不被裁掉；时间：2026-10-29；作者：lq。 */
function splitColumns(columnWidths: number[], scale: number, availableWidth: number, target: CellRange): ColumnChunk[] {
  const chunks: ColumnChunk[] = [];
  let start = target.c1;
  let width = 0;
  columnWidths.forEach((raw, index) => {
    const col = target.c1 + index;
    const scaled = raw * scale;
    if (width > 0 && width + scaled > availableWidth + 0.5) {
      chunks.push({ start, end: col - 1, width });
      start = col;
      width = 0;
    }
    width += scaled;
  });
  chunks.push({ start, end: target.c2, width });
  return chunks;
}

/** 功能：先渲染一次完整列组读取每行真实高度，再按页高切分；时间：2026-10-29；作者：lq。 */
function measureRowChunks(params: {
  workbook: WorkbookModel;
  sheet: SheetModel;
  target: CellRange;
  chunk: ColumnChunk;
  scale: number;
  titleRows: { r1: number; r2: number } | null;
  bodyStart: number;
  columnWidths: number[];
  printGridLines: boolean;
  pageBodyHeight: number;
  container: HTMLElement;
}): RowChunk[] {
  const { workbook, sheet, target, chunk, scale, titleRows, bodyStart, columnWidths, printGridLines, pageBodyHeight, container } = params;
  if (bodyStart > target.r2) return [{ start: target.r1, end: target.r1 }];

  const probe = document.createElement('div');
  probe.style.cssText = `position:absolute;left:0;top:0;width:${chunk.width}px;visibility:hidden;`;
  const table = buildChunkTable({
    workbook, sheet, target, chunk, columnWidths, scale, printGridLines,
    rowStart: bodyStart, rowEnd: target.r2, titleRows, repeatTitles: false, ignoreRowSpans: true,
  });
  probe.appendChild(table);
  container.appendChild(probe);
  const measured: number[] = [];
  table.querySelectorAll<HTMLTableRowElement>('tbody > tr').forEach((tr) => {
    const height = tr.getBoundingClientRect().height;
    measured.push(height > 0 ? height : parseFloat(tr.style.height) || 0);
  });
  probe.remove();

  const rowChunks: RowChunk[] = [];
  let start = bodyStart;
  let used = 0;
  for (let index = 0; index < measured.length; index += 1) {
    const row = bodyStart + index;
    const height = measured[index] ?? 0;
    // 功能：单行超过整页时独占一页，不再无限扩张或直接裁掉；时间：2026-10-29；作者：lq。
    if (used > 0 && used + height > pageBodyHeight + 0.5) {
      rowChunks.push({ start, end: row - 1 });
      start = row;
      used = 0;
    }
    used += height;
  }
  rowChunks.push({ start, end: target.r2 });
  return rowChunks;
}

/** 功能：组装单页纸，页边距由 padding 承担，内容区严格裁在纸内；时间：2026-10-29；作者：lq。 */
function buildPage(params: {
  workbook: WorkbookModel;
  sheet: SheetModel;
  setup: SheetPageSetup;
  target: CellRange;
  chunk: ColumnChunk;
  rows: RowChunk;
  scale: number;
  pageWidth: number;
  pageHeight: number;
  contentWidth: number;
  contentHeight: number;
  margins: { top: number; right: number; bottom: number; left: number };
  columnWidths: number[];
  headerRow: number | null;
  titleRows: { r1: number; r2: number } | null;
  pageNumber: number;
  totalPages: number;
  headingWidth: number;
  headingHeight: number;
}): HTMLElement {
  const {
    workbook, sheet, setup, target, chunk, rows, scale, pageWidth, pageHeight,
    contentWidth, contentHeight, margins, columnWidths, headerRow, titleRows, pageNumber, totalPages,
    headingWidth, headingHeight,
  } = params;

  const page = document.createElement('div');
  page.dataset.sheetPrintPage = 'true';
  page.style.cssText = `width:${pageWidth}px;height:${pageHeight}px;background:#fff;color:#1D2939;box-sizing:border-box;padding:${margins.top}px ${margins.right}px ${margins.bottom}px ${margins.left}px;font-family:等线,PingFang SC,sans-serif;position:relative;overflow:hidden;`;

  const title = document.createElement('div');
  title.style.cssText = `height:${TITLE_HEIGHT}px;display:flex;align-items:flex-end;justify-content:space-between;font-size:10px;color:#98A2B3;border-bottom:1px solid #E4E9F2;margin-bottom:6px;padding-bottom:3px;box-sizing:border-box;`;
  const pageLabel = totalPages > 1 ? `第 ${pageNumber} / ${totalPages} 页` : '';
  title.innerHTML = `<span style="font-weight:600;color:#475467;">${escapeHtml(sheet.name)}</span><span>${pageLabel}</span>`;
  page.appendChild(title);

  const wrapper = document.createElement('div');
  wrapper.style.cssText = `position:relative;width:${contentWidth}px;height:${Math.max(1, contentHeight - TITLE_HEIGHT)}px;overflow:hidden;`;

  if (setup.printHeadings) {
    wrapper.appendChild(buildHeadings(sheet, target, chunk, rows, columnWidths, scale, headingWidth, headingHeight));
  }

  const tableHost = document.createElement('div');
  tableHost.style.cssText = `position:absolute;left:${headingWidth}px;top:${headingHeight}px;`;
  tableHost.appendChild(buildChunkTable({
    workbook, sheet, target, chunk, columnWidths, scale,
    printGridLines: setup.printGridLines,
    rowStart: rows.start, rowEnd: rows.end, titleRows, repeatTitles: true, ignoreRowSpans: false,
  }));
  wrapper.appendChild(tableHost);

  appendImages(wrapper, sheet, chunk, rows, scale, headerRow, contentHeight - TITLE_HEIGHT);
  page.appendChild(wrapper);
  return page;
}

/** 功能：按行列范围生成表格，合并单元格按页裁剪，跨页部分保留边框与底纹；时间：2026-10-29；作者：lq。 */
function buildChunkTable(params: {
  workbook: WorkbookModel;
  sheet: SheetModel;
  target: CellRange;
  chunk: ColumnChunk;
  columnWidths: number[];
  scale: number;
  printGridLines: boolean;
  rowStart: number;
  rowEnd: number;
  titleRows: { r1: number; r2: number } | null;
  repeatTitles: boolean;
  ignoreRowSpans: boolean;
}): HTMLTableElement {
  const { workbook, sheet, target, chunk, columnWidths, scale, printGridLines, rowStart, rowEnd, titleRows, repeatTitles, ignoreRowSpans } = params;
  const table = document.createElement('table');
  table.style.cssText = `border-collapse:collapse;table-layout:fixed;width:${chunk.width}px;font-size:${workbook.defaultFontSizePt * scale}pt;`;
  const colgroup = document.createElement('colgroup');
  for (let col = chunk.start; col <= chunk.end; col += 1) {
    const colElement = document.createElement('col');
    colElement.style.width = `${(columnWidths[col - target.c1] ?? sheet.defaultColWidth) * scale}px`;
    colgroup.appendChild(colElement);
  }
  table.appendChild(colgroup);

  // 功能：打印标题行不属于正文分页区间，但每页都要重复输出（支持 1:2 这类多行标题）；时间：2026-11-01；作者：lq。
  if (titleRows && repeatTitles) {
    const thead = document.createElement('thead');
    for (let row = titleRows.r1; row <= titleRows.r2; row += 1) {
      thead.appendChild(buildPrintRow({ workbook, sheet, target, chunk, row, header: true, printGridLines, scale, rowStart: titleRows.r1, rowEnd: titleRows.r2, ignoreRowSpans }));
    }
    table.appendChild(thead);
  }
  const tbody = document.createElement('tbody');
  for (let row = rowStart; row <= rowEnd; row += 1) {
    // 功能：标题行已由 thead 输出，正文里跳过，避免同一行打印两次；时间：2026-11-01；作者：lq。
    if (titleRows && row >= titleRows.r1 && row <= titleRows.r2) continue;
    tbody.appendChild(buildPrintRow({ workbook, sheet, target, chunk, row, header: false, printGridLines, scale, rowStart, rowEnd, ignoreRowSpans }));
  }
  table.appendChild(tbody);
  return table;
}

/** 功能：生成一行，按页范围裁剪跨页合并单元格；时间：2026-10-29；作者：lq。 */
function buildPrintRow(params: {
  workbook: WorkbookModel;
  sheet: SheetModel;
  target: CellRange;
  chunk: ColumnChunk;
  row: number;
  header: boolean;
  printGridLines: boolean;
  scale: number;
  rowStart: number;
  rowEnd: number;
  ignoreRowSpans: boolean;
}): HTMLTableRowElement {
  const { workbook, sheet, target, chunk, row, header, printGridLines, scale, rowStart, rowEnd, ignoreRowSpans } = params;
  const tr = document.createElement('tr');
  tr.style.height = `${(sheet.rowHeights[row] ?? sheet.defaultRowHeight) * scale}px`;
  for (let col = chunk.start; col <= chunk.end; col += 1) {
    const merge = sheet.merges.find((item) => item.r1 <= row && row <= item.r2 && item.c1 <= col && col <= item.c2);
    const originRow = merge ? merge.r1 : row;
    const originCol = merge ? merge.c1 : col;
    // 功能：合并单元格只在裁剪后的左上角输出一次，其余位置跳过；时间：2026-10-29；作者：lq。
    const clampRow = Math.max(originRow, rowStart);
    const clampCol = Math.max(originCol, chunk.start);
    if (row !== clampRow || col !== clampCol) continue;
    const rowSpan = ignoreRowSpans ? 1 : Math.min(merge ? merge.r2 : row, rowEnd) - clampRow + 1;
    const colSpan = Math.min(merge ? merge.c2 : col, chunk.end) - clampCol + 1;
    const originInside = !merge || (originRow >= rowStart && originCol >= chunk.start);
    const td = document.createElement('td');
    const cell = cellAt(sheet, originRow, originCol);
    const style = cellStyle(workbook, cell);
    // 功能：跨页延续的合并单元格只保留边框与底纹，文字留在起始页；时间：2026-10-29；作者：lq。
    td.textContent = originInside ? cellDisplay(workbook, sheet, originRow, originCol) : '';
    const horizontal = style.horizontalAlign && style.horizontalAlign !== 'general'
      ? style.horizontalAlign
      : typeof cell?.value === 'number' ? 'right' : 'left';
    if (colSpan > 1) td.colSpan = colSpan;
    if (rowSpan > 1) td.rowSpan = rowSpan;
    td.style.cssText = `border:${printGridLines ? '1px solid #D0D5DD' : 'none'};padding:${2 * scale}px ${5 * scale}px;vertical-align:${
      style.verticalAlign ?? 'bottom'
    };text-align:${horizontal};overflow:hidden;white-space:${style.wrapText ? 'pre-wrap' : 'nowrap'};word-break:break-word;background:${
      style.fill ?? '#FFFFFF'
    };color:${style.color ?? '#1D2939'};font-weight:${header || style.bold ? 700 : 400};font-style:${
      style.italic ? 'italic' : 'normal'
    };font-size:${(style.fontSizePt ?? workbook.defaultFontSizePt) * scale}pt;font-family:${
      style.fontFamily ?? workbook.defaultFont
    };`;
    if (style.borders?.top && style.borders.top.style !== 'none') {
      td.style.borderTop = `${style.borders.top.style === 'medium' ? 2 : style.borders.top.style === 'thick' ? 3 : 1}px solid ${style.borders.top.color ?? '#000'}`;
    }
    if (style.borders?.right && style.borders.right.style !== 'none') {
      td.style.borderRight = `${style.borders.right.style === 'medium' ? 2 : style.borders.right.style === 'thick' ? 3 : 1}px solid ${style.borders.right.color ?? '#000'}`;
    }
    if (style.borders?.bottom && style.borders.bottom.style !== 'none') {
      td.style.borderBottom = `${style.borders.bottom.style === 'medium' ? 2 : style.borders.bottom.style === 'thick' ? 3 : 1}px solid ${style.borders.bottom.color ?? '#000'}`;
    }
    if (style.borders?.left && style.borders.left.style !== 'none') {
      td.style.borderLeft = `${style.borders.left.style === 'medium' ? 2 : style.borders.left.style === 'thick' ? 3 : 1}px solid ${style.borders.left.color ?? '#000'}`;
    }
    tr.appendChild(td);
  }
  return tr;
}

/** 功能：打印行列标题（列标 A/B/C 与行号）；时间：2026-10-29；作者：lq。 */
function buildHeadings(
  sheet: SheetModel,
  target: CellRange,
  chunk: ColumnChunk,
  rows: RowChunk,
  columnWidths: number[],
  scale: number,
  headingWidth: number,
  headingHeight: number,
): HTMLElement {
  const host = document.createElement('div');
  host.style.cssText = `position:absolute;left:0;top:0;width:${headingWidth + chunk.width}px;height:${headingHeight}px;font-size:9px;color:#667085;`;
  const cellStyleText = 'display:flex;align-items:center;justify-content:center;border:1px solid #E4E9F2;box-sizing:border-box;background:#F9FAFB;overflow:hidden;';
  const corner = document.createElement('div');
  corner.style.cssText = `${cellStyleText}position:absolute;left:0;top:0;width:${headingWidth}px;height:${headingHeight}px;`;
  host.appendChild(corner);
  let left = headingWidth;
  for (let col = chunk.start; col <= chunk.end; col += 1) {
    const width = (columnWidths[col - target.c1] ?? sheet.defaultColWidth) * scale;
    const cell = document.createElement('div');
    cell.style.cssText = `${cellStyleText}position:absolute;left:${left}px;top:0;width:${width}px;height:${headingHeight}px;`;
    cell.textContent = columnLabel(col);
    host.appendChild(cell);
    left += width;
  }
  const rowHost = document.createElement('div');
  rowHost.style.cssText = `position:absolute;left:0;top:${headingHeight}px;width:${headingWidth}px;display:flex;flex-direction:column;`;
  for (let row = rows.start; row <= rows.end; row += 1) {
    const cell = document.createElement('div');
    cell.style.cssText = `${cellStyleText}width:100%;height:${(sheet.rowHeights[row] ?? sheet.defaultRowHeight) * scale}px;`;
    cell.textContent = String(row + 1);
    rowHost.appendChild(cell);
  }
  host.appendChild(rowHost);
  return host;
}

/** 功能：把浮动图片按分页位置叠加，跨页图片逐页裁切；时间：2026-10-29；作者：lq。 */
function appendImages(
  wrapper: HTMLElement,
  sheet: SheetModel,
  chunk: ColumnChunk,
  rows: RowChunk,
  scale: number,
  headerRow: number | null,
  bodyHeight: number,
): void {
  const images = sheet.images ?? [];
  if (images.length === 0) return;
  const rowTop = (row: number) => { let total = 0; for (let i = 0; i < row; i += 1) total += sheet.rowHeights[i] ?? sheet.defaultRowHeight; return total; };
  const colLeft = (col: number) => { let total = 0; for (let i = 0; i < col; i += 1) total += sheet.colWidths[i] ?? sheet.defaultColWidth; return total; };
  const headerHeight = headerRow !== null ? (sheet.rowHeights[headerRow] ?? sheet.defaultRowHeight) : 0;
  const chunkWidth = colLeft(chunk.end + 1) - colLeft(chunk.start);
  for (const image of images) {
    const left = colLeft(image.col) + image.offsetX - colLeft(chunk.start);
    if (left + image.width <= 0 || left >= chunkWidth) continue;
    const absoluteTop = rowTop(image.row) + image.offsetY;
    const top = image.row === headerRow ? image.offsetY : headerHeight + absoluteTop - rowTop(rows.start);
    if (top + image.height <= 0 || top * scale >= bodyHeight) continue;
    const element = document.createElement('img');
    element.src = image.src;
    element.alt = image.name;
    element.style.cssText = `position:absolute;left:${left * scale}px;top:${top * scale}px;width:${image.width * scale}px;height:${image.height * scale}px;object-fit:fill;`;
    wrapper.appendChild(element);
  }
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
