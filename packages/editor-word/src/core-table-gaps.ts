/** 功能：分页装饰遵守表格的纵向合并网格，重复表头不新增列且不写入正文；时间：2026-11-11；作者：lq。 */
import { outerTableRows } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-lines';

/** 功能：计算页缝行中未被前面合并格占用的连续列段，支持中间列纵向合并；时间：2026-11-11；作者：lq。 */
export function wordTableGapSpans(table: HTMLElement, rowIndex: number): number[] {
  const rows = outerTableRows(table) as HTMLTableRowElement[];
  const busy: number[] = [];
  let width = 0;
  for (let index = 0; index < rows.length; index++) {
    let column = 0;
    for (const cell of Array.from(rows[index].cells)) {
      while ((busy[column] ?? 0) > index) column++;
      const colspan = cell.colSpan || 1;
      const rowspan = Number(cell.dataset.baseRowspan) || cell.rowSpan || 1;
      for (let col = column; col < column + colspan; col++) busy[col] = index + rowspan;
      column += colspan; width = Math.max(width, column);
    }
  }
  // 功能：第二遍只登记切点前的格子，切点所在行本身不能占用插入的装饰行；时间：2026-11-11；作者：lq。
  busy.fill(0);
  for (let index = 0; index < rowIndex; index++) {
    let column = 0;
    for (const cell of Array.from(rows[index]?.cells ?? [])) {
      while ((busy[column] ?? 0) > index) column++;
      for (let col = column; col < column + cell.colSpan; col++) busy[col] = index + (Number(cell.dataset.baseRowspan) || cell.rowSpan);
      column += cell.colSpan;
    }
  }
  const spans: number[] = []; let run = 0;
  for (let col = 0; col < width; col++) {
    if ((busy[col] ?? 0) > rowIndex) { if (run) spans.push(run); run = 0; }
    else run++;
  }
  if (run) spans.push(run);
  return spans;
}

/** 功能：重复表头沿原表真实列边界绘制，绝对定位克隆格避免与贯穿页缝的 rowspan 争用列；时间：2026-11-11；作者：lq。 */
export function wordRepeatedTableHeaders(table: HTMLElement, height: number, factor: number): { rows: HTMLElement[]; height: number } {
  const headers: HTMLElement[] = [];
  let total = 0;
  const tableBox = table.getBoundingClientRect();
  for (const source of outerTableRows(table) as HTMLTableRowElement[]) {
    if (total >= height - 1.5) break;
    const sourceBox = source.getBoundingClientRect(), rowHeight = sourceBox.height / factor;
    const clone = source.cloneNode(true) as HTMLTableRowElement;
    clone.classList.add('page-gap-inline', 'page-repeat-header', 'mo-word-repeat-header');
    clone.contentEditable = 'false'; clone.style.height = `${rowHeight}px`;
    Array.from(clone.cells).forEach((cell, index) => {
      const box = source.cells[index].getBoundingClientRect();
      cell.removeAttribute('rowspan'); cell.removeAttribute('colspan'); cell.removeAttribute('data-base-rowspan');
      cell.style.left = `${(box.left - tableBox.left) / factor}px`;
      cell.style.width = `${box.width / factor}px`; cell.style.height = `${rowHeight}px`;
    });
    headers.push(clone); total += rowHeight;
  }
  return { rows: headers, height: total };
}
