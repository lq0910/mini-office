/** 功能：解析真实换页行，区分正文行和表格末尾无文字留白；时间：2026-11-11；作者：lq。 */
import { lineStartAnchor, nextLineAnchor, outerTableRows, type DomLineRectsFn, type LineAnchor } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-lines';

/** 功能：表格切点只在计划行及覆盖该行的合并格中按视觉顺序选锚点，避免首列下一行抢占后列续行；时间：2026-11-11；作者：lq。 */
export function wordTableRowRects(element: HTMLElement, rowIndex: number, rectsOf: DomLineRectsFn): DomLineRectsFn {
  const rows = outerTableRows(element) as HTMLTableRowElement[];
  const cells = new Set<HTMLTableCellElement>();
  for (let index = 0; index <= rowIndex; index++) for (const cell of Array.from(rows[index]?.cells ?? [])) {
    if (index + (Number(cell.dataset.baseRowspan) || cell.rowSpan) > rowIndex) cells.add(cell);
  }
  return (host, factor) => rectsOf(host, factor).filter(line => {
    // 功能：嵌套表格中的行仍归属外层计划格，保留原有嵌套切点覆盖路径；时间：2026-11-11；作者：lq。
    let parent: Element | null = line.node instanceof Element ? line.node : line.node.parentElement;
    while (parent && parent !== element) {
      if (cells.has(parent as HTMLTableCellElement)) return true;
      parent = parent.parentElement;
    }
    return false;
  }).sort((a, b) => a.offset - b.offset || a.left - b.left);
}

/** 功能：仅在真实末行之后使用块末；切点落入最后字形时将完整末行移到下一页，禁止依赖像素容差丢失文字；时间：2026-11-11；作者：lq。 */
export function resolveWordPageAnchor(element: HTMLElement, offset: number, height: number, factor: number, rectsOf: DomLineRectsFn): (LineAnchor & { layoutOffset?: number }) | 'after-block' | null {
  if (!Number.isFinite(offset) || offset < 0 || offset > height + 1.5) return null;
  const anchor = lineStartAnchor(element, offset, factor, rectsOf) ?? nextLineAnchor(element, offset, factor, rectsOf);
  if (anchor) return anchor;
  const lines = rectsOf(element, factor);
  if (lines.length && lines.every(line => line.bottom <= offset)) return 'after-block';
  // 功能：表格行高补偿与字形墨迹高度不是同一几何量；有实际文字穿过切点时采用可定位的末行开头；时间：2026-11-11；作者：lq。
  const crossing = [...lines].reverse().find(line => line.offset < offset && line.bottom > offset);
  if (crossing) {
    const start = lineStartAnchor(element, crossing.offset, factor, rectsOf);
    if (start) return { ...start, layoutOffset: crossing.offset };
  }
  return null;
}
