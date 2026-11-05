import type { DocBlock, DocModel, DocSection, DocTable, DocTableRow } from '@mini-office/document-model';
import { tableCellPlacements } from './doc-utils';
import { documentSections } from './sections';

export interface PageSlice {
  index: number;
  blocks: DocBlock[];
  /**
   * 功能：本页中的表格跨页片段 id；这些 block 是切片而不是原表，渲染时不能被原表映射覆盖；
   * 时间：2026-10-29；作者：lq。
   */
  fragments?: string[];
  /**
   * 功能：续页片段 id；同一张表只在首片就地编辑，续页只读以免编辑落到错误的行上；
   * 时间：2026-10-29；作者：lq。
   */
  locked?: string[];
  /** 功能：批量报告的节内页序与节总页数；时间：2026-10-24；作者：lq。 */
  section?: DocSection;
  sectionPage?: number;
  sectionTotal?: number;
  /**
   * 功能：本页中跨页段落的裁剪窗口，键为片段 block id；时间：2026-11-03；作者：lq。
   *
   * 段落不像表格可以真正切片（切开会破坏可编辑文本模型），因此改为「整段渲染 + 窗口裁剪」：
   * 片段仍携带完整段落内容，渲染时用固定高度 + 负上边距只露出属于本页的那几行。
   */
  splits?: Record<string, ParagraphSplit>;
}

/** 功能：表格实测布局，用于按行切页并重复表头；时间：2026-10-29；作者：lq。 */
export interface TableLayout {
  /** 每行渲染高度（px），长度应与 table.rows 一致。 */
  rowHeights: number[];
  /** 连续表头行数，跨页时重复。 */
  headerRows: number;
}

/**
 * 功能：段落行盒布局，用于按行切分跨页段落；时间：2026-11-03；作者：lq。
 *
 * 行盒由渲染时的确定行高决定，因此只需记录每行顶部偏移与内容总高度。
 */
export interface ParagraphLayout {
  /** 每行相对段落文本框顶部的起始偏移（px），首行为 0。 */
  lineTops: number[];
  /** 段落内容总高度（px），等于行数 × 行高。 */
  height: number;
}

/** 功能：一个段落片段的裁剪窗口参数；时间：2026-11-03；作者：lq。 */
export interface ParagraphSplit {
  /** 被切分的原段落 id。 */
  sourceId: string;
  /** 起始行（含）。 */
  start: number;
  /** 结束行（不含）。 */
  end: number;
  /** 段落文本框顶到起始行的偏移（px），渲染时用负 margin-top 上移。 */
  offset: number;
  /** 可视窗口高度（px），不含段前段后距。 */
  height: number;
  /** 本片段保留的段前距（px）；续页与页首为 0。 */
  paddingTop: number;
  /** 本片段保留的段后距（px）；只有整段最后一片保留。 */
  paddingBottom: number;
}

/** 功能：每节独立分页并恢复编号，合并报告重新打开后仍保持各自页码；时间：2026-10-24；作者：lq。 */
export function paginateDocument(
  doc: DocModel,
  heights: Map<string, number>,
  tables?: Map<string, TableLayout>,
  paragraphs?: Map<string, ParagraphLayout>,
): PageSlice[] {
  const sections = documentSections(doc);
  const pages: PageSlice[] = [];
  let nextNumber = 1;
  for (const section of sections) {
    const slices = paginateBlocks(section.blocks, heights, pageMetrics(section.page).contentHeight, tables, paragraphs);
    const start = section.page.pageNumberStart ?? nextNumber;
    slices.forEach((slice, index) => pages.push({ ...slice, index: pages.length, section: { page: { ...section.page, pageNumberStart: start }, header: section.header, footer: section.footer }, sectionPage: index, sectionTotal: slices.length }));
    nextNumber = start + slices.length;
  }
  return pages;
}

/**
 * 功能：按实测高度把内容装进定高页面；长表格按行切分并重复表头，避免整表被裁掉丢行；
 * 时间：2026-10-29；作者：lq。
 */
export function paginateBlocks(
  blocks: DocBlock[],
  heights: Map<string, number>,
  contentHeight: number,
  tables?: Map<string, TableLayout>,
  paragraphs?: Map<string, ParagraphLayout>,
): PageSlice[] {
  const pages: PageSlice[] = [];
  let current: DocBlock[] = [];
  let fragments: string[] = [];
  let locked: string[] = [];
  let splits: Record<string, ParagraphSplit> = {};
  let used = 0;

  const flush = () => {
    if (current.length === 0) return;
    pages.push({
      index: pages.length,
      blocks: current,
      fragments: fragments.length ? fragments : undefined,
      locked: locked.length ? locked : undefined,
      splits: Object.keys(splits).length ? splits : undefined,
    });
    current = [];
    fragments = [];
    locked = [];
    splits = {};
    used = 0;
  };

  // 功能：页首不叠加段前距，与渲染时丢弃 padding-top 的行为一致；时间：2026-11-02；作者：lq。
  const heightOf = (block: DocBlock, firstOnPage: boolean): number => {
    const raw = Math.max(0, heights.get(block.id) ?? 0);
    if (!firstOnPage || block.type !== 'paragraph') return raw;
    return Math.max(0, raw - (block.style.spaceBeforePx ?? 0));
  };

  /**
   * 功能：收集 keepNext 链——标题与后续段落必须同页；时间：2026-11-02；作者：lq。
   *
   * 只把段落计入链，遇到表格或分页符即截止，避免与表格跨页切片逻辑互相干扰。
   */
  const keepChain = (start: number): DocBlock[] => {
    const chain: DocBlock[] = [blocks[start]];
    let index = start;
    while (index < blocks.length) {
      const node = blocks[index];
      if (node.type !== 'paragraph' || !node.style.keepNext) break;
      const next = blocks[index + 1];
      if (!next || next.type === 'pageBreak' || next.type === 'table') break;
      chain.push(next);
      index += 1;
    }
    return chain;
  };

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    if (block.type === 'pageBreak') {
      // 连续/段首硬分页有实际空白页语义，不能与一个分页符折叠。
      if (current.length === 0) pages.push({ index: pages.length, blocks: [] });
      else flush();
      continue;
    }
    if (block.type === 'paragraph' && block.style.pageBreakBefore && current.length > 0) {
      flush();
    }
    // 功能：keepNext 链整体放不下时先换页，避免标题落在页尾、正文翻到下一页；时间：2026-11-02；作者：lq。
    if (block.type === 'paragraph' && block.style.keepNext && current.length > 0) {
      const chain = keepChain(index);
      if (chain.length > 1) {
        const chainHeight = chain.reduce((sum, item) => sum + Math.max(0, heights.get(item.id) ?? 0), 0);
        if (chainHeight <= contentHeight + 0.5 && used + chainHeight > contentHeight + 0.5) flush();
      }
    }
    // 功能：段落行级切分——整段放不下时按行盒切开，避免整段移到下一页在页底留下半页空白；时间：2026-11-03；作者：lq。
    const paragraphLayout = block.type === 'paragraph' && !block.toc ? paragraphs?.get(block.id) : undefined;
    if (block.type === 'paragraph' && paragraphLayout && paragraphLayout.lineTops.length > 1) {
      const spaceBefore = block.style.spaceBeforePx ?? 0;
      const spaceAfter = block.style.spaceAfterPx ?? 0;
      const total = paragraphLayout.lineTops.length;
      const atTop = current.length === 0;
      const lead = atTop ? 0 : spaceBefore;
      if (used + lead + paragraphLayout.height <= contentHeight + 0.5) {
        current.push(block);
        // 功能：多行段落仍包含段后距，防止下一段被误放入当前页；时间：2026-11-03；作者：lq。
        used += lead + paragraphLayout.height + spaceAfter;
        continue;
      }
      // 功能：整段不拆页（w:keepLines）时，只要能放进一整页就整体移到下一页；时间：2026-11-03；作者：lq。
      if (block.style.keepLines && !atTop && paragraphLayout.height <= contentHeight + 0.5) {
        flush();
        current.push(block);
        used = paragraphLayout.height + spaceAfter;
        continue;
      }
      let start = 0;
      let part = 0;
      while (start < total) {
        const pageTop = current.length === 0;
        const padTop = pageTop ? 0 : spaceBefore;
        const base = paragraphLayout.lineTops[start];
        const budget = contentHeight - used - padTop;
        let end = start;
        while (end < total) {
          const nextTop = end + 1 < total ? paragraphLayout.lineTops[end + 1] : paragraphLayout.height;
          if (nextTop - base > budget + 0.5) break;
          end += 1;
        }
        if (end === start) {
          // 本页一行都放不下：换页重算；页首仍放不下时至少放一行，避免死循环。
          if (!pageTop) {
            flush();
            continue;
          }
          end = start + 1;
        }
        // 功能：孤行控制（w:widowControl，默认开启）——页尾不留单行，续页开头也不留单行；时间：2026-11-03；作者：lq。
        if (block.style.widowControl !== false && end < total) {
          if (end - start === 1 && !pageTop) {
            flush();
            continue;
          }
          if (total - end === 1 && end - start >= 2) {
            end -= 1;
            if (end - start === 1 && !pageTop) {
              flush();
              continue;
            }
          }
        }
        const last = end >= total;
        // 功能：换页后整段正好放下时按普通段落处理，不生成无意义的裁剪窗口；时间：2026-11-03；作者：lq。
        if (part === 0 && start === 0 && last) {
          current.push(block);
          used += padTop + paragraphLayout.height + spaceAfter;
          break;
        }
        const offset = paragraphLayout.lineTops[start];
        const windowHeight = (last ? paragraphLayout.height : paragraphLayout.lineTops[end]) - offset;
        const fragmentId = part === 0 ? block.id : `${block.id}-part-${part}`;
        splits[fragmentId] = {
          sourceId: block.id,
          start,
          end,
          offset,
          height: windowHeight,
          paddingTop: padTop,
          paddingBottom: last ? spaceAfter : 0,
        };
        current.push({ ...block, id: fragmentId });
        fragments.push(fragmentId);
        // 功能：续页片段只读，避免编辑落到被裁剪的不可见行上；时间：2026-11-03；作者：lq。
        if (part > 0) locked.push(fragmentId);
        used += padTop + windowHeight + (last ? spaceAfter : 0);
        part += 1;
        start = end;
        if (start < total) flush();
      }
      continue;
    }
    const table = block.type === 'table' ? block : undefined;
    const layout = table ? tables?.get(table.id) : undefined;
    const height = heightOf(block, current.length === 0);
    // 功能：整表放得下时保持原样，避免无意义的切片；时间：2026-10-29；作者：lq。
    if (!table || !layout || used + height <= contentHeight + 0.5) {
      if (current.length > 0 && used + height > contentHeight + 0.5) flush();
      current.push(block);
      used += heightOf(block, current.length === 1);
      if (used > contentHeight + 0.5) flush();
      continue;
    }
    // 功能：长表格逐行装页，续页重复表头，行不会被裁掉；时间：2026-10-29；作者：lq。
    let part = 0;
    let start = 0;
    let continuation = false;
    while (start < table.rows.length) {
      const headerRows = continuation ? layout.headerRows : 0;
      const headerHeight = sumHeights(layout.rowHeights, 0, headerRows);
      let budget = contentHeight - used - headerHeight;
      if (budget <= 0) {
        if (current.length > 0) { flush(); continue; }
        // 功能：表头本身高于整页时仍允许输出，避免死循环；时间：2026-10-29；作者：lq。
        budget = Math.max(1, contentHeight - headerHeight);
      }
      let end = start;
      let acc = 0;
      while (end < table.rows.length) {
        const rowHeight = layout.rowHeights[end] ?? 0;
        if (end > start && acc + rowHeight > budget + 0.5) break;
        acc += rowHeight;
        end += 1;
      }
      // 功能：单行超过整页高度时仍独占一页，宁可轻微溢出也不丢行；时间：2026-10-29；作者：lq。
      if (end === start) end = start + 1;
      // 功能：首片沿用原表 id，保证点击表格仍能编辑原表；续页用独立 id 并锁定编辑；时间：2026-10-29；作者：lq。
      const fragmentId = part === 0 ? table.id : `${table.id}-part-${part}`;
      const fragment = sliceTable(table, start, end - 1, headerRows);
      current.push({ ...fragment, id: fragmentId });
      fragments.push(fragmentId);
      if (part > 0) locked.push(fragmentId);
      part += 1;
      used += headerHeight + sumHeights(layout.rowHeights, start, end);
      start = end;
      continuation = true;
      if (start < table.rows.length) flush();
    }
  }
  flush();
  if (pages.length === 0) pages.push({ index: 0, blocks: [] });
  return pages;
}

/** 功能：累加区间行高；时间：2026-10-29；作者：lq。 */
function sumHeights(heights: number[], start: number, end: number): number {
  let total = 0;
  for (let index = start; index < end; index += 1) total += heights[index] ?? 0;
  return total;
}

/**
 * 功能：截取表格的行区间，续页重复表头，跨页合并单元格按页裁剪并保留边框底纹；
 * 时间：2026-10-29；作者：lq。
 */
export function sliceTable(table: DocTable, start: number, end: number, repeatHeader: number): DocTable {
  const placements = tableCellPlacements(table);
  const byRow = new Map<number, Array<{ col: number; cell: DocTableRow['cells'][number] }>>();
  for (let row = start; row <= end; row += 1) byRow.set(row, []);
  for (const placement of placements) {
    const placementEnd = placement.row + placement.cell.rowSpan - 1;
    if (placement.row > end || placementEnd < start) continue;
    const rowStart = Math.max(placement.row, start);
    const rowEnd = Math.min(placementEnd, end);
    const cell = { ...placement.cell, rowSpan: rowEnd - rowStart + 1 };
    // 功能：从上一页延续下来的合并单元格只保留边框与底纹，正文留在起始页；时间：2026-10-29；作者：lq。
    if (placement.row < start) cell.blocks = [];
    byRow.get(rowStart)!.push({ col: placement.col, cell });
  }
  const rows: DocTableRow[] = [];
  for (let row = start; row <= end; row += 1) {
    const source = table.rows[row];
    const cells = (byRow.get(row) ?? []).sort((a, b) => a.col - b.col).map((item) => item.cell);
    rows.push({ ...source, cells });
  }
  const headers = repeatHeader > 0
    ? table.rows.slice(0, repeatHeader).map((row) => ({ ...row, cells: row.cells.map((cell) => ({ ...cell })) }))
    : [];
  return { ...table, rows: [...headers, ...rows] };
}

export interface PageMetrics {
  contentWidth: number;
  contentHeight: number;
}

export function pageMetrics(page: {
  widthPx: number;
  heightPx: number;
  margins: { top: number; right: number; bottom: number; left: number };
}): PageMetrics {
  return {
    contentWidth: Math.max(120, page.widthPx - page.margins.left - page.margins.right),
    contentHeight: Math.max(80, page.heightPx - page.margins.top - page.margins.bottom),
  };
}
