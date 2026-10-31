import type { DocBlock, DocModel, DocParagraph, DocSection, DocTable, DocTableRow } from '@mini-office/document-model';
import { paragraphHasContent, tableCellPlacements } from './doc-utils';
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
  /**
   * 功能：可编辑首片在模型段落中的结束偏移；编辑该片段时据此回写，保留被切到续页的尾部；
   * 时间：2026-10-30；作者：lq。
   */
  fragmentEnds?: Record<string, number>;
  /** 功能：批量报告的节内页序与节总页数；时间：2026-10-24；作者：lq。 */
  section?: DocSection;
  sectionPage?: number;
  sectionTotal?: number;
  /** 功能：本页已用内容高度（px），连续分节的后续节据此计算剩余空间；时间：2026-10-30；作者：hou0。 */
  usedHeight?: number;
  /**
   * 功能：页内多栏分区——连续分节在同一页切换栏数时，按分节记录栏参数与所属块，
   * 渲染时同一页的标题（单栏）与三栏段互不干扰；时间：2026-10-30；作者：hou0。
   */
  columnGroups?: ColumnGroup[];
}

/** 功能：页内一个分节的栏参数与块范围；时间：2026-10-30；作者：hou0。 */
export interface ColumnGroup {
  count: number;
  widthsPx?: number[];
  spacesPx?: number[];
  blockIds: string[];
}

/** 功能：表格实测布局，用于按行切页并重复表头；时间：2026-10-29；作者：lq。 */
export interface TableLayout {
  /** 每行渲染高度（px），长度应与 table.rows 一致。 */
  rowHeights: number[];
  /** 连续表头行数，跨页时重复。 */
  headerRows: number;
}

/** 功能：每节独立分页并恢复编号，合并报告重新打开后仍保持各自页码；时间：2026-10-24；作者：lq。 */
/** 功能：段落跨页切分结果：首片（head）留在本页，续片（tail）落到下一页；时间：2026-10-30；作者：lq。 */
export interface ParagraphSplit {
  /** 本页内的前半段（保留原段落 id 可编辑，或续页时改用独立 id 只读）。 */
  head: DocParagraph;
  /** 溢出到下一页的后半段（续页片段，已去掉列表符号与首行缩进）。 */
  tail: DocParagraph;
  /** 前半段实测高度（含段前/段后间距的外边距盒）。 */
  headHeight: number;
  /** 后半段实测高度（已按续页样式，去掉段前/段后间距）。 */
  tailHeight: number;
  /** 功能：首片在模型段落中的结束偏移（模型字符单位）；时间：2026-10-30；作者：lq。 */
  headEnd: number;
  /**
   * 功能：断点来自段内分页符时为真——续片必须另起一页，即使当前页还有空间；
   * 时间：2026-10-30；作者：lq。
   */
  forceNewPage?: boolean;
}

/**
 * 功能：段落断页器——由宿主提供，因为它依赖真实 DOM 测量；
 * needsBreak 用于判断段落是否含必须断页的标记（段内分页符），split 负责算出断点与实测高度；
 * 时间：2026-10-30；作者：lq。
 */
export interface ParagraphBreaker {
  /** 功能：段落是否含强制断点（段内分页符）；时间：2026-10-30；作者：lq。 */
  needsBreak(paragraph: DocParagraph): boolean;
  /** 功能：计算断点；返回 null 表示无需断开；时间：2026-10-30；作者：lq。 */
  split(paragraph: DocParagraph, maxHeight: number): ParagraphSplit | null;
}

export function paginateDocument(
  doc: DocModel,
  heights: Map<string, number>,
  tables?: Map<string, TableLayout>,
  breaker?: ParagraphBreaker,
): PageSlice[] {
  const sections = documentSections(doc);
  const pages: PageSlice[] = [];
  let nextNumber = 1;
  for (const section of sections) {
    // 功能：documentSections 展开的节携带首页/奇偶页页眉页脚变体；时间：2026-10-31；作者：hou0。
    const sectionVariants = section as DocModel & DocSection;
    const metrics = pageMetrics(section.page);
    // 功能：连续分节不翻页——把本节内容装进上一页的剩余空间，装不下的部分才另起新页；
    // 同一页内出现多栏分节时按分节记录栏分组，标题（单栏）与三栏段各自排版；时间：2026-10-30；作者：hou0。
    const previous = pages.at(-1);
    const remaining = section.page.type === 'continuous' && previous
      ? Math.max(0, metrics.contentHeight - (previous.usedHeight ?? 0))
      : undefined;
    if (remaining !== undefined && remaining > 1) {
      const slices = paginateBlocks(section.blocks, heights, metrics.contentHeight, tables, breaker, remaining);
      const [first, ...rest] = slices;
      if (first) {
        previous!.blocks = [...previous!.blocks, ...first.blocks];
        if (first.fragments) previous!.fragments = [...(previous!.fragments ?? []), ...first.fragments];
        if (first.locked) previous!.locked = [...(previous!.locked ?? []), ...first.locked];
        if (first.fragmentEnds) previous!.fragmentEnds = { ...(previous!.fragmentEnds ?? {}), ...first.fragmentEnds };
        if (section.page.columns && section.page.columns > 1) {
          previous!.columnGroups = [
            ...(previous!.columnGroups ?? []),
            {
              count: section.page.columns,
              widthsPx: section.page.columnWidthsPx,
              spacesPx: section.page.columnSpacePx,
              blockIds: first.blocks.map((block) => block.id),
            },
          ];
        }
        previous!.usedHeight = (previous!.usedHeight ?? 0) + (first.usedHeight ?? 0);
      }
      const start = previous?.section?.page.pageNumberStart ?? nextNumber;
      rest.forEach((slice, offset) => {
        pages.push({
          ...slice,
          index: pages.length,
          section: { page: { ...section.page, pageNumberStart: start }, header: section.header, footer: section.footer, headerFirst: sectionVariants.headerFirst, headerEven: sectionVariants.headerEven, footerFirst: sectionVariants.footerFirst, footerEven: sectionVariants.footerEven },
          sectionPage: (previous?.sectionPage ?? 0) + offset + 1,
          sectionTotal: (previous?.sectionTotal ?? 0) + rest.length,
        });
      });
      // 功能：接续节溢出的页也占用页码，避免后续节从这里重新编号；时间：2026-10-30；作者：hou0。
      nextNumber = start + (previous?.sectionPage ?? 0) + rest.length + 1;
      continue;
    }
    const slices = paginateBlocks(section.blocks, heights, metrics.contentHeight, tables, breaker);
    const start = section.page.pageNumberStart ?? nextNumber;
    slices.forEach((slice, index) => pages.push({ ...slice, index: pages.length, section: { page: { ...section.page, pageNumberStart: start }, header: section.header, footer: section.footer, headerFirst: sectionVariants.headerFirst, headerEven: sectionVariants.headerEven, footerFirst: sectionVariants.footerFirst, footerEven: sectionVariants.footerEven }, sectionPage: index, sectionTotal: slices.length }));
    nextNumber = start + slices.length;
  }
  return pages;
}

/**
 * 功能：keepNext 链检查——把 index 处的块放到当前页 `used` 高度之后，是否会因「与下段同页」
 * 的连带要求而放不下后续块；是则该块应推到新页。链长超过整页时放弃推页避免死循环；
 * 时间：2026-10-31；作者：hou0。
 */
function keepNextChainOverflows(blocks: DocBlock[], index: number, heights: Map<string, number>, used: number, limit: number): boolean {
  let sum = 0;
  let k = index;
  while (k < blocks.length) {
    const block = blocks[k];
    sum += Math.max(0, heights.get(block.id) ?? 0);
    const keeps = block.type === 'paragraph' && block.style.keepNext === true && k + 1 < blocks.length;
    if (!keeps) return false;
    if (sum > limit) return false;
    k += 1;
    const nextHeight = Math.max(0, heights.get(blocks[k].id) ?? 0);
    if (used + sum + nextHeight > limit + 0.5) return true;
  }
  return false;
}

/**
 * 功能：孤寡控制（Word 默认开启）——跨页断点两侧至少保留两行：
 * 头部只剩一行时整段移到下页（返回 null）；尾部只剩一行时把断点前移凑足两行；
 * 时间：2026-10-31；作者：hou0。
 */
export interface WidowControlInput {
  headEnd: number;
  /** 全段包装高度（含段前/段后 padding，与 measure 同口径）。 */
  totalHeight: number;
  /** 测量任意模型偏移处的前缀包装高度（含 padding）。 */
  measure: (offset: number) => number;
  /** 段落行高（CSS px）。 */
  lineHeightPx: number;
  /** 段前+段后 padding 合计。 */
  padY: number;
  /** 本页剩余可用高度（与 measure 同口径）。 */
  maxHeight: number;
  /** 段内分页符强制断点：不参与孤寡调整。 */
  forced?: boolean;
}

export function applyWidowControl(input: WidowControlInput): number | null {
  const { headEnd, totalHeight, measure, lineHeightPx, padY, maxHeight, forced } = input;
  if (forced || lineHeightPx <= 0) return headEnd;
  const content = (wrapperHeight: number) => Math.max(0, wrapperHeight - padY);
  const headLines = Math.max(1, Math.round(content(measure(headEnd)) / lineHeightPx));
  const totalLines = Math.max(1, Math.round(content(totalHeight) / lineHeightPx));
  if (totalLines < 3) return headEnd;
  if (headLines < 2) return null;
  let lo = 0;
  let hi = headEnd;
  let best = -1;
  while (lo <= hi) {
    const mid = Math.floor((lo + hi) / 2);
    const wrapperHeight = measure(mid);
    if (wrapperHeight <= maxHeight + 0.5 && totalHeight - wrapperHeight >= 2 * lineHeightPx - 0.5) {
      best = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  if (best < 0) return null;
  return best < headEnd ? best : headEnd;
}

/**
 * 功能：按实测高度把内容装进定高页面；长表格按行切分并重复表头，避免整表被裁掉丢行；
 * 长段落按行断页（首片可编辑、续片只读），段内分页符作为强制断点；
 * 时间：2026-10-29；作者：lq。
 */
export function paginateBlocks(
  blocks: DocBlock[],
  heights: Map<string, number>,
  contentHeight: number,
  tables?: Map<string, TableLayout>,
  breaker?: ParagraphBreaker,
  /** 功能：首页可用高度——连续分节接入上一页剩余空间时小于整页高；时间：2026-10-30；作者：hou0。 */
  firstPageBudget?: number,
): PageSlice[] {
  const pages: PageSlice[] = [];
  let current: DocBlock[] = [];
  let fragments: string[] = [];
  let locked: string[] = [];
  let fragmentEnds: Record<string, number> = {};
  let used = 0;

  const flush = () => {
    if (current.length === 0) return;
    pages.push({
      index: pages.length,
      blocks: current,
      fragments: fragments.length ? fragments : undefined,
      locked: locked.length ? locked : undefined,
      fragmentEnds: Object.keys(fragmentEnds).length ? fragmentEnds : undefined,
      usedHeight: used,
    });
    current = [];
    fragments = [];
    locked = [];
    fragmentEnds = {};
    used = 0;
  };

  // 功能：跨页切分产生的续页片段；分轮次消费，避免跳过一个原始块；时间：2026-10-30；作者：lq。
  let pending: { block: DocBlock; isCont: boolean; forceNewPage: boolean } | null = null;
  let i = 0;
  while (i < blocks.length) {
    // 功能：首页可用高度可小于整页（连续分节接续上一页）；后续页恢复整页高；时间：2026-10-30；作者：hou0。
    const limit = pages.length === 0 && firstPageBudget !== undefined ? firstPageBudget : contentHeight;
    const item: { block: DocBlock; isCont: boolean; forceNewPage: boolean } = pending ?? { block: blocks[i], isCont: false, forceNewPage: false };
    pending = null;
    const block: DocBlock = item.block;
    const cont = item.isCont;
    if (block.type === 'pageBreak') {
      flush();
      i += 1;
      continue;
    }
    // 功能：段内分页符断出的续片必须另起一页，即使当前页还有空间；时间：2026-10-30；作者：lq。
    if (item.forceNewPage && current.length > 0) flush();
    if (block.type === 'paragraph' && block.style.pageBreakBefore && current.length > 0) {
      flush();
    }
    const height = Math.max(0, heights.get(block.id) ?? 0);

    if (block.type === 'paragraph') {
      const fits = used + height <= limit + 0.5;
      // 功能：段内分页符是强制断点，即使整段放得下也必须在此断开；时间：2026-10-30；作者：lq。
      const mustBreak = breaker?.needsBreak(block) === true;
      // 功能：keepNext 链——本块放得下但「与下段同页」的连带要求放不下时，连本块一起推到新页；时间：2026-10-31；作者：hou0。
      const keepNextPush =
        fits && !mustBreak && current.length > 0 && i + 1 < blocks.length &&
        block.style.keepNext === true && keepNextChainOverflows(blocks, i, heights, used, limit);
      if (keepNextPush) flush();
      if (fits && !mustBreak) {
        current.push(block);
        if (cont) {
          // 功能：续页片段只读，不可编辑，避免编辑落到错误位置；时间：2026-10-30；作者：lq。
          fragments.push(block.id);
          locked.push(block.id);
        }
        used += height;
        if (used > limit + 0.5) flush();
        i += 1;
        continue;
      }
      // 功能：段高超过本页剩余空间——先换页，再按行切分；强制断点则留在本页；时间：2026-10-30；作者：lq。
      if (!fits && current.length > 0) flush();
      const remaining = limit - used;
      const split = breaker?.split(block, remaining) ?? null;
      if (split) {
        // 功能：首片为空表示此处只断页（段首分页符），不产出多余空行；时间：2026-10-30；作者：lq。
        if (paragraphHasContent(split.head)) {
          if (cont) {
            const headId = `${block.id}-head`;
            const head = { ...split.head, id: headId };
            current.push(head);
            fragments.push(headId);
            locked.push(headId);
          } else {
            // 功能：首片沿用原段落 id 与样式（含列表符号、首行缩进），保持可编辑；时间：2026-10-30；作者：lq。
            current.push(split.head);
            fragments.push(split.head.id);
            fragmentEnds[split.head.id] = split.headEnd;
          }
          used += split.headHeight;
        } else {
          flush();
        }
        const tailId: string = `${block.id}-cont`;
        const tail: DocBlock = { ...split.tail, id: tailId };
        heights.set(tailId, split.tailHeight);
        pending = { block: tail, isCont: true, forceNewPage: split.forceNewPage === true };
        continue;
      }
      // 功能：无法切分（如单行即超页）时整段留到下一页；时间：2026-10-30；作者：lq。
      if (current.length > 0) flush();
      current.push(block);
      if (cont) {
        fragments.push(block.id);
        locked.push(block.id);
      }
      used += height;
      if (used > limit + 0.5) flush();
      i += 1;
      continue;
    }

    // 功能：表格：整表放得下时保持原样，避免无意义的切片；时间：2026-10-29；作者：lq。
    const table = block.type === 'table' ? block : undefined;
    const layout = table ? tables?.get(table.id) : undefined;
    if (!table || !layout || used + height <= limit + 0.5) {
      if (current.length > 0 && used + height > limit + 0.5) flush();
      current.push(block);
      used += height;
      if (used > limit + 0.5) flush();
      i += 1;
      continue;
    }
    // 功能：长表格逐行装页，续页重复表头，行不会被裁掉；时间：2026-10-29；作者：lq。
    let part = 0;
    let start = 0;
    let continuation = false;
    while (start < table.rows.length) {
      const headerRows = continuation ? layout.headerRows : 0;
      const headerHeight = sumHeights(layout.rowHeights, 0, headerRows);
      let budget = limit - used - headerHeight;
      if (budget <= 0) {
        if (current.length > 0) { flush(); continue; }
        // 功能：表头本身高于整页时仍允许输出，避免死循环；时间：2026-10-29；作者：lq。
        budget = Math.max(1, limit - headerHeight);
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
    i += 1;
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
