/** 功能：验证 Word / Excel 打印分页的页数、跨页切分与纸张规则；时间：2026-10-29；作者：lq。 */
import { describe, expect, it } from 'vitest';
import { createBlankDoc, createBlankWorkbook, type DocTable, type ParagraphStyle } from '@mini-office/document-model';
import { createPrintRoot, renderPrintPreview, resolvePrintPaper } from '../packages/platform-adapter/src/print';
import { paginateBlocks, sliceTable, type ParagraphBreaker } from '../packages/editor-word/src/pagination';
import { makeParagraph, mergeFragmentRuns, splitRunsAtOffset } from '../packages/editor-word/src/doc-utils';
import { runsPlainText } from '../packages/editor-word/src/editing';
import { buildSheetPrintPages, sheetPageSize } from '../packages/editor-sheet/src/print';
import { defaultSheetPageSetup } from '@mini-office/document-model';

/**
 * 功能：模拟真实断页器——每字 20px，按本页容量切分；段内分页符作为强制断点并让续片另起一页；
 * 时间：2026-10-30；作者：lq。
 */
function capacityBreaker(): ParagraphBreaker {
  const perChar = 20;
  const continuation = (style: ParagraphStyle) => ({ ...style, firstLineIndentPx: 0, spaceBeforePx: 0, spaceAfterPx: 0 });
  return {
    needsBreak: (paragraph) => paragraph.runs.some((run) => run.pageBreak),
    split: (paragraph, maxHeight) => {
      const total = runsPlainText(paragraph.runs).length;
      let forced = -1;
      let acc = 0;
      for (const run of paragraph.runs) {
        if (run.pageBreak) { forced = acc; break; }
        acc += run.image || run.tab || run.lineBreak || run.mergeField ? 1 : (run.text?.length ?? 0);
      }
      const capacity = Math.floor(maxHeight / perChar);
      // 功能：强制断点在本页放得下时优先按它断开；时间：2026-10-30；作者：lq。
      if (forced >= 0 && forced < total && forced <= capacity) {
        const [headRuns] = splitRunsAtOffset(paragraph.runs, forced, {});
        const [, tailRuns] = splitRunsAtOffset(paragraph.runs, forced + 1, {});
        return {
          head: { ...paragraph, runs: headRuns },
          tail: { ...paragraph, runs: tailRuns, style: continuation(paragraph.style), list: undefined },
          headEnd: forced,
          forceNewPage: true,
          headHeight: forced * perChar,
          tailHeight: (total - forced - 1) * perChar,
        };
      }
      if (capacity >= total) return null;
      const offset = Math.max(1, capacity);
      const [headRuns, tailRuns] = splitRunsAtOffset(paragraph.runs, offset, {});
      return {
        head: { ...paragraph, runs: headRuns },
        tail: { ...paragraph, runs: tailRuns, style: continuation(paragraph.style), list: undefined },
        headEnd: offset,
        headHeight: offset * perChar,
        tailHeight: (total - offset) * perChar,
      };
    },
  };
}

/** 功能：构造带表头、可指定行数的表格；时间：2026-10-29；作者：lq。 */
function tableFixture(rowCount: number, headerRows = 1, columnCount = 2): DocTable {
  return {
    type: 'table',
    id: 'table-1',
    columnWidthsPx: Array.from({ length: columnCount }, () => 120),
    rows: Array.from({ length: rowCount }, (_, row) => ({
      id: `row-${row}`,
      isHeader: row < headerRows,
      cells: Array.from({ length: columnCount }, (_, col) => ({
        id: `cell-${row}-${col}`,
        colSpan: 1,
        rowSpan: 1,
        blocks: [makeParagraph([{ text: `r${row}c${col}`, style: {} }])],
      })),
    })),
  };
}

describe('打印纸张', () => {
  it('A4 像素尺寸吸附成毫米纸张，页面盒与 @page 使用同一物理长度', () => {
    const portrait = resolvePrintPaper({ widthPx: 794, heightPx: 1123 });
    expect(portrait.name).toBe('A4');
    expect(portrait.pageRule).toBe('size: 210mm 297mm;');
    expect(portrait.width).toBe('793.701px');
    // 功能：页面盒高度比纸张矮一点，避免按可打印宽度缩放后纵向溢出一张空白纸；时间：2026-10-29；作者：lq。
    expect(portrait.widthMm).toBe(210);
    expect(portrait.heightMm).toBe(297);
    const sheetHeight = Number.parseFloat(portrait.sheetHeight);
    expect(sheetHeight).toBeLessThan(1122.52);
    expect(sheetHeight).toBeGreaterThan(1080);
    const landscape = resolvePrintPaper({ widthPx: 1123, heightPx: 794 });
    expect(landscape.pageRule).toBe('size: 297mm 210mm;');
    expect(landscape.widthMm).toBe(297);
  });

  it('非标准尺寸保留像素纸张，不会被错误吸附', () => {
    const custom = resolvePrintPaper({ widthPx: 500, heightPx: 900 });
    expect(custom.name).toBeUndefined();
    expect(custom.pageRule).toBe('size: 500px 900px;');
    expect(custom.sheetHeight).toBe('877.323px');
  });
});

describe('打印文档', () => {
  it('页数严格等于页面数量，且最后一页不再分页', () => {
    const pages = Array.from({ length: 3 }, () => document.createElement('div'));
    const handle = createPrintRoot({ pages, widthPx: 794, heightPx: 1123 });
    try {
      expect(handle.root.children).toHaveLength(3);
      const style = document.querySelector('style[data-mo-print-style]')!.textContent!;
      expect(style).toContain('@page { size: 210mm 297mm; margin: 0; }');
      expect(style).toContain('.mo-print-page:last-child');
      // 功能：应用界面在打印时必须整体隐藏，否则窗口高度会把打印文档裁成一两页；时间：2026-10-29；作者：lq。
      expect(style).toContain('body > *:not([data-mo-print-root])');
      expect(style).toContain('height: auto !important');
      // 功能：布局宽度必须锁到纸张宽度，否则 WebKit 会按窗口宽度排版再整体缩放，内容会变小且贴左；时间：2026-10-29；作者：lq。
      expect(style).toContain('width: 210mm !important');
    } finally {
      handle.cleanup();
    }
    expect(document.querySelector('style[data-mo-print-style]')).toBeNull();
  });
});

describe('Word 分页', () => {
  it('长表格按行跨页并重复表头，行不丢失不重复', () => {
    const table = tableFixture(40, 1);
    const heights = new Map([[table.id, 40 * 30]]);
    const layouts = new Map([[table.id, { rowHeights: Array.from({ length: 40 }, () => 30), headerRows: 1 }]]);
    const pages = paginateBlocks([table], heights, 300, layouts);
    expect(pages.length).toBeGreaterThan(1);
    const fragments = pages.flatMap((page) => page.blocks).filter((block): block is DocTable => block.type === 'table');
    // 首片自带表头，续片重复表头
    expect(fragments[0].rows[0].isHeader).toBe(true);
    expect(fragments[1].rows[0].isHeader).toBe(true);
    const bodyTexts = fragments.flatMap((fragment) => fragment.rows.filter((row) => !row.isHeader)
      .flatMap((row) => row.cells.flatMap((cell) => cell.blocks.flatMap((block) => (block.type === 'paragraph' ? block.runs.map((run) => run.text ?? '') : [])))));
    expect(bodyTexts).toHaveLength(39 * 2);
    expect(new Set(bodyTexts).size).toBe(39 * 2);
    // 每片都标记为片段，避免被原表映射覆盖；首片沿用原表 id 可编辑，续页锁定
    expect(pages[0].fragments).toEqual([table.id]);
    expect(pages[0].locked).toBeUndefined();
    pages.slice(1).forEach((page) => {
      expect(page.fragments?.length).toBe(1);
      expect(page.locked?.length).toBe(1);
    });
  });

  it('整表放得下时不切片', () => {
    const table = tableFixture(4, 1);
    const heights = new Map([[table.id, 120]]);
    const layouts = new Map([[table.id, { rowHeights: [30, 30, 30, 30], headerRows: 1 }]]);
    const pages = paginateBlocks([table], heights, 300, layouts);
    expect(pages).toHaveLength(1);
    expect(pages[0].fragments).toBeUndefined();
    expect(pages[0].blocks[0]).toBe(table);
  });

  it('跨页合并单元格在续页保留占位与跨度，不产生空洞', () => {
    const table = tableFixture(6, 1);
    table.rows[1].cells[0].rowSpan = 4;
    table.rows[1].cells[1].rowSpan = 4;
    for (let row = 2; row <= 4; row += 1) table.rows[row].cells = [];
    const slice = sliceTable(table, 3, 5, 1);
    expect(slice.rows[0].isHeader).toBe(true);
    const continuation = slice.rows[1];
    expect(continuation.cells[0].rowSpan).toBe(2);
    expect(continuation.cells[0].blocks).toHaveLength(0);
  });

  it('页面分割保持块顺序，分页符强制换页', () => {
    const doc = createBlankDoc();
    const blocks = [
      makeParagraph([{ text: '第一段', style: {} }]),
      { type: 'pageBreak' as const, id: 'pb' },
      makeParagraph([{ text: '第二段', style: {} }]),
    ];
    const heights = new Map(blocks.filter((block) => block.type === 'paragraph').map((block) => [block.id, 20]));
    const pages = paginateBlocks(blocks, heights, 500);
    expect(pages).toHaveLength(2);
    expect(doc.blocks.length).toBeGreaterThan(0);
  });

  it('长段落按行跨页切分，首片沿用原 id 可编辑、续页片段只读', () => {
    const p = makeParagraph([{ text: '一二三四五六七八九十', style: {} }]);
    const heights = new Map([[p.id, 200]]);
    const breaker = capacityBreaker();
    const pages = paginateBlocks([p], heights, 100, undefined, breaker);
    expect(pages.length).toBeGreaterThanOrEqual(2);
    const first = pages[0];
    expect(first.blocks[0].id).toBe(p.id);
    expect(first.fragments).toContain(p.id);
    expect(first.locked).toBeUndefined();
    // 功能：首片的模型边界必须记录下来，编辑回写据此保留续页尾部；时间：2026-10-30；作者：lq。
    expect(first.fragmentEnds?.[p.id]).toBe(5);
    const cont = pages.find((page) => page.locked && page.locked.length > 0);
    expect(cont).toBeDefined();
    expect(cont!.blocks[0].id).toContain('-cont');
    expect(heights.get(cont!.blocks[0].id)).toBe(100);
  });

  it('段内分页符是强制断点，即使整段放得下也要断开', () => {
    const p = makeParagraph([
      { text: '前半', style: {} },
      { text: '', style: {}, pageBreak: true },
      { text: '后半', style: {} },
    ]);
    const heights = new Map([[p.id, 40]]);
    const breaker = capacityBreaker();
    const pages = paginateBlocks([p], heights, 100, undefined, breaker);
    // 功能：分页符后的内容必须落到下一页，不能留在本页；时间：2026-10-30；作者：lq。
    expect(pages.length).toBe(2);
    expect(pages[0].blocks[0].id).toBe(p.id);
    expect(pages[1].locked).toBeDefined();
    expect(runsPlainText(pages[1].blocks[0].type === 'paragraph' ? pages[1].blocks[0].runs : [])).toBe('后半');
  });

  it('跨页首片编辑回写时保留续页尾部', () => {
    const p = makeParagraph([{ text: '前半', style: {} }, { text: '后半', style: {} }]);
    const merged = mergeFragmentRuns(p.runs, 2, [{ text: '新前半', style: {} }], {});
    expect(runsPlainText(merged)).toBe('新前半后半');
  });
});

describe('Excel 分页', () => {
  it('长表格按行分页且每页重复表头', () => {
    const book = createBlankWorkbook();
    const sheet = book.sheets[0];
    sheet.cells['0:0'] = { value: '表头' };
    for (let row = 1; row <= 80; row += 1) sheet.cells[`${row}:0`] = { value: `行${row}` };
    const host = document.createElement('div');
    try {
      const pages = buildSheetPrintPages(book, sheet, host);
      expect(pages.length).toBeGreaterThan(1);
      pages.forEach((page) => expect(page.querySelector('thead td')?.textContent).toBe('表头'));
      const texts = pages.flatMap((page) => Array.from(page.querySelectorAll('tbody td')).map((td) => td.textContent));
      expect(texts).toHaveLength(80);
      expect(new Set(texts).size).toBe(80);
    } finally {
      host.remove();
    }
  });

  it('列宽超出可打印宽度时横向分页，右侧列不再被裁掉', () => {
    const book = createBlankWorkbook();
    const sheet = book.sheets[0];
    for (let col = 0; col < 20; col += 1) {
      sheet.colWidths[col] = 120;
      sheet.cells[`0:${col}`] = { value: `列${col}` };
    }
    const host = document.createElement('div');
    try {
      const pages = buildSheetPrintPages(book, sheet, host);
      expect(pages.length).toBeGreaterThan(1);
      const texts = pages.flatMap((page) => Array.from(page.querySelectorAll('td')).map((td) => td.textContent));
      for (let col = 0; col < 20; col += 1) expect(texts).toContain(`列${col}`);
    } finally {
      host.remove();
    }
  });

  it('跨页合并单元格在续页保留边框底纹且文字只出现一次', () => {
    const book = createBlankWorkbook();
    const sheet = book.sheets[0];
    sheet.cells['0:0'] = { value: '表头' };
    sheet.cells['1:0'] = { value: '合并内容' };
    sheet.merges.push({ r1: 1, c1: 0, r2: 30, c2: 0 });
    for (let row = 2; row <= 60; row += 1) sheet.cells[`${row}:1`] = { value: `行${row}` };
    const host = document.createElement('div');
    try {
      const pages = buildSheetPrintPages(book, sheet, host);
      expect(pages.length).toBeGreaterThan(1);
      const merged = pages.flatMap((page) => Array.from(page.querySelectorAll('tbody td[rowspan]')));
      expect(merged.length).toBeGreaterThan(1);
      expect(merged.filter((td) => td.textContent === '合并内容')).toHaveLength(1);
      expect(merged.every((td) => (td as HTMLElement).style.border !== '')).toBe(true);
    } finally {
      host.remove();
    }
  });


  it('“调整为一页宽”只缩小不放大小表', () => {
    const book = createBlankWorkbook();
    const sheet = book.sheets[0];
    sheet.pageSetup = { ...sheet.pageSetup!, fitToWidth: 1 };
    sheet.cells['0:0'] = { value: '标题' };
    sheet.cells['1:0'] = { value: '内容' };
    const host = document.createElement('div');
    try {
      const page = buildSheetPrintPages(book, sheet, host)[0];
      const table = page.querySelector('table') as HTMLTableElement;
      const expected = (sheet.colWidths[0] ?? sheet.defaultColWidth);
      expect(parseFloat(table.style.width)).toBeCloseTo(expected, 1);
    } finally {
      host.remove();
    }
  });

  it('“调整为页高”把整表压到一页内', () => {
    const book = createBlankWorkbook();
    const sheet = book.sheets[0];
    sheet.pageSetup = { ...sheet.pageSetup!, fitToHeight: 1 };
    sheet.cells['0:0'] = { value: '表头' };
    for (let row = 1; row <= 60; row += 1) sheet.cells[`${row}:0`] = { value: `行${row}` };
    const host = document.createElement('div');
    try {
      const pages = buildSheetPrintPages(book, sheet, host);
      expect(pages).toHaveLength(1);
      expect(pages[0].querySelectorAll('tbody td')).toHaveLength(60);
    } finally {
      host.remove();
    }
  });

  it('纸张尺寸与页面设置方向一致，页面盒带内边距不叠加 @page 边距', () => {    const setup = defaultSheetPageSetup();
    const landscape = sheetPageSize(setup);
    expect(landscape.width).toBeGreaterThan(landscape.height);
    const portrait = sheetPageSize({ ...setup, orientation: 'portrait' });
    expect(portrait.height).toBeGreaterThan(portrait.width);
    const book = createBlankWorkbook();
    const sheet = book.sheets[0];
    sheet.cells['0:0'] = { value: '内容' };
    const host = document.createElement('div');
    try {
      const page = buildSheetPrintPages(book, sheet, host)[0];
      expect(page.style.padding).not.toBe('');
      expect(page.style.width).toBe(`${landscape.width}px`);
      expect(page.style.height).toBe(`${landscape.height}px`);
    } finally {
      host.remove();
    }
  });
});

describe('多节混合方向', () => {
  const portrait = { widthPx: 793.7, heightPx: 1122.52 };
  const landscape = { widthPx: 1122.52, heightPx: 793.7 };

  it('按页分配纸张，横向页与纵向页各自一套 @page 规则', () => {
    const pages = Array.from({ length: 3 }, () => document.createElement('div'));
    const handle = createPrintRoot({
      pages,
      pageSizes: [portrait, landscape, portrait],
      widthPx: portrait.widthPx,
      heightPx: portrait.heightPx,
    });
    try {
      const style = document.querySelector('style[data-mo-print-style]')!.textContent!;
      expect(style).toContain('@page mo-sheet-1 { size: 210mm 297mm; margin: 0; }');
      expect(style).toContain('@page mo-sheet-2 { size: 297mm 210mm; margin: 0; }');
      expect(style).toContain('.mo-print-page.mo-sheet-2 { page: mo-sheet-2; }');
      // 功能：横向页必须拿到横向页面盒，不能被统一纸张拉成纵向；时间：2026-10-29；作者：lq。
      expect(pages[0].classList.contains('mo-sheet-1')).toBe(true);
      expect(pages[1].classList.contains('mo-sheet-2')).toBe(true);
      expect(pages[1].style.width).toBe('1122.52px');
      expect(pages[1].style.height).toBe('771.024px');
      expect(pages[2].style.width).toBe('793.701px');
      // 功能：打印根宽度取最大页宽，横向页不会被裁；时间：2026-10-29；作者：lq。
      expect(style).toContain('width: 297mm !important');
    } finally {
      handle.cleanup();
    }
  });

  it('单一纸张时不生成命名页组，保持原来的单条 @page', () => {
    const pages = Array.from({ length: 2 }, () => document.createElement('div'));
    const handle = createPrintRoot({ pages, widthPx: portrait.widthPx, heightPx: portrait.heightPx });
    try {
      const style = document.querySelector('style[data-mo-print-style]')!.textContent!;
      expect(style).toContain('@page { size: 210mm 297mm; margin: 0; }');
      expect(style).not.toContain('mo-sheet-');
    } finally {
      handle.cleanup();
    }
  });

  it('预览按每页纸张渲染画框，并在页码后标出该页方向', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const pages = Array.from({ length: 2 }, () => document.createElement('div'));
    const cleanup = renderPrintPreview({
      host,
      pages,
      pageSizes: [portrait, landscape],
      widthPx: portrait.widthPx,
      heightPx: portrait.heightPx,
    });
    try {
      const columns = Array.from(host.querySelectorAll<HTMLElement>(':scope > div > div'));
      expect(columns).toHaveLength(2);
      const frames = columns.map((column) => column.lastElementChild as HTMLElement);
      const [first, second] = frames.map((frame) => Number.parseFloat(frame.style.width));
      expect(second).toBeGreaterThan(first);
      const labels = columns.map((column) => column.firstElementChild?.textContent ?? '');
      expect(labels[0]).toContain('纵向');
      expect(labels[1]).toContain('横向');
    } finally {
      cleanup();
      host.remove();
    }
  });
});
