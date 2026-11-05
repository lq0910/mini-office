// 功能：锁定 OOXML 子元素顺序与图形 ID 唯一性，防止 Word/Excel 因模式校验失败拒绝打开；时间：2026-10-29；作者：lq。
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createBlankDoc, createBlankWorkbook, type DocModel, type DocTable } from '@mini-office/document-model';
import { parseXmlDom } from '@mini-office/ooxml';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import { serializeXlsx } from '../packages/editor-sheet/src/serialize';

/** 功能：按 ECMA-376 声明的序列校验子元素顺序，返回违规描述；时间：2026-10-29；作者：lq。 */
function orderViolations(xml: string, tag: string, order: string[]): string[] {
  const doc = parseXmlDom(xml);
  const problems: string[] = [];
  for (const container of Array.from(doc.getElementsByTagName(`w:${tag}`))) {
    let last = -1;
    for (const child of Array.from(container.children)) {
      const index = order.indexOf(child.localName ?? '');
      if (index < 0) problems.push(`${tag} 出现未知子元素 ${child.localName}`);
      else if (index < last) problems.push(`${tag} 子元素顺序错误：${child.localName}`);
      else last = index;
    }
  }
  return problems;
}

const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
const TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange'];
const TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription', 'tblPrChange'];
const SECTPR_ORDER = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'];

/** 功能：构造覆盖全部属性分支的文档，确保每种属性组合都参与顺序校验；时间：2026-10-29；作者：lq。 */
function richDoc(): DocModel {
  const doc = createBlankDoc();
  doc.header = { blocks: [{ type: 'paragraph', id: 'h1', style: { alignment: 'center', borders: { bottom: { style: 'single', widthPx: 1, color: '#000' } } }, runs: [{ text: '页眉', style: { fontSizePt: 9 } }] }], edited: true, distancePx: 38 };
  doc.blocks = [
    {
      type: 'paragraph',
      id: 'p1',
      // 功能：同时覆盖 snapToGrid、边框和底纹，校验新增属性组合的真实模式顺序；时间：2026-11-03；作者：lq。
      style: { alignment: 'justify', spaceBeforePx: 6, spaceAfterPx: 12, lineHeight: 1.5, firstLineIndentPx: 28, backgroundColor: '#fff', keepLines: true, widowControl: false, snapToGrid: false, borders: { bottom: { style: 'single', widthPx: 1, color: '#333' } } },
      runs: [{ text: '标题', style: { bold: true, italic: true, underline: true, strike: true, allCaps: true, smallCaps: true, color: '#f00', highlight: '#ff0', fontSizePt: 12, letterSpacingPx: 1, verticalAlign: 'super' } }],
    },
    {
      type: 'paragraph',
      id: 'p2',
      style: { alignment: 'left', lineHeightExactPx: 24 },
      runs: [
        { text: '第一行', style: { fontSizePt: 10.5 } },
        { text: '', style: {}, lineBreak: true },
        { text: '第二行', style: { fontSizePt: 10.5 } },
        { text: '', style: {}, image: { id: 'img1', src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', dataUrl: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', widthPx: 40, heightPx: 40, wrap: 'square-left' } },
        { text: '', style: {}, image: { id: 'img2', src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', dataUrl: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', widthPx: 20, heightPx: 20, wrap: 'inline' } },
      ],
    },
    {
      type: 'table',
      id: 't1',
      columnWidthsPx: [120, 120],
      widthPct: 100,
      alignment: 'center',
      // 功能：确保表格缩进与内边距的新增写入分支参与顺序验证；时间：2026-11-03；作者：lq。
      indentPx: 12,
      cellMarginPx: { top: 0, right: 7, bottom: 0, left: 7 },
      borders: { top: { style: 'single', widthPx: 1, color: '#000' }, insideV: { style: 'single', widthPx: 1, color: '#000' } },
      rows: [
        { id: 'r1', heightPx: 24, isHeader: true, cells: [
          { id: 'c1', colSpan: 1, rowSpan: 2, verticalAlign: 'middle', marginPx: { top: 2, right: 5, bottom: 2, left: 5 }, backgroundColor: '#eee', borders: { top: { style: 'single', widthPx: 1, color: '#000' } }, blocks: [{ type: 'paragraph', id: 'p3', style: {}, runs: [{ text: '合并', style: {} }] }] },
          { id: 'c2', colSpan: 1, rowSpan: 1, verticalAlign: 'bottom', blocks: [{ type: 'paragraph', id: 'p4', style: {}, runs: [{ text: '右一', style: {} }] }] },
        ] },
        { id: 'r2', cells: [
          { id: 'c3', colSpan: 1, rowSpan: 1, verticalAlign: 'top', blocks: [{ type: 'paragraph', id: 'p5', style: {}, runs: [{ text: '右二', style: {} }] }] },
        ] },
      ],
    } as DocTable,
    { type: 'pageBreak', id: 'b1' },
  ];
  doc.page.pageNumberStart = 3;
  return doc;
}

describe('DOCX 部件符合 ECMA-376 顺序', () => {
  it('段落、字符、单元格、表格与节属性都按模式顺序输出', async () => {
    const bytes = await serializeDocx(richDoc(), { name: 'order.docx' });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')!.async('text');
    expect(orderViolations(xml, 'pPr', PPR_ORDER)).toEqual([]);
    expect(orderViolations(xml, 'rPr', RPR_ORDER)).toEqual([]);
    expect(orderViolations(xml, 'tcPr', TCPR_ORDER)).toEqual([]);
    expect(orderViolations(xml, 'tblPr', TBLPR_ORDER)).toEqual([]);
    expect(orderViolations(xml, 'sectPr', SECTPR_ORDER)).toEqual([]);
    // 功能：确认分页属性确实写出，且 widowControl 仅在显式关闭时输出 w:val="0"；时间：2026-11-03；作者：lq。
    expect(xml).toContain('<w:keepLines/>');
    expect(xml).toContain('<w:widowControl w:val="0"/>');
    for (const name of ['word/header1.xml', 'word/footer1.xml']) {
      const part = zip.file(name);
      if (!part) continue;
      const partXml = await part.async('text');
      expect(orderViolations(partXml, 'pPr', PPR_ORDER)).toEqual([]);
      expect(orderViolations(partXml, 'rPr', RPR_ORDER)).toEqual([]);
    }
  });

  it('浮动图的环绕方式排在 wp:docPr 之前，且图形 ID 全文档唯一', async () => {
    const bytes = await serializeDocx(richDoc(), { name: 'order.docx' });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')!.async('text');
    const doc = parseXmlDom(xml);
    const anchor = doc.getElementsByTagName('wp:anchor')[0];
    const children = Array.from(anchor.children).map((child) => child.localName);
    expect(children.indexOf('wrapSquare')).toBeLessThan(children.indexOf('docPr'));
    expect(children.indexOf('docPr')).toBeLessThan(children.indexOf('graphic'));
    const ids = Array.from(doc.getElementsByTagName('wp:docPr')).map((element) => element.getAttribute('id'));
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('颜色值统一为 6 位十六进制，避免 ST_HexColor 校验失败', async () => {
    const bytes = await serializeDocx(richDoc(), { name: 'order.docx' });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')!.async('text');
    const doc = parseXmlDom(xml);
    for (const color of Array.from(doc.getElementsByTagName('w:color'))) {
      expect(color.getAttribute('w:val')).toMatch(/^[0-9A-F]{6}$/);
    }
    for (const shd of Array.from(doc.getElementsByTagName('w:shd'))) {
      expect(shd.getAttribute('w:fill')).toMatch(/^(auto|[0-9A-F]{6})$/);
    }
  });
});

describe('XLSX 工作表符合 CT_Worksheet 顺序', () => {
  const WORKSHEET_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

  it('筛选、合并、打印设置按顺序输出', async () => {
    const workbook = createBlankWorkbook();
    const sheet = workbook.sheets[0];
    sheet.cells['0:0'] = { value: '编号' };
    sheet.autoFilter = { r1: 0, c1: 0, r2: 1, c2: 1 };
    sheet.merges = [{ r1: 0, c1: 0, r2: 0, c2: 1 }];
    sheet.pageSetup = { orientation: 'portrait', scalePercent: 100, margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }, printGridLines: true, printHeadings: false };
    const bytes = await serializeXlsx(workbook, { name: 'order.xlsx' });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('text');
    const doc = parseXmlDom(xml);
    const worksheet = doc.getElementsByTagName('worksheet')[0];
    let last = -1;
    for (const child of Array.from(worksheet.children)) {
      const index = WORKSHEET_ORDER.indexOf(child.localName ?? '');
      expect(index).toBeGreaterThanOrEqual(0);
      expect(index).toBeGreaterThan(last);
      last = index;
    }
  });
});
