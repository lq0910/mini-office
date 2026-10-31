/**
 * 功能：验证 Word 版式兼容——连续分节不翻页、不等宽多栏（w:cols + 栏宽 + 分栏符）、
 * VML 复选框绝对定位浮层、文档网格行距、默认制表位；时间：2026-10-30；作者：hou0。
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createBlankDoc, type DocBlock, type DocModel } from '@mini-office/document-model';
import {
  distributeColumns,
  documentSections,
  paginateDocument,
  parseDocx,
  paragraphTextStyle,
  renderParagraphHtml,
  serializeDocx,
} from '@mini-office/editor-word';

/** 功能：构造最小但合法的 docx，方便验证 sectPr/浮动形状的解析与往返；时间：2026-10-30；作者：hou0。 */
async function makeDocx(body: string, sectPr: string, settings?: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  );
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>${body}${sectPr}</w:body></w:document>`,
  );
  if (settings) zip.file('word/settings.xml', settings);
  return zip.generateAsync({ type: 'uint8array' });
}

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

describe('Word 版式兼容', () => {
  it('连续分节不翻页，同一页按分节记录多栏分组', () => {
    const base = createBlankDoc().page;
    const doc: DocModel = {
      ...createBlankDoc(),
      blocks: [
        // 与真实文档同构：标题单独一节（连续），三栏节含三个词并以连续节收尾，之后接正文。
        { type: 'paragraph', id: 'a', style: {}, runs: [{ text: '标题', style: {} }], sectionEnd: { page: { ...base, type: 'continuous' } } },
        { type: 'paragraph', id: 'b', style: {}, runs: [{ text: '第一栏', style: {} }] },
        { type: 'paragraph', id: 'c', style: { columnBreakBefore: true }, runs: [{ text: '第二栏', style: {} }] },
        { type: 'paragraph', id: 'd', style: { columnBreakBefore: true }, runs: [{ text: '第三栏', style: {} }] },
        {
          type: 'paragraph',
          id: 'e',
          style: {},
          runs: [{ text: '', style: {} }],
          sectionEnd: { page: { ...base, type: 'continuous', columns: 3, columnWidthsPx: [200, 100, 220], columnSpacePx: [10, 10] } },
        },
        { type: 'paragraph', id: 'f', style: {}, runs: [{ text: '结尾', style: {} }], sectionEnd: { page: { ...base, type: 'continuous' } } },
      ],
    };
    const heights = new Map(doc.blocks.map((block) => [block.id, 20]));
    const pages = paginateDocument(doc, heights);
    // 连续分节不另起页：全部内容落在第 1 页。
    expect(pages).toHaveLength(1);
    expect(pages[0].blocks.map((block) => block.id)).toEqual(['a', 'b', 'c', 'd', 'e', 'f']);
    expect(pages[0].columnGroups).toHaveLength(1);
    expect(pages[0].columnGroups?.[0]).toMatchObject({ count: 3, widthsPx: [200, 100, 220] });
    expect(pages[0].columnGroups?.[0].blockIds).toEqual(['b', 'c', 'd', 'e']);
    expect(pages[0].usedHeight).toBeGreaterThan(0);
  });

  it('多栏节末尾的分节标记空段落不占行', () => {
    const base = createBlankDoc().page;
    const blocks: DocBlock[] = [
      { type: 'paragraph', id: 'w1', style: {}, runs: [{ text: 'Retrospective', style: {} }] },
      { type: 'paragraph', id: 'w2', style: { columnBreakBefore: true }, runs: [{ text: 'Prospective', style: {} }] },
      { type: 'paragraph', id: 'w3', style: { columnBreakBefore: true }, runs: [{ text: 'Re-qualification', style: {} }] },
      { type: 'paragraph', id: 'marker', style: {}, runs: [{ text: '', style: {} }], sectionEnd: { page: { ...base, type: 'continuous', columns: 3 } } },
    ];
    const columns = distributeColumns(blocks, 3);
    // 分节标记空段落不占行（Word 实测：勾选行→下一段=一行正文+一行空段，不含标记行）。
    expect(columns.map((column) => column.map((block) => block.id))).toEqual([['w1'], ['w2'], ['w3']]);
  });

  it('nextPage 节仍强制翻页', () => {
    const base = createBlankDoc();
    const doc: DocModel = {
      ...base,
      blocks: [
        { type: 'paragraph', id: 'a', style: {}, runs: [{ text: '第一节', style: {} }], sectionEnd: { page: { ...base.page, type: 'nextPage' } } },
        { type: 'paragraph', id: 'b', style: {}, runs: [{ text: '第二节', style: {} }] },
      ],
    };
    const heights = new Map([['a', 20], ['b', 20]]);
    const pages = paginateDocument(doc, heights);
    expect(pages).toHaveLength(2);
    expect(pages[1].blocks.map((block) => block.id)).toEqual(['b']);
  });

  it('分栏符决定栏内块分布，缺分栏符时按块数均衡', () => {
    const blocks = [
      { type: 'paragraph' as const, id: 'p1', style: {}, runs: [{ text: 'Retrospective', style: {} }] },
      { type: 'paragraph' as const, id: 'p2', style: { columnBreakBefore: true }, runs: [{ text: 'Prospective', style: {} }] },
      { type: 'paragraph' as const, id: 'p3', style: { columnBreakBefore: true }, runs: [{ text: 'Re-qualification', style: {} }] },
      { type: 'paragraph' as const, id: 'p4', style: {}, runs: [{ text: '', style: {} }] },
    ];
    const columns = distributeColumns(blocks, 3);
    expect(columns.map((column) => column.map((block) => block.id))).toEqual([['p1'], ['p2'], ['p3', 'p4']]);
    // 无分栏符：9 块 3 栏 → 3/3/3，而不是轮询。
    const even = Array.from({ length: 9 }, (_, index) => ({ type: 'paragraph' as const, id: `q${index}`, style: {}, runs: [{ text: String(index), style: {} }] }));
    expect(distributeColumns(even, 3).map((column) => column.map((block) => block.id))).toEqual([
      ['q0', 'q1', 'q2'],
      ['q3', 'q4', 'q5'],
      ['q6', 'q7', 'q8'],
    ]);
  });

  it('解析不等宽多栏、连续节、网格行距与 VML 复选框浮层', async () => {
    const body =
      '<w:p><w:r><w:t>标题</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:ind w:left="2545" w:firstLine="0"/></w:pPr>' +
      '<w:r><w:pict><v:rect style="position:absolute;margin-left:151.34pt;margin-top:5.93pt;width:9.84pt;height:9.84pt;mso-position-horizontal-relative:page" stroked="true" strokecolor="#000000" strokeweight=".72pt"/></w:pict></w:r>' +
      '<w:r><w:t>Retrospective</w:t></w:r></w:p>' +
      '<w:p><w:r><w:br w:type="column"/></w:r><w:r><w:t>Prospective</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11910" w:h="16840"/><w:pgMar w:top="1580" w:bottom="280" w:left="880" w:right="680"/>' +
      '<w:cols w:num="3" w:equalWidth="0"><w:col w:w="4060" w:space="40"/><w:col w:w="1827" w:space="39"/><w:col w:w="4384"/></w:cols>' +
      '<w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr></w:pPr></w:p>';
    const bytes = await makeDocx(
      body,
      `<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="11910" w:h="16840"/><w:pgMar w:top="1580" w:bottom="280" w:left="880" w:right="680"/></w:sectPr>`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${NS}><w:defaultTabStop w:val="420"/></w:settings>`,
    );
    const doc = await parseDocx(bytes);
    expect(doc.defaultTabStopPx).toBeCloseTo(28, 0);

    const sections = documentSections(doc);
    const columnSection = sections.find((section) => section.page.columns === 3);
    expect(columnSection).toBeTruthy();
    expect(columnSection!.page.type).toBe('continuous');
    expect(columnSection!.page.columnWidthsPx?.map((width) => Math.round(width))).toEqual([271, 122, 292]);
    expect(columnSection!.page.columnSpacePx?.map((space) => Math.round(space))).toEqual([3, 3]);
    expect(Math.round(columnSection!.page.linePitchPx ?? 0)).toBe(21);

    const paragraph = doc.blocks.find((block) => block.type === 'paragraph' && block.runs.some((run) => run.image?.vmlRect));
    expect(paragraph).toBeTruthy();
    const shape = paragraph!.type === 'paragraph' ? paragraph!.runs.find((run) => run.image?.vmlRect)?.image : undefined;
    expect(shape?.vmlRect).toMatchObject({ xRelative: 'page' });
    expect(Math.round(shape!.vmlRect!.left)).toBe(202);
    expect(Math.round(shape!.widthPx)).toBe(13);
    expect(shape?.strokeColor).toBe('#000000');

    // 浮动形状不占行内空间：渲染为绝对定位方框而非 □ 字符。
    const html = renderParagraphHtml(paragraph as never, { defaultTabStopPx: doc.defaultTabStopPx });
    expect(html).toContain('data-vml-rect="1"');
    expect(html).toContain('position:absolute');
    expect(html).toContain('data-vml-xrel="page"');
    expect(html).not.toContain('□');

    // 段首分栏符提升为段落属性。
    const prospective = doc.blocks.find((block) => block.type === 'paragraph' && (block.runs[0]?.text ?? '').includes('Prospective'));
    expect(prospective?.type === 'paragraph' && prospective.style.columnBreakBefore).toBe(true);
  });

  it('保存后栏宽、节类型、网格行距与复选框浮层仍可往返', async () => {
    const doc = await parseDocx(
      await makeDocx(
        '<w:p><w:pPr><w:ind w:left="2545"/></w:pPr><w:r><w:pict><v:rect style="position:absolute;margin-left:151.34pt;margin-top:-25.62pt;width:9.84pt;height:9.84pt;mso-position-horizontal-relative:page" stroked="true" strokecolor="#000000"/></w:pict></w:r><w:r><w:t>项</w:t></w:r></w:p>' +
          '<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11910" w:h="16840"/><w:pgMar w:top="1580" w:bottom="280" w:left="880" w:right="680"/><w:cols w:num="3" w:equalWidth="0"><w:col w:w="4060" w:space="40"/><w:col w:w="1827" w:space="39"/><w:col w:w="4384"/></w:cols><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr></w:pPr></w:p>',
        `<w:sectPr><w:pgSz w:w="11910" w:h="16840"/><w:pgMar w:top="1580" w:bottom="280" w:left="880" w:right="680"/></w:sectPr>`,
      ),
    );
    const saved = await serializeDocx(doc, { name: 'layout.docx' });
    const reopened = await parseDocx(saved);
    const section = documentSections(reopened).find((item) => item.page.columns === 3);
    expect(section).toBeTruthy();
    expect(section!.page.type).toBe('continuous');
    expect(section!.page.columnWidthsPx?.map((width) => Math.round(width))).toEqual([271, 122, 292]);
    expect(Math.round(section!.page.linePitchPx ?? 0)).toBe(21);
    const shape = reopened.blocks
      .flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
      .find((run) => run.image?.vmlRect)?.image?.vmlRect;
    expect(shape).toMatchObject({ xRelative: 'page' });
    // 负 margin-top 是第三列复选框对齐到同一行用的，往返不能丢。
    expect(Math.round(shape!.top)).toBe(-34);
  });

  it('文档网格：行高向上取整到网格整数倍，关闭对齐的段落不取整', () => {
    const grid = 21;
    const normal = paragraphTextStyle({ type: 'paragraph', id: 'n', runs: [{ text: 'a', style: {} }], style: { lineHeight: 1.15 } }, '宋体', 10.5, grid);
    const raw = paragraphTextStyle({ type: 'paragraph', id: 'n', runs: [{ text: 'a', style: {} }], style: { lineHeight: 1.15 } }, '宋体', 10.5);
    expect(parseFloat(String(normal.lineHeight))).toBeGreaterThan(parseFloat(String(raw.lineHeight)));
    expect(parseFloat(String(normal.lineHeight)) % grid).toBeCloseTo(0, 5);
    const optedOut = paragraphTextStyle({ type: 'paragraph', id: 'o', runs: [{ text: 'a', style: {} }], style: { lineHeight: 1.15, snapToGrid: false } }, '宋体', 10.5, grid);
    expect(optedOut.lineHeight).toBe(raw.lineHeight);
  });

  it('普通段落按默认制表位跳转，右制表位内容过宽不折行', () => {
    const html = renderParagraphHtml({
      type: 'paragraph',
      id: 't',
      style: {},
      runs: [
        { text: '', style: {}, tab: true },
        { text: '深圳市XXX有限公司', style: {} },
      ],
    }, { defaultTabStopPx: 48 });
    // 行首制表符跳到第一个默认制表位（48px），不再是固定两个全角空格。
    expect(html).toContain('display:inline-block;width:48px');
    expect(html).not.toContain('\u2003\u2003');

    const rightTab = renderParagraphHtml({
      type: 'paragraph',
      id: 'r',
      style: { tabStops: [{ posPx: 20, kind: 'right' }] },
      runs: [
        { text: '', style: {}, tab: true },
        { text: '深圳市XXX有限公司', style: {} },
      ],
    });
    expect(rightTab).toContain('text-align:right');
    expect(rightTab).toContain('white-space:nowrap');
    // 内容宽于空隙时向左溢出，右缘仍落在制表位上。
    expect(rightTab).toMatch(/margin-left:-\d/);
  });
});
