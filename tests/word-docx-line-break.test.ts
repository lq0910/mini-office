import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { loadOoxmlPackage, parseXmlDom } from '@mini-office/ooxml';
import { parseDocxPackage } from '../packages/editor-word/src/parse';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import { paragraphTextStyle, renderParagraphHtml, renderRunsHtml } from '../packages/editor-word/src/html';
import type { DocBlock, DocImage, DocParagraph } from '@mini-office/document-model';
// 功能：对比图片倍数行距保留的额外间距与文字单行度量；时间：2026-11-05；作者：lq。
import { singleLineHeightPx } from '../packages/editor-word/src/metrics';
// 功能：检查导入后相邻段落实际使用的上下文间距；时间：2026-11-05；作者：lq。
import { effectiveParagraphSpacing } from '../packages/editor-word/src/paragraph-spacing';

/**
 * 功能：DOCX 换行与浮动图片回归测试；时间：2026-10-31；作者：lq。
 *
 * 背景：WPS/Word 常把签名下划线做成「锚定在空段落里的浮动图片」。
 * 旧渲染给浮动图硬编码 8px 下边距，使 8px 高的图片实际占 16px 行盒，
 * 正好压住下一段首行，把「乙方：xxx」整行文字挤成两行，表现为「换行不对」。
 */

function installBlobShim(): void {
  if (typeof URL.createObjectURL !== 'function') {
    (URL as unknown as { createObjectURL: (blob: Blob) => string }).createObjectURL = () => 'blob:test';
    (URL as unknown as { revokeObjectURL: (url: string) => void }).revokeObjectURL = () => {};
  }
}

function paragraphsOf(blocks: DocBlock[]): DocParagraph[] {
  return blocks.filter((block): block is DocParagraph => block.type === 'paragraph');
}

function paragraphText(paragraph: DocParagraph): string {
  return paragraph.runs.map((run) => run.text ?? '').join('');
}

async function parseFixture(path: string) {
  installBlobShim();
  const bytes = new Uint8Array(readFileSync(path));
  const pkg = await loadOoxmlPackage(bytes);
  const doc = await parseDocxPackage(pkg, { fileName: path });
  return { doc, bytes, pkg };
}

function wrapXml(body: string): string {
  return `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;
}

async function docxWithBody(body: string, parts: Record<string, string> = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.file('word/document.xml', wrapXml(body));
  for (const [path, xml] of Object.entries(parts)) zip.file(path, xml);
  return zip.generateAsync({ type: 'uint8array' });
}

async function parseBody(body: string, parts: Record<string, string> = {}) {
  installBlobShim();
  const bytes = await docxWithBody(body, parts);
  const pkg = await loadOoxmlPackage(bytes);
  const doc = await parseDocxPackage(pkg, { fileName: 'inline.docx' });
  return { doc, pkg, bytes };
}

describe('DOCX 段内换行语义', () => {
  it('深色文字带 data-ink 标记，浅色文字不带，供深色纸张提亮与打印还原', () => {
    const dark = renderRunsHtml([{ text: '正文', style: { color: '#000000' } }]);
    const light = renderRunsHtml([{ text: '强调', style: { color: '#C00000' } }]);
    expect(dark).toContain('data-ink="1"');
    expect(dark).toContain('color:#000000');
    expect(light).not.toContain('data-ink');
    expect(light).toContain('color:#C00000');
  });

  it('显示换行标记不占行宽，不能把页底正文多挤出一行', () => {
    const html = renderRunsHtml([{ text: '', style: {}, lineBreak: true }], { showFormattingMarks: true });
    expect(html).toContain('width:0;height:0;');
    expect(html).toContain('</span><br>');
  });

  it('w:br 与 w:cr 生成段内换行而不是段落', async () => {
    const { doc, pkg } = await parseBody(
      '<w:p><w:r><w:t>第一行</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>第二行</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>甲</w:t><w:cr/><w:t>乙</w:t></w:r></w:p>',
    );
    try {
      const paragraphs = paragraphsOf(doc.blocks);
      expect(paragraphs).toHaveLength(2);
      expect(paragraphs[0].runs.filter((run) => run.lineBreak)).toHaveLength(1);
      expect(paragraphText(paragraphs[0])).toBe('第一行第二行');
      expect(paragraphs[1].runs.filter((run) => run.lineBreak)).toHaveLength(1);
      // 功能：段内换行渲染成 <br>，不能退化成段落分隔；时间：2026-10-31；作者：lq。
      expect(renderRunsHtml(paragraphs[0].runs)).toContain('<br>');
    } finally {
      pkg.dispose();
    }
  });

  it('w:firstLineChars 按字符换算首行缩进，且优先于 twips', async () => {
    const { doc, pkg } = await parseBody(
      '<w:p><w:pPr><w:ind w:firstLineChars="200"/></w:pPr><w:r><w:t>缩进两字</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:ind w:firstLineChars="200" w:firstLine="0"/></w:pPr><w:r><w:t>字符优先</w:t></w:r></w:p>',
    );
    try {
      const [first, second] = paragraphsOf(doc.blocks);
      // 默认 10.5pt 的汉字宽约 14px，2 字符约 28px。
      expect(first.style.firstLineIndentPx).toBeCloseTo(28, 1);
      expect(second.style.firstLineIndentPx).toBeCloseTo(28, 1);
    } finally {
      pkg.dispose();
    }
  });

  // 功能：保留无 linePitch 的字符网格和显式零 charSpace，避免解析或另存时丢弃设置；时间：2026-11-05；作者：lq。
  it('读取并往返保存仅含字符间距的 snapToChars 网格', async () => {
    const { doc, pkg } = await parseBody(
      '<w:p><w:r><w:t>网格文本</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:docGrid w:type="snapToChars" w:charSpace="0"/></w:sectPr>',
    );
    try {
      expect(doc.page.docGrid).toEqual({ type: 'snapToChars', charSpacePx: 0 });
      const reopened = await loadOoxmlPackage(await serializeDocx(doc, { name: 'character-grid.docx' }));
      try {
        const xml = await reopened.text('word/document.xml');
        expect(xml).toContain('<w:docGrid w:type="snapToChars" w:charSpace="0"');
        const roundTrip = await parseDocxPackage(reopened);
        expect(roundTrip.page.docGrid).toEqual({ type: 'snapToChars', charSpacePx: 0 });
      } finally {
        reopened.dispose();
      }
    } finally {
      pkg.dispose();
    }
  });

  // 功能：字符缩进使用段落标记明确声明的字号，不能退回到文档默认字号；时间：2026-11-05；作者：lq。
  it('大字号段落的字符缩进按该段字号计算', async () => {
    const { doc, pkg } = await parseBody('<w:p><w:pPr><w:ind w:firstLineChars="300"/><w:rPr><w:sz w:val="32"/></w:rPr></w:pPr><w:r><w:t>三字缩进</w:t></w:r></w:p>');
    try {
      expect(paragraphsOf(doc.blocks)[0].style.firstLineIndentPx).toBeCloseTo(64, 1);
    } finally { pkg.dispose(); }
  });

  // 功能：段落样式继承 contextualSpacing，直接 false 可关闭，并在重建保存后保留；时间：2026-11-05；作者：lq。
  it('继承和保存 contextualSpacing，显式 false 覆盖样式', async () => {
    const { doc, pkg, bytes } = await parseBody(
      '<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>同样式一</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>同样式二</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Body"/><w:contextualSpacing w:val="0"/></w:pPr><w:r><w:t>显式关闭</w:t></w:r></w:p>',
      { 'word/styles.xml': '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Body"><w:pPr><w:spacing w:before="120" w:after="80"/><w:contextualSpacing/></w:pPr></w:style></w:styles>' },
    );
    try {
      const [first, second, disabled] = paragraphsOf(doc.blocks);
      expect(first.style).toMatchObject({ spaceBeforePx: 8, contextualSpacing: true });
      expect(first.style.spaceAfterPx).toBeCloseTo(5.3333333333, 8);
      expect(second.style.contextualSpacing).toBe(true);
      expect(disabled.style.contextualSpacing).toBe(false);
      expect(effectiveParagraphSpacing(doc.blocks, 0)).toEqual({ beforePx: 8, afterPx: 0 });
      expect(effectiveParagraphSpacing(doc.blocks, 1)).toMatchObject({ beforePx: 0 });
      expect(effectiveParagraphSpacing(doc.blocks, 1)?.afterPx).toBe(0);
      expect(effectiveParagraphSpacing(doc.blocks, 2)).toMatchObject({ beforePx: 8 });
      expect(effectiveParagraphSpacing(doc.blocks, 2)?.afterPx).toBeCloseTo(5.3333333333, 8);
      expect(await serializeDocx(doc, { sourceBytes: bytes, name: 'context-spacing.docx' })).toEqual(bytes);

      first.runs[0].style.color = '#123456';
      disabled.runs[0].style.color = '#654321';
      const saved = await serializeDocx(doc, { sourceBytes: bytes, name: 'context-spacing.docx' });
      const reopened = await loadOoxmlPackage(saved);
      try {
        const xml = await reopened.text('word/document.xml');
        expect(xml).toContain('<w:contextualSpacing w:val="0"/>');
        const roundTrip = paragraphsOf((await parseDocxPackage(reopened)).blocks);
        expect(roundTrip[0].style.contextualSpacing).toBe(true);
        expect(roundTrip[2].style.contextualSpacing).toBe(false);
      } finally { reopened.dispose(); }
    } finally { pkg.dispose(); }
  });

  // 功能：段落结束标记的加粗、字号与正文独立，空段、样式继承及编辑重建均保留各自格式；时间：2026-11-05；作者：lq。
  it('段落标记不污染正文，格式修改保存后仍独立', async () => {
    const mark = '<w:pPr><w:rPr><w:b/><w:i/><w:sz w:val="40"/></w:rPr></w:pPr>';
    const { doc, pkg, bytes } = await parseBody(
      `<w:p>${mark}<w:r><w:t>普通正文</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>加粗正文</w:t></w:r></w:p>` +
      `<w:p>${mark}</w:p>` +
      '<w:p><w:pPr><w:pStyle w:val="Bold"/><w:rPr><w:b w:val="0"/></w:rPr></w:pPr><w:r><w:t>样式加粗</w:t></w:r></w:p>',
      { 'word/styles.xml': '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Bold"><w:rPr><w:b/></w:rPr></w:style></w:styles>' },
    );
    try {
      const [body, empty, styled] = paragraphsOf(doc.blocks);
      expect(body.markStyle).toMatchObject({ bold: true, italic: true, fontSizePt: 20 });
      expect(body.runs[0].style.bold).not.toBe(true);
      expect(body.runs[0].style.italic).not.toBe(true);
      expect(body.runs[0].style.fontSizePt).not.toBe(20);
      expect(body.runs[1].style.bold).toBe(true);
      expect(empty.runs[0].style).toMatchObject({ bold: true, italic: true, fontSizePt: 20 });
      expect(styled.markStyle?.bold).toBe(false);
      expect(styled.runs[0].style.bold).toBe(true);
      expect(await serializeDocx(doc, { sourceBytes: bytes, name: 'marks.docx' })).toEqual(bytes);
      // 功能：修改正文格式强制走段落重建，防止仅靠原 XML 透传掩盖标记丢失；时间：2026-11-05；作者：lq。
      body.runs[0].style.color = '#123456';
      empty.runs[0].style.fontSizePt = 24;
      const saved = await serializeDocx(doc, { sourceBytes: bytes, name: 'marks.docx' });
      const reopenedPkg = await loadOoxmlPackage(saved);
      try {
        const reopenedParagraphs = paragraphsOf((await parseDocxPackage(reopenedPkg)).blocks);
        const reopened = reopenedParagraphs[0];
        expect(reopened.markStyle).toMatchObject({ bold: true, italic: true, fontSizePt: 20 });
        expect(reopened.runs[0].style).toMatchObject({ color: '#123456' });
        expect(reopened.runs[0].style.bold).not.toBe(true);
        expect(reopened.runs[0].style.fontSizePt).not.toBe(20);
        expect(reopenedParagraphs[1].markStyle?.fontSizePt).toBe(24);
        expect(reopenedParagraphs[1].runs[0].style.fontSizePt).toBe(24);
      } finally { reopenedPkg.dispose(); }
    } finally { pkg.dispose(); }
  });

  it('w:keepLines 与 w:widowControl 解析为分页属性', async () => {
    const { doc, pkg } = await parseBody(
      '<w:p><w:pPr><w:keepLines/><w:widowControl w:val="0"/></w:pPr><w:r><w:t>条款</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:widowControl/></w:pPr><w:r><w:t>正文</w:t></w:r></w:p>',
    );
    try {
      const [first, second] = paragraphsOf(doc.blocks);
      expect(first.style.keepLines).toBe(true);
      expect(first.style.widowControl).toBe(false);
      // 功能：w:widowControl 不带 w:val 等价于开启，缺省同样视为开启；时间：2026-11-03；作者：lq。
      expect(second.style.widowControl).toBe(true);
    } finally {
      pkg.dispose();
    }
  });

  it('w:br w:type="page" 生成分页块而不是空行', async () => {
    const { doc, pkg } = await parseBody(
      '<w:p><w:r><w:t>第一页</w:t></w:r></w:p>' +
        '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' +
        '<w:p><w:r><w:t>第二页</w:t></w:r></w:p>',
    );
    try {
      expect(doc.blocks.filter((block) => block.type === 'pageBreak')).toHaveLength(1);
      // 功能：只有分页标记的段落不额外占行，避免多出空白页；时间：2026-10-31；作者：lq。
      expect(paragraphsOf(doc.blocks).map(paragraphText)).toEqual(['第一页', '第二页']);
    } finally {
      pkg.dispose();
    }
  });

  it('w:t 内的裸 CR/LF 按 Word 语义归一为空格', async () => {
    const { doc, pkg } = await parseBody(
      '<w:p><w:r><w:t xml:space="preserve">上一行\n下一行</w:t></w:r></w:p>',
    );
    try {
      const paragraph = paragraphsOf(doc.blocks)[0];
      expect(paragraphText(paragraph)).toBe('上一行 下一行');
      expect(paragraph.runs.some((run) => run.lineBreak)).toBe(false);
    } finally {
      pkg.dispose();
    }
  });

  it('未声明 xml:space="preserve" 的 w:t 丢弃首尾空白', async () => {
    const { doc, pkg } = await parseBody(
      '<w:p><w:r><w:t>  无保留空格  </w:t></w:r><w:r><w:t xml:space="preserve"> 保留空格 </w:t></w:r></w:p>',
    );
    try {
      const paragraph = paragraphsOf(doc.blocks)[0];
      // 功能：行尾多余空格会占掉行宽，导致 Word 里不换行的地方在这里提前换行；时间：2026-10-31；作者：lq。
      expect(paragraph.runs.map((run) => run.text)).toEqual(['无保留空格', ' 保留空格 ']);
    } finally {
      pkg.dispose();
    }
  });

  it('段内换行在保存后仍写回 w:br', async () => {
    const { doc, pkg } = await parseBody(
      '<w:p><w:r><w:t>第一行</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>第二行</w:t></w:r></w:p>',
    );
    try {
      const saved = await serializeDocx(doc, { name: 'br.docx' });
      const reparsedPkg = await loadOoxmlPackage(saved);
      try {
        const xml = await reparsedPkg.text('word/document.xml');
        expect(xml).toContain('<w:br/>');
        const reparsed = await parseDocxPackage(reparsedPkg, { fileName: 'br.docx' });
        const paragraph = paragraphsOf(reparsed.blocks)[0];
        expect(paragraph.runs.filter((run) => run.lineBreak)).toHaveLength(1);
        expect(paragraphText(paragraph)).toBe('第一行第二行');
      } finally {
        reparsedPkg.dispose();
      }
    } finally {
      pkg.dispose();
    }
  });
});

describe('DOCX 硬分页保留段内顺序', () => {
  const pageBreak = '<w:br w:type="page"/>';
  const textRun = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
  const sequence = (blocks: DocBlock[]) => blocks.map((block) =>
    block.type === 'paragraph' ? `paragraph:${paragraphText(block)}` : block.type,
  );

  it.each([
    ['同一 run', `<w:r><w:rPr><w:b/></w:rPr><w:t>封面</w:t>${pageBreak}<w:t>正文</w:t></w:r>`],
    ['不同 run', `${textRun('封面')}<w:r>${pageBreak}</w:r>${textRun('正文')}`],
  ])('%s 内分页把封面和正文拆成有序块', async (_name, inline) => {
    const { doc, pkg } = await parseBody(`<w:p>${inline}</w:p>`);
    try {
      expect(sequence(doc.blocks)).toEqual(['paragraph:封面', 'pageBreak', 'paragraph:正文']);
      expect(new Set(doc.blocks.map((block) => block.id)).size).toBe(doc.blocks.length);
      const paragraphs = paragraphsOf(doc.blocks);
      expect(paragraphs.every((paragraph) => paragraph.runs.every((run) => run.field !== '__PAGE_BREAK__'))).toBe(true);
      if (_name === '同一 run') {
        expect(paragraphs.map((paragraph) => paragraph.runs[0].style.bold)).toEqual([true, true]);
      }
    } finally {
      pkg.dispose();
    }
  });

  it.each([
    ['段落直接子节点', pageBreak],
    ['超链接直接子节点', `<w:hyperlink>${pageBreak}</w:hyperlink>`],
    ['内容控件直接子节点', `<w:sdt><w:sdtContent>${pageBreak}</w:sdtContent></w:sdt>`],
  ])('%s 的分页也保留位置', async (_name, inline) => {
    const { doc, pkg } = await parseBody(`<w:p>${textRun('封面')}${inline}${textRun('正文')}</w:p>`);
    try {
      expect(sequence(doc.blocks)).toEqual(['paragraph:封面', 'pageBreak', 'paragraph:正文']);
    } finally {
      pkg.dispose();
    }
  });

  it.each(['PAGE', 'NUMPAGES', 'SECTIONPAGES'])('%s 简式页码域不能覆盖内部硬分页标记', async (field) => {
    const { doc, pkg } = await parseBody(
      `<w:p><w:fldSimple w:instr="${field}"><w:r><w:t>1</w:t>${pageBreak}<w:t>2</w:t></w:r></w:fldSimple></w:p>`,
    );
    try {
      expect(sequence(doc.blocks)).toEqual(['paragraph:1', 'pageBreak', 'paragraph:2']);
      expect(paragraphsOf(doc.blocks).map((paragraph) => paragraph.runs[0].field)).toEqual([field, field]);
    } finally {
      pkg.dispose();
    }
  });

  it('连续与分散的多个分页符不折叠，换行和制表符留在原位置', async () => {
    const { doc, pkg } = await parseBody(
      `<w:p><w:r><w:t>封面</w:t><w:br/>${pageBreak}${pageBreak}<w:tab/><w:t>正文</w:t>${pageBreak}<w:t>附录</w:t><w:cr/></w:r></w:p>`,
    );
    try {
      expect(sequence(doc.blocks)).toEqual([
        'paragraph:封面', 'pageBreak', 'pageBreak', 'paragraph:正文', 'pageBreak', 'paragraph:附录',
      ]);
      const paragraphs = paragraphsOf(doc.blocks);
      expect(paragraphs[0].runs.at(-1)?.lineBreak).toBe(true);
      expect(paragraphs[1].runs[0].tab).toBe(true);
      expect(paragraphs[2].runs.at(-1)?.lineBreak).toBe(true);
    } finally {
      pkg.dispose();
    }
  });

  it.each([
    ['run 内', `<w:r>${pageBreak}${pageBreak}</w:r>`],
    ['段落下', `${pageBreak}${pageBreak}`],
    ['独立分页段', `<w:r>${pageBreak}</w:r></w:p><w:p><w:r>${pageBreak}</w:r>`],
  ])('%s 的纯连续分页不插入额外空段', async (_name, inline) => {
    const { doc, pkg } = await parseBody(`<w:p>${inline}</w:p>`);
    try {
      expect(sequence(doc.blocks)).toEqual(['pageBreak', 'pageBreak']);
    } finally {
      pkg.dispose();
    }
  });

  it.each([
    ['段首', `${pageBreak}${textRun('正文')}`, ['pageBreak', 'paragraph:正文']],
    ['段尾', `${textRun('正文')}${pageBreak}`, ['paragraph:正文', 'pageBreak']],
    ['首尾连续分页', `${pageBreak}${pageBreak}${textRun('正文')}${pageBreak}${pageBreak}`,
      ['pageBreak', 'pageBreak', 'paragraph:正文', 'pageBreak', 'pageBreak']],
  ])('%s 分页保留顺序且不占空行', async (_name, inline, expected) => {
    const { doc, pkg } = await parseBody(`<w:p>${inline}</w:p>`);
    try {
      expect(sequence(doc.blocks)).toEqual(expected);
    } finally {
      pkg.dispose();
    }
  });

  it('分页旁的真实空段与段内空行不会被当作纯分页丢弃', async () => {
    const { doc, pkg } = await parseBody(`<w:p/><w:p><w:cr/>${pageBreak}<w:tab/></w:p><w:p/>`);
    try {
      expect(sequence(doc.blocks)).toEqual(['paragraph:', 'paragraph:', 'pageBreak', 'paragraph:', 'paragraph:']);
      const paragraphs = paragraphsOf(doc.blocks);
      expect(paragraphs[1].runs[0].lineBreak).toBe(true);
      expect(paragraphs[2].runs[0].tab).toBe(true);
    } finally {
      pkg.dispose();
    }
  });

  it('w:lastRenderedPageBreak 只是渲染缓存，不是硬分页或换行', async () => {
    const { doc, pkg } = await parseBody(
      `<w:p><w:r><w:t>封面</w:t><w:lastRenderedPageBreak/><w:t>正文</w:t></w:r><w:lastRenderedPageBreak/>${textRun('附录')}</w:p>`,
    );
    try {
      expect(sequence(doc.blocks)).toEqual(['paragraph:封面正文附录']);
      expect(paragraphsOf(doc.blocks)[0].runs.some((run) => run.lineBreak)).toBe(false);
    } finally {
      pkg.dispose();
    }
  });

  it('样式间距仅落首尾部分，续段不重复段前分页和列表编号，保存重开仍如此', async () => {
    const parts = {
      'word/styles.xml': '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Split"><w:name w:val="拆分段落"/><w:pPr><w:pageBreakBefore/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="180" w:after="360" w:line="360"/><w:ind w:left="720" w:hanging="360"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/></w:rPr></w:style></w:styles>',
      'word/numbering.xml': '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
    };
    const { doc, pkg, bytes } = await parseBody(
      `<w:p><w:pPr><w:pStyle w:val="Split"/></w:pPr>${textRun('第一部分')}${pageBreak}${textRun('第二部分')}${pageBreak}${textRun('第三部分')}</w:p>` +
        `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${textRun('下一项')}</w:p>`,
      parts,
    );
    const check = (blocks: DocBlock[]) => {
      const paragraphs = paragraphsOf(blocks);
      expect(paragraphs).toHaveLength(4);
      expect(paragraphs.slice(0, 3).map((paragraph) => paragraph.style.spaceBeforePx)).toEqual([12, 0, 0]);
      expect(paragraphs.slice(0, 3).map((paragraph) => paragraph.style.spaceAfterPx)).toEqual([0, 0, 24]);
      expect(paragraphs.slice(0, 3).map((paragraph) => !!paragraph.style.pageBreakBefore)).toEqual([true, false, false]);
      expect(paragraphs.slice(0, 3).map((paragraph) => paragraph.style.firstLineIndentPx)).toEqual([-24, 0, 0]);
      expect(paragraphs.map((paragraph) => paragraph.list?.marker)).toEqual(['1.', undefined, undefined, '2.']);
      for (const paragraph of paragraphs.slice(0, 3)) {
        expect(paragraph.style).toMatchObject({ alignment: 'center', indentLeftPx: 48, lineHeight: 1.5 });
        expect(paragraph.runs[0].style.italic).toBe(true);
      }
    };
    try {
      check(doc.blocks);
      const reopenedPkg = await loadOoxmlPackage(await serializeDocx(doc, { sourceBytes: bytes, name: 'list-break.docx' }));
      try {
        const reopened = await parseDocxPackage(reopenedPkg);
        expect(sequence(reopened.blocks)).toEqual(sequence(doc.blocks));
        check(reopened.blocks);
      } finally {
        reopenedPkg.dispose();
      }
    } finally {
      pkg.dispose();
    }
  });

  it.each([
    ['段中', `${textRun('封面')}${pageBreak}${textRun('正文')}`, ['paragraph:封面', 'pageBreak', 'paragraph:正文'], '正文'],
    ['段首', `${pageBreak}${textRun('正文')}`, ['pageBreak', 'paragraph:正文'], '正文'],
    ['段尾', `${textRun('正文')}${pageBreak}`, ['paragraph:正文', 'pageBreak', 'paragraph:'], ''],
    ['仅连续分页', `${pageBreak}${pageBreak}`, ['pageBreak', 'pageBreak', 'paragraph:'], ''],
  ])('%s 分页的 sectionEnd 只在最后部分且保存重开保留位置', async (_name, inline, expected, lastText) => {
    const { doc, pkg } = await parseBody(
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgNumType w:start="7"/></w:sectPr></w:pPr>' +
        `${inline}</w:p><w:p>${textRun('下一节')}</w:p>`,
    );
    const check = (blocks: DocBlock[]) => {
      expect(sequence(blocks)).toEqual([...expected, 'paragraph:下一节']);
      const paragraphs = paragraphsOf(blocks);
      const withSection = paragraphs.filter((paragraph) => paragraph.sectionEnd);
      expect(withSection).toHaveLength(1);
      expect(withSection[0]).toBe(paragraphs.at(-2));
      expect(paragraphText(withSection[0])).toBe(lastText);
      expect(withSection[0].sectionEnd?.page).toMatchObject({ sizeName: 'Letter', pageNumberStart: 7 });
    };
    try {
      check(doc.blocks);
      const reopenedPkg = await loadOoxmlPackage(await serializeDocx(doc, { name: 'section-break.docx' }));
      try {
        check((await parseDocxPackage(reopenedPkg)).blocks);
      } finally {
        reopenedPkg.dispose();
      }
    } finally {
      pkg.dispose();
    }
  });

  it('页眉、页脚和表格兼容拆分后的段落块，页码域仍可识别', async () => {
    const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const { doc, pkg } = await parseBody(
      `<w:tbl><w:tr><w:tc><w:p>${textRun('表格上')}${pageBreak}${textRun('表格下')}</w:p></w:tc></w:tr></w:tbl>` +
        '<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:pPr></w:p>',
      {
        'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>',
        'word/header1.xml': `<w:hdr xmlns:w="${ns}"><w:p>${textRun('页眉上')}${pageBreak}<w:fldSimple w:instr="PAGE">${textRun('1')}</w:fldSimple></w:p></w:hdr>`,
        'word/footer1.xml': `<w:ftr xmlns:w="${ns}"><w:p>${pageBreak}${textRun('页脚')}</w:p></w:ftr>`,
      },
    );
    try {
      expect(sequence(doc.header!.blocks)).toEqual(['paragraph:页眉上', 'pageBreak', 'paragraph:1']);
      expect(doc.header!.pageNumbers).toBe(true);
      expect(sequence(doc.footer!.blocks)).toEqual(['pageBreak', 'paragraph:页脚']);
      const table = doc.blocks[0];
      expect(table.type).toBe('table');
      if (table.type === 'table') {
        expect(sequence(table.rows[0].cells[0].blocks)).toEqual(['paragraph:表格上', 'pageBreak', 'paragraph:表格下']);
      }
    } finally {
      pkg.dispose();
    }
  });

  it.each([
    ['原有独立分页段', `<w:p>${textRun('封面')}</w:p><w:p><w:r>${pageBreak}</w:r></w:p><w:p>${textRun('正文')}</w:p>`,
      ['paragraph:封面', 'pageBreak', 'paragraph:正文', 'paragraph:后续']],
    ['段内连续及首尾分页', `<w:p>${pageBreak}${textRun('封面')}<w:r>${pageBreak}${pageBreak}</w:r>${textRun('正文')}${pageBreak}</w:p>`,
      ['pageBreak', 'paragraph:封面', 'pageBreak', 'pageBreak', 'paragraph:正文', 'pageBreak', 'paragraph:后续']],
  ])('%s 保存两次仍是独立分页段且不增减空段', async (_name, body, expected) => {
    const { doc, pkg } = await parseBody(`${body}<w:p>${textRun('后续')}</w:p>`);
    try {
      expect(sequence(doc.blocks)).toEqual(expected);
      let current = doc;
      for (let pass = 0; pass < 2; pass += 1) {
        const reopenedPkg = await loadOoxmlPackage(await serializeDocx(current, { name: 'page-break.docx' }));
        try {
          const dom = parseXmlDom(await reopenedPkg.text('word/document.xml'));
          const breakParagraphs = Array.from(dom.getElementsByTagName('w:p')).filter((paragraph) =>
            Array.from(paragraph.getElementsByTagName('w:br')).some((br) => br.getAttribute('w:type') === 'page'),
          );
          expect(breakParagraphs).toHaveLength(expected.filter((part) => part === 'pageBreak').length);
          for (const paragraph of breakParagraphs) {
            expect(paragraph.getElementsByTagName('w:br')).toHaveLength(1);
            expect(paragraph.getElementsByTagName('w:t')).toHaveLength(0);
            expect(paragraph.getElementsByTagName('w:spacing')[0].getAttribute('w:line')).toBe('1');
          }
          current = await parseDocxPackage(reopenedPkg);
          expect(sequence(current.blocks)).toEqual(expected);
        } finally {
          reopenedPkg.dispose();
        }
      }
    } finally {
      pkg.dispose();
    }
  });
});

describe('浮动图片锚定不影响后续段落', () => {
  it('解析签名下划线图片的锚定距离与环绕方向', async () => {
    const { doc, pkg } = await parseFixture('tests/fixtures/docx/contract-underline.docx');
    try {
      const paragraphs = paragraphsOf(doc.blocks);
      const imageParagraph = paragraphs.find((paragraph) => paragraph.runs.some((run) => run.image));
      expect(imageParagraph).toBeTruthy();
      const image = imageParagraph!.runs.find((run) => run.image)!.image!;
      expect(image.wrap).toBe('square-left');
      expect(image.anchor).toMatchObject({
        hRelative: 'column',
        hOffsetPx: 0,
        vRelative: 'paragraph',
        wrapType: 'square',
        wrapText: 'left',
        distTopPx: 0,
        distBottomPx: 0,
        distLeftPx: 12,
        distRightPx: 12,
      });
      // 功能：下划线细条改为绝对定位，完全脱离行内流，不再挤占后续段落的行盒；时间：2026-10-31；作者：lq。
      const html = renderRunsHtml([imageParagraph!.runs.find((run) => run.image)!], { indentLeftPx: 0 });
      expect(html).toContain('position:absolute');
      expect(html).toContain('left:0px');
      expect(html).toContain('top:0px');
      expect(html).not.toContain('float:left');
      expect(html).not.toContain('margin:0px 12px 0px 0px');
    } finally {
      pkg.dispose();
    }
  });

  it('图片后的乙方段落保持单段单行文本', async () => {
    const { doc, pkg } = await parseFixture('tests/fixtures/docx/contract-underline.docx');
    try {
      const paragraphs = paragraphsOf(doc.blocks);
      const imageIndex = paragraphs.findIndex((paragraph) => paragraph.runs.some((run) => run.image));
      const nextText = paragraphs
        .slice(imageIndex + 1)
        .map(paragraphText)
        .find((text) => text.trim().length > 0);
      expect(nextText?.replace(/\s/g, '')).toBe('乙方：深圳市XXX有限公司');
    } finally {
      pkg.dispose();
    }
  });

  it('浮动图片保存后仍保留锚定与环绕参数', async () => {
    const { doc, bytes, pkg } = await parseFixture('tests/fixtures/docx/contract-underline.docx');
    pkg.dispose();
    const saved = await serializeDocx(doc, { sourceBytes: bytes, name: 'contract-underline.docx' });
    const reparsedPkg = await loadOoxmlPackage(saved);
    try {
      const xml = await reparsedPkg.text('word/document.xml');
      expect(xml).toContain('relativeFrom="column"');
      expect(xml).toContain('<wp:wrapSquare wrapText="left"/>');
      expect(xml).not.toContain('distB="76200"');
      const reparsed = await parseDocxPackage(reparsedPkg, { fileName: 'contract-underline.docx' });
      const imageParagraph = paragraphsOf(reparsed.blocks).find((paragraph) =>
        paragraph.runs.some((run) => run.image),
      );
      expect(imageParagraph?.runs.find((run) => run.image)?.image?.anchor?.distBottomPx).toBe(0);
    } finally {
      reparsedPkg.dispose();
    }
  });
});

describe('浮动图定位策略', () => {
  const imageOf = (over: Partial<DocImage> = {}): DocImage => ({
    id: 'img-test',
    src: 'data:image/png;base64,AAAA',
    widthPx: 240,
    heightPx: 120,
    ...over,
  });

  // 功能：回归纯图片自动行距不重复预留下伸部空间，同时保留显式行距、网格与混排语义；时间：2026-11-05；作者：lq。
  it('纯内嵌图片自动行距清除文字行盒，文字混排、网格和固定/最小行距保持原规则', () => {
    const paragraph: DocParagraph = { id: 'inline-only', type: 'paragraph', style: {},
      runs: [{ style: {}, image: imageOf() }, { style: {}, image: imageOf({ heightPx: 80 }) }] };
    expect(paragraphTextStyle(paragraph, 'Calibri', 11).lineHeight).toBe('0px');
    for (const style of [{ lineHeightExactPx: 16 }, { lineHeightAtLeastPx: 0 }]) {
      expect(paragraphTextStyle({ ...paragraph, style }, 'Calibri', 11).lineHeight).not.toBe('0px');
    }
    expect(paragraphTextStyle(paragraph, 'Calibri', 11, 24).lineHeight).toBe('24px');
    expect(paragraphTextStyle({ ...paragraph, runs: [...paragraph.runs, { text: ' tail text', style: {} }] }, 'Calibri', 11).lineHeight).not.toBe('0px');
    expect(paragraphTextStyle({ ...paragraph, runs: [] }, 'Calibri', 11).lineHeight).not.toBe('0px');
  });

  // 功能：图片倍数行距只增加字体行高的增量，不叠加整份文字下伸部，也不按倍数拉伸图片；时间：2026-11-05；作者：lq。
  it('纯图片的自动倍数保留额外间距，低倍数不产生负占位', () => {
    const paragraph: DocParagraph = { id: 'image-multiple', type: 'paragraph', style: { lineHeight: 1.15 },
      runs: [{ style: {}, image: imageOf() }] };
    const style = paragraphTextStyle(paragraph, 'Calibri', 11);
    expect(style.lineHeight).toBe('0px');
    const extra = (value: ReturnType<typeof paragraphTextStyle>) => Number.parseFloat(
      (value as Record<string, string>)['--mo-inline-image-line-extra']);
    expect(extra(style)).toBeCloseTo(singleLineHeightPx('Calibri', 11) * 0.15, 5);
    expect(extra(paragraphTextStyle({ ...paragraph, style: { lineHeight: 0.5 } }, 'Calibri', 11))).toBe(0);
    expect(extra(paragraphTextStyle({ ...paragraph, style: { lineHeight: 1 } }, 'Calibri', 11))).toBe(0);
  });

  it('普通四周型图片仍走 float，保留文字绕排能力', () => {
    const html = renderRunsHtml([
      { text: '', style: {}, image: imageOf({ wrap: 'square-left', anchor: { wrapType: 'square', hRelative: 'column', hOffsetPx: 0 } }) },
    ]);
    // 功能：手柄本身是 position:absolute，这里只断言「图片帧」的定位方式；时间：2026-10-31；作者：lq。
    expect(html).toContain('height:120px;position:relative;float:left');
  });

  it('wp:wrapNone（浮于文字上方）改为绝对定位，不参与行内流', () => {
    const html = renderRunsHtml([
      { text: '', style: {}, image: imageOf({ wrap: 'square-left', anchor: { wrapType: 'none', hRelative: 'column', hOffsetPx: 36, vOffsetPx: 12 } }) },
    ]);
    expect(html).toContain('height:120px;position:absolute;left:36px;top:12px;z-index:3');
    expect(html).not.toContain('float:left');
  });

  it('装饰性细条（高 <= 24px 且长宽比 >= 8）改为绝对定位', () => {
    const html = renderRunsHtml([
      { text: '', style: {}, image: imageOf({ widthPx: 387, heightPx: 8, wrap: 'square-left', anchor: { wrapType: 'square', hRelative: 'column', hOffsetPx: 0 } }) },
    ]);
    expect(html).toContain('height:8px;position:absolute;left:0px;top:0px;z-index:3');
  });

  it('衬于文字下方的图片放到正文之下', () => {
    const html = renderRunsHtml([
      { text: '', style: {}, image: imageOf({ wrap: 'square-left', anchor: { wrapType: 'none', behindDoc: true } }) },
    ]);
    expect(html).toContain('height:120px;position:absolute;left:0px;top:0px;z-index:0');
  });

  it('段落盒提供定位上下文，并保住固定行距的行盒高度', () => {
    const paragraph: DocParagraph = {
      type: 'paragraph',
      id: 'p1',
      runs: [{ text: '', style: {}, image: imageOf({ widthPx: 387, heightPx: 8, wrap: 'square-left', anchor: { wrapType: 'square' } }) }],
      style: { lineHeightExactPx: 1.33 },
    };
    const style = paragraphTextStyle(paragraph, '宋体', 10.5);
    expect(style.position).toBe('relative');
    expect(style.minHeight).toBe('1.33px');
    // 功能：只含浮动图的段落不能因为绝对定位塌成 0 高，渲染仍输出图片帧；时间：2026-10-31；作者：lq。
    expect(renderParagraphHtml(paragraph)).toContain('data-image-frame');
  });
});
