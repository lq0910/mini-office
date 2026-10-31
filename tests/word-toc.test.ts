/**
 * 功能：验证 Word 目录(TOC)兼容——前导点+页码右对齐渲染（标准 TOC 样式与手动目录制表位两种）、
 * 目录项锚点链接、书签采集、w:leader 解析与内置 TOC1~TOC9 样式；时间：2026-10-30；作者：hou0。
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createBlankDoc, type DocParagraph } from '@mini-office/document-model';
import {
  BUILTIN_STYLES,
  parseDocx,
  paragraphTextStyle,
  renderParagraphHtml,
  serializeDocx,
} from '@mini-office/editor-word';

function tocParagraph(): DocParagraph {
  return {
    type: 'paragraph',
    id: 'p1',
    styleId: 'TOC1',
    style: {},
    runs: [
      { text: '第一章 引言', style: {}, anchor: '_Toc1' },
      { text: '', style: {}, tab: true },
      { text: '1', style: {} },
    ],
  };
}

// 功能：手动目录段落（如 WPS 用自定义样式 + 右对齐点前导符制表位），非标准 TOC 样式；时间：2026-10-30；作者：hou0。
function manualTocParagraph(): DocParagraph {
  return {
    type: 'paragraph',
    id: 'p2',
    styleId: 'MyDir',
    style: { tabStops: [{ posPx: 400, kind: 'right', leader: 'dot' }] },
    runs: [
      { text: '1概述', style: {} },
      { text: '', style: {}, tab: true },
      { text: '3', style: {} },
    ],
  };
}

describe('Word 目录(TOC)兼容', () => {
  it('renderParagraphHtml 把 TOC 段落渲染为「标题+前导点+页码右对齐」结构', () => {
    const html = renderParagraphHtml(tocParagraph());
    expect(html).toContain('mo-toc-title');
    expect(html).toContain('mo-toc-leader');
    expect(html).toContain('mo-toc-page');
  });

  it('标准 TOC 样式兜底：无制表位信息也走前导点布局', () => {
    const style = paragraphTextStyle(tocParagraph(), '宋体', 10.5);
    expect(style.display).toBe('flex');
  });

  it('手动目录（自定义样式+右对齐点前导符制表位）同样渲染为前导点+页码右对齐', () => {
    const html = renderParagraphHtml(manualTocParagraph());
    expect(html).toContain('mo-toc-title');
    expect(html).toContain('mo-toc-leader');
    expect(html).toContain('mo-toc-page');
    expect(html).toContain('1概述');
    expect(html).toContain('3');
    const style = paragraphTextStyle(manualTocParagraph(), '宋体', 10.5);
    expect(style.display).toBe('flex');
  });

  it('目录项锚点渲染为可点击链接', () => {
    const html = renderParagraphHtml(tocParagraph());
    expect(html).toContain('data-anchor="_Toc1"');
    expect(html).toContain('第一章 引言');
    expect(html).toContain('mo-toc-page');
  });

  it('非目录段落（无 leader 制表位）不启用 flex，保持原有行内布局', () => {
    const normal: DocParagraph = {
      type: 'paragraph',
      id: 'p3',
      styleId: 'Normal',
      style: {},
      runs: [{ text: '正文', style: {} }],
    };
    const style = paragraphTextStyle(normal, '宋体', 10.5);
    expect(style.display).toBeUndefined();
    expect(renderParagraphHtml(normal)).not.toContain('mo-toc-leader');
  });

  it('带右对齐制表位但无 leader 的普通段落不渲染前导点（避免误判为目录）', () => {
    const plain: DocParagraph = {
      type: 'paragraph',
      id: 'p4',
      styleId: 'Normal',
      style: { tabStops: [{ posPx: 400, kind: 'right' }] },
      runs: [
        { text: '左', style: {} },
        { text: '', style: {}, tab: true },
        { text: '右', style: {} },
      ],
    };
    const html = renderParagraphHtml(plain);
    expect(html).not.toContain('mo-toc-leader');
  });

  it('BUILTIN_STYLES 包含 TOC1~TOC9 内置目录样式', () => {
    for (let level = 1; level <= 9; level += 1) {
      expect(BUILTIN_STYLES.some((s) => s.id === `TOC${level}`)).toBe(true);
    }
  });

  it('parseDocx 解析 w:leader 前导符与右对齐制表位，供目录渲染', async () => {
    const doc = createBlankDoc();
    doc.blocks = [{ type: 'paragraph', id: 'p0', runs: [{ text: '占位', style: {} }], style: {} }];
    const bytes = await serializeDocx(doc, { name: 'toc.docx' });
    const zip = await JSZip.loadAsync(bytes);
    let docXml = await zip.file('word/document.xml')!.async('text');
    const manualXml =
      '<w:p><w:pPr><w:pStyle w:val="MyDir"/><w:tabs><w:tab w:val="right" w:pos="9076" w:leader="dot"/></w:tabs></w:pPr>' +
      '<w:r><w:t>1概述</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r>' +
      '</w:p>';
    docXml = docXml.replace('</w:body>', `${manualXml}</w:body>`);
    zip.file('word/document.xml', docXml);
    const reopened = await parseDocx(await zip.generateAsync({ type: 'uint8array' }), { fileName: 'toc.docx' });
    const manual = reopened.blocks.find(
      (b): b is DocParagraph => b.type === 'paragraph' && b.styleId === 'MyDir',
    );
    expect(manual).toBeDefined();
    if (!manual) return;
    expect(manual.style.tabStops?.[0]?.leader).toBe('dot');
    expect(manual.style.tabStops?.[0]?.kind).toBe('right');
    expect(renderParagraphHtml(manual)).toContain('mo-toc-leader');
  });

  it('parseDocx 采集目录项超链接锚点与书签，供点击跳转', async () => {
    const doc = createBlankDoc();
    doc.blocks = [{ type: 'paragraph', id: 'p0', runs: [{ text: '占位', style: {} }], style: {} }];
    const bytes = await serializeDocx(doc, { name: 'toc.docx' });
    const zip = await JSZip.loadAsync(bytes);
    let docXml = await zip.file('word/document.xml')!.async('text');
    const tocXml =
      '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
      '<w:bookmarkStart w:id="1" w:name="_Toc_Title1"/>' +
      '<w:hyperlink w:anchor="_Toc_Title1"><w:r><w:t>标题一</w:t></w:r></w:hyperlink>' +
      '<w:r><w:tab/></w:r>' +
      '<w:r><w:t>3</w:t></w:r>' +
      '<w:bookmarkEnd w:id="1"/>' +
      '</w:p>';
    docXml = docXml.replace('</w:body>', `${tocXml}</w:body>`);
    zip.file('word/document.xml', docXml);
    const reopened = await parseDocx(await zip.generateAsync({ type: 'uint8array' }), { fileName: 'toc.docx' });
    const toc = reopened.blocks.find(
      (b): b is DocParagraph => b.type === 'paragraph' && b.styleId === 'TOC1',
    );
    expect(toc).toBeDefined();
    if (!toc) return;
    expect(toc.bookmarks).toContain('_Toc_Title1');
    const linkRun = toc.runs.find((r) => r.anchor === '_Toc_Title1');
    expect(linkRun).toBeDefined();
    expect(linkRun?.text).toBe('标题一');
  });
});

describe('目录显示与保存', () => {
  it('serialize 写出制表位前导符 w:leader，Word/WPS 打开仍有点线', async () => {
    const doc = createBlankDoc();
    doc.blocks = [
      {
        type: 'paragraph',
        id: 'toc1',
        styleId: 'TOC1',
        style: { tabStops: [{ posPx: 620, kind: 'right', leader: 'dot' }] },
        runs: [
          { text: '第一章 总则', style: {} },
          { text: '', style: {}, tab: true },
          { text: '3', style: {} },
        ],
      },
    ];
    const bytes = await serializeDocx(doc, { name: 'toc.docx' });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')!.async('text');
    expect(xml).toContain('w:leader="dot"');
  });
});

describe('中西文字体（w:rFonts 双槽）', () => {
  it('parse 分离西文与中文字体，渲染时中文字体进 CSS 栈', async () => {
    const doc = createBlankDoc();
    doc.blocks = [{ type: 'paragraph', id: 'p0', runs: [{ text: '占位', style: {} }], style: {} }];
    const bytes = await serializeDocx(doc, { name: 'font.docx' });
    const zip = await JSZip.loadAsync(bytes);
    let docXml = await zip.file('word/document.xml')!.async('text');
    const xml =
      '<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体"/></w:rPr></w:pPr>' +
      '<w:r><w:t>混合文字</w:t></w:r></w:p>';
    docXml = docXml.replace('</w:body>', `${xml}</w:body>`);
    zip.file('word/document.xml', docXml);
    const reopened = await parseDocx(await zip.generateAsync({ type: 'uint8array' }), { fileName: 'font.docx' });
    const p = reopened.blocks.find(
      (b): b is DocParagraph => b.type === 'paragraph' && b.runs.some((r) => r.text === '混合文字'),
    );
    const run = p?.runs.find((r) => r.text === '混合文字');
    expect(run?.style.fontFamily).toBe('Times New Roman');
    expect(run?.style.eastAsiaFont).toBe('黑体');
    expect(renderParagraphHtml(p!)).toContain('Heiti SC');
  });

  it('run 级缺 eastAsia 时继承样式层的中文字体，不被西文覆盖', async () => {
    const doc = createBlankDoc();
    doc.blocks = [{ type: 'paragraph', id: 'p0', runs: [{ text: '占位', style: {} }], style: {} }];
    const bytes = await serializeDocx(doc, { name: 'font.docx' });
    const zip = await JSZip.loadAsync(bytes);
    let stylesXml = await zip.file('word/styles.xml')!.async('text');
    const styleDef =
      '<w:style w:type="paragraph" w:styleId="DirFont"><w:name w:val="DirFont"/>' +
      '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体"/></w:rPr></w:style>';
    stylesXml = stylesXml.replace('</w:styles>', `${styleDef}</w:styles>`);
    zip.file('word/styles.xml', stylesXml);
    let docXml = await zip.file('word/document.xml')!.async('text');
    const bodyXml =
      '<w:p><w:pPr><w:pStyle w:val="DirFont"/></w:pPr>' +
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman"/></w:rPr><w:t>目录项</w:t></w:r></w:p>';
    docXml = docXml.replace('</w:body>', `${bodyXml}</w:body>`);
    zip.file('word/document.xml', docXml);
    const reopened = await parseDocx(await zip.generateAsync({ type: 'uint8array' }), { fileName: 'font.docx' });
    const p = reopened.blocks.find(
      (b): b is DocParagraph => b.type === 'paragraph' && b.styleId === 'DirFont',
    );
    const run = p?.runs.find((r) => r.text === '目录项');
    expect(run?.style.fontFamily).toBe('Times New Roman');
    expect(run?.style.eastAsiaFont).toBe('黑体');
  });

  it('serialize 分槽回写 w:ascii 与 w:eastAsia', async () => {
    const doc = createBlankDoc();
    doc.blocks = [
      {
        type: 'paragraph',
        id: 'p1',
        style: {},
        runs: [{ text: '混合', style: { fontFamily: 'Times New Roman', eastAsiaFont: '黑体' } }],
      },
    ];
    const bytes = await serializeDocx(doc, { name: 'font.docx' });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')!.async('text');
    expect(xml).toContain('w:ascii="Times New Roman"');
    expect(xml).toContain('w:eastAsia="黑体"');
  });
});

describe('目录域保真（docx 往返）', () => {
  // 功能：Word 标准目录结构——TOC 域跨段落（begin/separate 在首段、end 在末段），条目页码是 PAGEREF 子域；时间：2026-10-30；作者：hou0。
  const TOC_FIELD_XML =
    '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:hyperlink w:anchor="_Toc_Title1"><w:r><w:t>标题一</w:t></w:r></w:hyperlink>' +
    '<w:r><w:tab/></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc_Title1 \\h </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:t>3</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '</w:p>' +
    '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';

  async function roundTrip(bodyXml: string): Promise<string> {
    const doc = createBlankDoc();
    doc.blocks = [{ type: 'paragraph', id: 'p0', runs: [{ text: '占位', style: {} }], style: {} }];
    const bytes = await serializeDocx(doc, { name: 'toc.docx' });
    const zip = await JSZip.loadAsync(bytes);
    let docXml = await zip.file('word/document.xml')!.async('text');
    docXml = docXml.replace('</w:body>', `${bodyXml}</w:body>`);
    zip.file('word/document.xml', docXml);
    const reopened = await parseDocx(await zip.generateAsync({ type: 'uint8array' }), { fileName: 'toc.docx' });
    const out = await serializeDocx(reopened, { name: 'toc.docx' });
    const outZip = await JSZip.loadAsync(out);
    return outZip.file('word/document.xml')!.async('text');
  }

  it('保存后仍保留 TOC 域结构（begin/separate/end + instrText），Word 打开后仍可更新目录', async () => {
    const xml = await roundTrip(TOC_FIELD_XML);
    expect(xml).toContain('w:fldCharType="begin"');
    expect(xml).toContain('w:fldCharType="separate"');
    expect(xml).toContain('w:fldCharType="end"');
    expect(xml).toContain('TOC \\o &quot;1-3&quot; \\h \\z \\u');
    expect(xml).toContain('PAGEREF _Toc_Title1');
  });

  it('保存后仍保留目录项超链接（w:hyperlink w:anchor），Word 内可点击跳转', async () => {
    const xml = await roundTrip(TOC_FIELD_XML);
    expect(xml).toContain('w:hyperlink');
    expect(xml).toContain('w:anchor="_Toc_Title1"');
  });

  it('保存后仍保留锚点目标书签（w:bookmarkStart/w:bookmarkEnd），否则目录跳转失效', async () => {
    const withBookmark =
      '<w:p><w:bookmarkStart w:id="7" w:name="_Toc_Title1"/><w:r><w:t>标题一</w:t></w:r>' +
      '<w:bookmarkEnd w:id="7"/></w:p>';
    const xml = await roundTrip(withBookmark);
    expect(xml).toContain('<w:bookmarkStart');
    expect(xml).toContain('w:name="_Toc_Title1"');
    expect(xml).toContain('<w:bookmarkEnd');
  });

  it('保真写回不加 w:dirty，避免 Word 打开即重排静态页码缓存', async () => {
    const xml = await roundTrip(TOC_FIELD_XML);
    expect(xml).not.toContain('w:dirty');
  });

  it('目录条目文字与页码不因写回域结构而丢失', async () => {
    const xml = await roundTrip(TOC_FIELD_XML);
    expect(xml).toContain('标题一');
    expect(xml).toContain('<w:tab/>');
  });

  it('合并域仍只写一个 fldSimple，不与复杂域结构重复', async () => {
    const mergeXml =
      '<w:p>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> MERGEFIELD "name" \\* MERGEFORMAT </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>«name»</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '</w:p>';
    const xml = await roundTrip(mergeXml);
    expect(xml).toContain('<w:fldSimple w:instr=" MERGEFIELD');
    expect(xml).not.toContain('w:fldChar');
  });
});
