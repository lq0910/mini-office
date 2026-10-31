// 临时验证 2a：首页/奇偶页页眉页脚
import { describe, expect, it } from 'vitest';
import { applyRevisionDecision, documentSections, parseDocx, renderParagraphHtml, serializeDocx } from '@mini-office/editor-word';
import JSZip from 'jszip';

async function makeDocx(body: string, sectPr: string, settings?: string, docNs?: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rH2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/><Relationship Id="rF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>');
  zip.file('word/header1.xml', '<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>首页页眉</w:t></w:r></w:p></w:hdr>');
  zip.file('word/header2.xml', '<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>默认页眉</w:t></w:r></w:p></w:hdr>');
  zip.file('word/footer1.xml', '<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>偶数页脚</w:t></w:r></w:p></w:ftr>');
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${docNs ?? 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'}><w:body>${body}${sectPr}</w:body></w:document>`);
  if (settings) zip.file('word/settings.xml', settings);
  return zip.generateAsync({ type: 'uint8array' });
}

const SECT = '<w:sectPr><w:pgSz w:w="11910" w:h="16840"/><w:pgMar w:top="1580" w:bottom="280" w:left="880" w:right="680"/></w:sectPr>';

describe('首页/奇偶页页眉页脚', () => {
  it('解析 titlePg + first/even 引用', async () => {
    const sect = '<w:sectPr><w:titlePg/><w:pgSz w:w="11910" w:h="16840"/><w:pgMar w:top="1580" w:bottom="280" w:left="880" w:right="680"/><w:headerReference w:type="default" r:id="rH2"/><w:headerReference w:type="first" r:id="rH1"/><w:footerReference w:type="even" r:id="rF1"/></w:sectPr>';
    const doc = await parseDocx(await makeDocx('<w:p><w:r><w:t>正文</w:t></w:r></w:p>', sect, '<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:evenAndOddHeaders/></w:settings>'));
    expect(doc.page.titlePage).toBe(true);
    expect(doc.evenAndOddHeaders).toBe(true);
    const section = documentSections(doc)[0];
    console.log('DEBUG header=', JSON.stringify(section.header?.blocks), 'first=', JSON.stringify(section.headerFirst?.blocks?.length), 'rels=', doc.images);
    expect(section.header?.blocks[0].type === 'paragraph' && (section.header.blocks[0] as any).runs[0].text).toBe('默认页眉');
    expect(section.headerFirst?.blocks[0].type === 'paragraph' && (section.headerFirst!.blocks[0] as any).runs[0].text).toBe('首页页眉');
    expect(section.footerEven).toBeTruthy();
  });

  it('往返保存后 titlePg 与变体引用保留', async () => {
    const sect = '<w:sectPr><w:titlePg/><w:pgSz w:w="11910" w:h="16840"/><w:pgMar w:top="1580" w:bottom="280" w:left="880" w:right="680"/><w:headerReference w:type="default" r:id="rH2"/><w:headerReference w:type="first" r:id="rH1"/></w:sectPr>';
    const doc = await parseDocx(await makeDocx('<w:p><w:r><w:t>正文</w:t></w:r></w:p>', sect));
    const saved = await serializeDocx(doc, { name: 'hf.docx' });
    const reopened = await parseDocx(saved);
    const section = documentSections(reopened)[0];
    expect(reopened.page.titlePage).toBe(true);
    expect(section.headerFirst).toBeTruthy();
    expect(section.header).toBeTruthy();
  });

  it('文本框：解析（AlternateContent/DrawingML）、渲染与只读往返', async () => {
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const body = `<w:p><w:r>${''}<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="1828800" cy="914400"/><wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>91440</wp:posOffset></wp:positionV><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>文本框内容</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape style="position:absolute;margin-left:72pt;margin-top:7.2pt;width:144pt;height:72pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>文本框内容</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r><w:r><w:t>正文</w:t></w:r></w:p>`;
    const doc = await parseDocx(await makeDocx(body, '<w:sectPr><w:pgSz w:w="11910" w:h="16840"/><w:pgMar w:top="1580" w:bottom="280" w:left="880" w:right="680"/></w:sectPr>', undefined, NS));
    const tbRun = doc.blocks.flatMap((b) => (b.type === 'paragraph' ? b.runs : [])).find((r) => r.textbox);
    console.log('DEBUG runs:', JSON.stringify(doc.blocks.flatMap((b) => (b.type === 'paragraph' ? b.runs.map((r) => ({ t: r.text, tb: !!r.textbox, img: !!r.image })) : []))));
    expect(tbRun?.textbox).toBeTruthy();
    expect(Math.round(tbRun!.textbox!.left)).toBe(96);
    expect(Math.round(tbRun!.textbox!.top)).toBe(10);
    const text = tbRun!.textbox!.blocks.map((b) => (b.type === 'paragraph' ? b.runs.map((r) => r.text ?? '').join('') : '')).join('');
    expect(text).toBe('文本框内容');

    // 渲染：绝对定位容器 + 内容
    const html = renderParagraphHtml(doc.blocks[0] as never);
    expect(html).toContain('data-textbox="1"');
    expect(html).toContain('文本框内容');

    // 往返：保存为 VML v:shape，重开仍解析出文本框
    const saved = await serializeDocx(doc, { name: 'tb.docx' });
    const reopened = await parseDocx(saved);
    const reopenedTb = reopened.blocks.flatMap((b) => (b.type === 'paragraph' ? b.runs : [])).find((r) => r.textbox);
    expect(reopenedTb?.textbox?.blocks.length).toBeGreaterThan(0);
    const reopenedText = reopenedTb!.textbox!.blocks.map((b) => (b.type === 'paragraph' ? b.runs.map((r) => r.text ?? '').join('') : '')).join('');
    expect(reopenedText).toBe('文本框内容');
  });
});

describe('批注与修订', () => {
  it('批注：comments.xml 解析、范围高亮渲染、编辑回读与写回往返', async () => {
    const body =
      '<w:p><w:commentRangeStart w:id="1"/>' +
      '<w:r><w:t>被批注文字</w:t></w:r>' +
      '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>' +
      '<w:r><w:t>正文</w:t></w:r></w:p>';
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>');
    zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file('word/_rels/document.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rC1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>');
    zip.file('word/comments.xml', '<?xml version="1.0"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1" w:author="张三" w:date="2026-10-31T00:00:00Z"><w:p><w:r><w:t>请核对这句话</w:t></w:r></w:p></w:comment></w:comments>');
    zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}<w:sectPr><w:pgSz w:w="11910" w:h="16840"/><w:pgMar w:top="1580" w:bottom="280" w:left="880" w:right="680"/></w:sectPr></w:body></w:document>`);
    const doc = await parseDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(doc.comments).toHaveLength(1);
    expect(doc.comments![0]).toMatchObject({ id: '1', author: '张三', text: '请核对这句话' });
    const html = renderParagraphHtml(doc.blocks[0] as never);
    expect(html).toContain('data-comment-ids="1"');
    expect(html).toContain('background-color:#fff3bf');

    // 写回：comments.xml + commentRangeStart/End/Reference
    const saved = await serializeDocx(doc, { name: 'c.docx' });
    const savedZip = await JSZip.loadAsync(saved);
    const savedComments = await savedZip.file('word/comments.xml')!.async('text');
    expect(savedComments).toContain('w:author="张三"');
    const savedDoc = await savedZip.file('word/document.xml')!.async('text');
    expect(savedDoc).toContain('<w:commentRangeStart w:id="1"/>');
    expect(savedDoc).toContain('<w:commentReference w:id="1"/>');
    const reopened = await parseDocx(saved);
    expect(reopened.comments![0].text).toBe('请核对这句话');
    const reopenedHtml = renderParagraphHtml(reopened.blocks[0] as never);
    expect(reopenedHtml).toContain('data-comment-ids="1"');
  });

  it('修订：w:ins/w:del 保留（不再静默接受），渲染配色，接受/拒绝操作与写回', async () => {
    const body =
      '<w:p><w:ins w:id="1" w:author="李四" w:date="2026-10-31T00:00:00Z"><w:r><w:t>新增文字</w:t></w:r></w:ins>' +
      '<w:del w:id="2" w:author="李四" w:date="2026-10-31T00:00:00Z"><w:r><w:delText>删除文字</w:delText></w:r></w:del>' +
      '<w:r><w:t>正文</w:t></w:r></w:p>';
    const doc = await parseDocx(await makeDocx(body, SECT));
    const runs = doc.blocks[0].type === 'paragraph' ? doc.blocks[0].runs : [];
    expect(runs).toHaveLength(3);
    expect(runs[0]).toMatchObject({ text: '新增文字', revision: { type: 'insert', author: '李四' } });
    expect(runs[1]).toMatchObject({ text: '删除文字', revision: { type: 'delete' } });
    expect(doc.hasTrackedChanges).toBe(true);

    // 渲染：插入绿底、删除红底删除线
    const html = renderParagraphHtml(doc.blocks[0] as never);
    expect(html).toContain('data-revision="insert"');
    expect(html).toContain('data-revision="delete"');
    expect(html).toContain('text-decoration:line-through');

    // 接受插入 → 去标记；拒绝插入 → 删除 run；接受删除 → 删 run；拒绝删除 → 恢复
    const acceptInsert = applyRevisionDecision(doc.blocks, doc.blocks[0].type === 'paragraph' ? doc.blocks[0].id : '', 0, true);
    expect((acceptInsert[0] as any).runs[0].revision).toBeUndefined();
    const rejectInsert = applyRevisionDecision(doc.blocks, (doc.blocks[0] as any).id, 0, false);
    expect((rejectInsert[0] as any).runs.some((r: any) => r.text === '新增文字')).toBe(false);
    const acceptDelete = applyRevisionDecision(doc.blocks, (doc.blocks[0] as any).id, 1, true);
    expect((acceptDelete[0] as any).runs.some((r: any) => r.text === '删除文字')).toBe(false);
    const rejectDelete = applyRevisionDecision(doc.blocks, (doc.blocks[0] as any).id, 1, false);
    expect((rejectDelete[0] as any).runs[1].revision).toBeUndefined();

    // 写回：w:ins/w:del + w:delText
    const saved = await serializeDocx(doc, { name: 'r.docx' });
    const savedZip = await JSZip.loadAsync(saved);
    const savedDoc = await savedZip.file('word/document.xml')!.async('text');
    expect(savedDoc).toContain('<w:ins w:author="李四"');
    expect(savedDoc).toContain('<w:del w:author="李四"');
    expect(savedDoc).toContain('<w:delText');
    const reopened = await parseDocx(saved);
    const reopenedRuns = reopened.blocks[0].type === 'paragraph' ? reopened.blocks[0].runs : [];
    expect(reopenedRuns[0].revision?.type).toBe('insert');
    expect(reopenedRuns[1].revision?.type).toBe('delete');
  });
});

