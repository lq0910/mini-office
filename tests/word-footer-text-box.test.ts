/** 功能：回归 DrawingML 页脚文本框的域解析、锚定与编辑其他页脚文字后的保留；时间：2026-11-05；作者：lq。 */
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseDocx } from '../packages/editor-word/src/parse';
import { serializeDocx } from '../packages/editor-word/src/serialize';

/** 功能：构造只包含页码域的居中文本框和独立页脚关系，不依赖私人文件；时间：2026-11-05；作者：lq。 */
async function textBoxFixture(): Promise<Uint8Array> {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>');
  zip.file('word/document.xml', `<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body><w:p><w:r><w:t>正文</w:t></w:r></w:p><w:sectPr><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFooter" Type="${r}/footer" Target="footer1.xml"/></Relationships>`);
  zip.file('word/footer1.xml', `<w:ftr xmlns:w="${w}" xmlns:wp="${wp}" xmlns:a="${a}" xmlns:wps="${wps}"><w:p><w:r><w:drawing><wp:anchor><wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="1828800"/><wp:wrapNone/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent><w:p><w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p></w:txbxContent></wps:txbx><wps:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p></w:ftr>`);
  return zip.generateAsync({ type: 'uint8array' });
}

describe('页脚页码文本框', () => {
  it('没有图片也保留文本框和 PAGE 域，不把锚定文字重复展开到普通段落', async () => {
    const bytes = await textBoxFixture();
    const doc = await parseDocx(bytes);
    expect(doc.footer?.pageNumbers).toBe(true);
    expect(doc.footer?.textBoxes).toHaveLength(1);
    const box = doc.footer!.textBoxes![0];
    expect(box).toMatchObject({ widthPx: 192, heightPx: 192, autoFit: true, anchor: { hAlign: 'center', hRelative: 'margin' } });
    expect(box.blocks[0]).toMatchObject({ type: 'paragraph', runs: [{ field: 'PAGE', text: '1' }] });
    expect(doc.footer!.blocks[0]).toMatchObject({ type: 'paragraph', runs: [{ text: '' }] });
    expect(await serializeDocx(doc, { name: 'footer.docx', sourceBytes: bytes })).toEqual(bytes);
  });

  it('编辑普通页脚文字后保存回读，文本框和居中 PAGE 域仍存在且不重复', async () => {
    const bytes = await textBoxFixture();
    const doc = await parseDocx(bytes);
    const footer = doc.footer!;
    const paragraph = footer.blocks[0];
    if (paragraph.type !== 'paragraph') throw new Error('缺少页脚段落');
    paragraph.runs = [{ text: '页脚说明', style: {} }];
    footer.edited = true;
    const saved = await serializeDocx(doc, { name: 'footer.docx', sourceBytes: bytes });
    const reopened = await parseDocx(saved);
    expect(reopened.footer?.textBoxes).toHaveLength(1);
    expect(reopened.footer?.textBoxes?.[0].anchor?.hAlign).toBe('center');
    expect(reopened.footer?.pageNumbers).toBe(true);
    expect(reopened.footer?.blocks[0]).toMatchObject({ runs: [{ text: '页脚说明' }] });
  });
});
