/** 功能：验证原始 XML 保存、未知对象保留与连续保存，避免只测试重新解析后的可见文字；时间：2026-11-03；作者：lq。 */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { parseDocx } from '../packages/editor-word/src/parse';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import type { DocParagraph } from '@mini-office/document-model';

const first = '<w:p w:rsidR="ABCDEF12"><w:pPr><w:keepNext w:val="0"/></w:pPr><w:bookmarkStart w:id="42" w:name="kept"/><w:r><w:rPr><w:lang w:val="zh-CN"/></w:rPr><w:t xml:space="preserve">原始第一段</w:t></w:r><w:bookmarkEnd w:id="42"/></w:p>';
const second = '<w:p w:rsidR="00000002"><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Second paragraph</w:t></w:r></w:p>';
const opaque = '<w:customXml w:element="unknown"><w:customXmlPr/></w:customXml>';

/** 功能：用仓库样本的有效包结构构造带未知属性和书签的正文；时间：2026-11-03；作者：lq。 */
async function fixture(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(readFileSync('tests/fixtures/docx/contract.docx'));
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${first}${opaque}${second}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

/** 功能：直接检查保存后的 XML，有效负载不通过模型二次解释；时间：2026-11-03；作者：lq。 */
async function xmlOf(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('text');
}

describe('原始 XML 局部保存', () => {
  it('没有修改时返回完全相同的原文件字节', async () => {
    const bytes = await fixture();
    const doc = await parseDocx(bytes);
    expect(doc.source).toBeDefined();
    expect(await serializeDocx(doc, { sourceBytes: bytes, name: 'test.docx' })).toEqual(bytes);
  });

  it('文字修改保留书签、语言、未知元素、其他段落和原样式部件', async () => {
    const bytes = await fixture();
    const doc = await parseDocx(bytes);
    (doc.blocks[0] as DocParagraph).runs[0].text = '修改第一段';
    const saved = await serializeDocx(doc, { sourceBytes: bytes, name: 'test.docx' });
    expect(await xmlOf(saved)).toBe((await xmlOf(bytes)).replace('原始第一段', '修改第一段'));
    const zip = await JSZip.loadAsync(saved);
    expect(await zip.file('word/styles.xml')!.async('text')).toBe(await (await JSZip.loadAsync(bytes)).file('word/styles.xml')!.async('text'));
  });

  it('连续保存后仍使用原始锚点，撤销能恢复原 XML', async () => {
    const bytes = await fixture();
    const doc = await parseDocx(bytes);
    const p = doc.blocks[0] as DocParagraph;
    p.runs[0].text = '第一次';
    const once = await serializeDocx(doc, { sourceBytes: bytes, name: 'test.docx' });
    p.runs[0].text = '第二次';
    const twice = await serializeDocx(doc, { sourceBytes: once, name: 'test.docx' });
    expect(await xmlOf(twice)).toBe((await xmlOf(bytes)).replace('原始第一段', '第二次'));
    p.runs[0].text = '原始第一段';
    expect(await xmlOf(await serializeDocx(doc, { sourceBytes: twice, name: 'test.docx' }))).toBe(await xmlOf(bytes));
  });

  it('插入和删除段落后仍保留未改的 XML 和未知顶层元素', async () => {
    const bytes = await fixture();
    const doc = await parseDocx(bytes);
    doc.blocks.unshift({ type: 'paragraph', id: 'new', runs: [{ text: '新段', style: {} }], style: {} });
    doc.blocks.splice(1, 1);
    const saved = await serializeDocx(doc, { sourceBytes: bytes, name: 'test.docx' });
    const xml = await xmlOf(saved);
    expect(xml).toContain(second);
    expect(xml).toContain(opaque);
    expect(xml).not.toContain(first);
    expect(xml).toContain('新段');
  });
});
