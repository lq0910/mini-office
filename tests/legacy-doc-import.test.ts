/** 功能：验证旧 DOC 导入保留字体、表格和页面结构，且不改变输入字节；时间：2026-11-03；作者：lq。 */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { convertLegacyDoc } from '../packages/platform-adapter/src/legacy-doc';

describe('GenOffice DOC 转换接入', () => {
  /** 功能：扩展名为 DOC 而实际是 DOCX 时保留原字节，不误报已转换或格式损失；时间：2026-11-11；作者：lq。 */
  it('改名的 DOCX 直接打开，非 Word ZIP 拒绝冒充 DOC', async () => {
    const bytes = new Uint8Array(readFileSync('tests/fixtures/docx/contract.docx'));
    const result = await convertLegacyDoc(bytes, '改名合同.doc');
    expect(result).toMatchObject({ name: '改名合同.docx', sourceName: '改名合同.doc', conversionPerformed: false, conversionWarnings: [] });
    expect(result.bytes).toBe(bytes);
    const zip = new JSZip(); zip.file('xl/workbook.xml', '<workbook/>');
    await expect(convertLegacyDoc(await zip.generateAsync({ type: 'uint8array' }), '表格.doc')).rejects.toThrow('不是 Word');
  });
  it.each(['formatting', 'table', 'page-setup'])('公开样本 %s 转换为 OOXML 并保留输入', async (name) => {
    const bytes = new Uint8Array(readFileSync(`third_party/genoffice/fixtures/${name}.doc`));
    const original = bytes.slice();
    const result = await convertLegacyDoc(bytes, `/temporary/${name}.doc`);
    expect(result.name).toBe(`${name}.docx`);
    expect(result.sourceName).toBe(`${name}.doc`);
    expect(bytes).toEqual(original);
    const xml = await (await JSZip.loadAsync(result.bytes)).file('word/document.xml')!.async('text');
    expect(xml).toContain('<w:pgSz');
    expect(xml).toContain('<w:pgMar');
    if (name === 'formatting') expect(xml).toContain('<w:rFonts');
    if (name === 'table') expect(xml).toContain('<w:tbl>');
  });

  it('非法输入报错，不静默降级为文本', async () => {
    await expect(convertLegacyDoc(new TextEncoder().encode('not a DOC'), 'invalid.doc')).rejects.toThrow();
  });
});
