/** 功能：验证 DOC FFN 备用名、DOCX 字体表解析及局部编辑保存不替换主字体；时间：2026-11-05；作者：lq。 */
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseFontTable, type FontEntry } from '../third_party/genoffice/doc2docx/fonts';
import type { FibInfo } from '../third_party/genoffice/doc2docx/fib';
import { parseDocx } from '../packages/editor-word/src/parse';
import { renderRunsHtml } from '../packages/editor-word/src/html';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import { fontStackFor } from '@mini-office/design-system';
// 功能：回归 DOC 孤行控制的 Bool8 开关，不依赖私人模板；时间：2026-11-05；作者：lq。
import { applyParagraphGrpprl, type ParagraphFormatting } from '../third_party/genoffice/doc2docx/formatting';

/** 功能：构造独立的 Unicode FFN 主名和备用名，不依赖私人 DOC 原件；时间：2026-11-05；作者：lq。 */
function fontRecord(name: string, alternate: string): Uint8Array {
  const chars = `${name}\0${alternate}\0`;
  const record = new Uint8Array(40 + chars.length * 2);
  record[0] = record.length - 1;
  record[5] = name.length + 1;
  const view = new DataView(record.buffer);
  for (let index = 0; index < chars.length; index++) view.setUint16(40 + index * 2, chars.charCodeAt(index), true);
  return record;
}

describe('文件声明的备用字体', () => {
  // 功能：DOC 样式中的关闭值必须经过解析保留，后续直接格式仍可显式开启；时间：2026-11-05；作者：lq。
  it('读取 DOC 段落的显式孤行控制关闭和开启值', () => {
    const format: ParagraphFormatting = {};
    const context = { source: {}, warnings: [] } as unknown as Parameters<typeof applyParagraphGrpprl>[2];
    applyParagraphGrpprl(format, new Uint8Array([0x31, 0x24, 0]), context, 'inherit');
    expect(format.widowControl).toBe(false);
    applyParagraphGrpprl(format, new Uint8Array([0x31, 0x24, 1]), context);
    expect(format.widowControl).toBe(true);
    // 功能：网格关闭与开启也应在样式继承和直接格式间完整传递；时间：2026-11-05；作者：lq。
    applyParagraphGrpprl(format, new Uint8Array([0x47, 0x24, 0]), context, 'inherit');
    expect(format.snapToGrid).toBe(false);
    applyParagraphGrpprl(format, new Uint8Array([0x47, 0x24, 1]), context);
    expect(format.snapToGrid).toBe(true);
  });
  it('读取 FFN 备用名，备用名索引越界时不读取下一条记录', () => {
    const first = fontRecord('旧字体', '备用字体');
    const second = fontRecord('另一字体', '另一备用');
    const table = new Uint8Array(4 + first.length + second.length);
    table[0] = 2; table.set(first, 4); table.set(second, 4 + first.length);
    const fib = { fcSttbfFfn: 0, lcbSttbfFfn: table.length } as FibInfo;
    const entries: FontEntry[] = [];
    expect(parseFontTable(fib, table, [], entries)).toEqual(['旧字体', '另一字体']);
    expect(entries).toEqual([{ name: '旧字体', alternateName: '备用字体' }, { name: '另一字体', alternateName: '另一备用' }]);
    table[4 + 5] = 255;
    const invalidEntries: FontEntry[] = [];
    parseFontTable(fib, table, [], invalidEntries);
    expect(invalidEntries).toEqual([{ name: '旧字体' }, { name: '另一字体', alternateName: '另一备用' }]);
  });

  it('备用名只作用于本文件，且实际主字体仍有最高优先级', () => {
    expect(fontStackFor('仿宋_GB2312', '微软雅黑')).toMatch(/^"仿宋_GB2312", "微软雅黑", "Microsoft YaHei"/);
    expect(fontStackFor('仿宋_GB2312')).not.toContain('Microsoft YaHei');
    expect(fontStackFor('Arial', 'Arial')).toBe(fontStackFor('Arial'));
  });

  it('解析和显示遵守备用名，修改文字后保存仍保留主名和原字体表', async () => {
    const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const zip = new JSZip();
    const fonts = `<w:fonts xmlns:w="${ns}"><w:font w:name="旧字体"><w:altName w:val="微软雅黑"/></w:font></w:fonts>`;
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
    zip.file('word/fontTable.xml', fonts);
    zip.file('word/document.xml', `<w:document xmlns:w="${ns}"><w:body><w:p><w:r><w:rPr><w:rFonts w:eastAsia="旧字体"/><w:sz w:val="24"/></w:rPr><w:t>原有文字</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
    const input = await zip.generateAsync({ type: 'uint8array' });
    const doc = await parseDocx(input);
    const paragraph = doc.blocks[0];
    if (paragraph.type !== 'paragraph') throw new Error('未解析出段落');
    expect(paragraph.runs[0].style).toMatchObject({ fontFamily: '旧字体', fontFallback: '微软雅黑' });
    expect(renderRunsHtml(paragraph.runs)).toContain("'旧字体', '微软雅黑', 'Microsoft YaHei'");
    expect(await serializeDocx(doc, { sourceBytes: input, name: 'font-test.docx' })).toEqual(input);
    paragraph.runs[0].text += '修改';
    const saved = await JSZip.loadAsync(await serializeDocx(doc, { sourceBytes: input, name: 'font-test.docx' }));
    expect(await saved.file('word/fontTable.xml')!.async('text')).toBe(fonts);
    const reopened = await parseDocx(await saved.generateAsync({ type: 'uint8array' }));
    expect(reopened.blocks[0]).toMatchObject({ runs: [{ text: '原有文字修改', style: { fontFamily: '旧字体', fontFallback: '微软雅黑' } }] });
    // 功能：用户切换字体后，绑定旧字体的显示备用名不能污染新字体；时间：2026-11-05；作者：lq。
    paragraph.runs[0].style.fontFamily = 'Arial';
    expect(renderRunsHtml(paragraph.runs)).not.toContain('Microsoft YaHei');
  });

  // 功能：只有安装清单确认缺失旧 GB 字体时才走兼容替代，原字体存在或环境未知时均保留；时间：2026-11-05；作者：lq。
  it('缺失 GB 字体才采用 Word 兼容备用字体', async () => {
    const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
    zip.file('word/document.xml', `<w:document xmlns:w="${ns}"><w:body><w:p><w:r><w:rPr><w:rFonts w:eastAsia="仿宋_GB2312"/></w:rPr><w:t>中文</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const style = async (knownFonts?: string[]) => {
      const doc = await parseDocx(bytes, { knownFonts });
      const paragraph = doc.blocks[0];
      if (paragraph.type !== 'paragraph') throw new Error('未解析出段落');
      return paragraph.runs[0].style;
    };
    expect(await style([])).toMatchObject({ fontFamily: '仿宋_GB2312', fontFallback: '微软雅黑', fontFallbackFor: '仿宋_GB2312' });
    expect((await style(['仿宋_GB2312'])).fontFallback).toBeUndefined();
    expect((await style()).fontFallback).toBeUndefined();
  });
});
