/** 功能：验证非英文 ID 表格样式的边框、边距继承和直接格式覆盖；时间：2026-11-03；作者：lq。 */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { parseDocx } from '../packages/editor-word/src/parse';

/** 功能：使用中文 Word 常见的 a1/a4 样式 ID 构造表格，防止依靠 Grid 名称猜测边框；时间：2026-11-03；作者：lq。 */
async function tableDoc(direct = '') {
  const zip = await JSZip.loadAsync(readFileSync('tests/fixtures/docx/contract.docx'));
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  zip.file('word/styles.xml', `<w:styles xmlns:w="${ns}"><w:style w:type="table" w:styleId="a1" w:default="1"><w:tblPr><w:tblCellMar><w:left w:w="300" w:type="dxa"/><w:right w:w="450" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="a4"><w:basedOn w:val="a1"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr></w:style></w:styles>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="a4"/>${direct}</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcMar><w:top w:w="90" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:t>内容</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`);
  const doc = await parseDocx(await zip.generateAsync({ type: 'uint8array' }));
  const table = doc.blocks.find((block) => block.type === 'table')!;
  return table;
}

describe('表格样式继承', () => {
  it('a4 继承 a1 边距并恢复边框，单元格逐边覆盖不抹掉左右边距', async () => {
    const table = await tableDoc();
    expect(table.borders?.top?.style).toBe('single');
    expect(table.borders?.insideH?.style).toBe('single');
    expect(table.borders?.insideV?.style).toBe('single');
    expect(table.rows[0].cells[0].marginPx).toEqual({ top: 6, right: 30, bottom: 0, left: 20 });
  });
  it('显式 nil 边框覆盖样式，不被自动恢复', async () => {
    const table = await tableDoc('<w:tblBorders><w:top w:val="nil"/></w:tblBorders>');
    expect(table.borders?.top?.style).toBe('none');
    expect(table.borders?.insideH?.style).toBe('single');
  });
});
