/** 功能：通过环境变量显式运行用户语料验证，不把私人文档提交到仓库；时间：2026-11-03；作者：lq。 */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { convertLegacyDoc } from '../packages/platform-adapter/src/legacy-doc';
import { parseDocx } from '../packages/editor-word/src/parse';
import { serializeDocx } from '../packages/editor-word/src/serialize';

const directory = process.env.MINI_OFFICE_FIDELITY_SAMPLES;

describe.skipIf(!directory)('用户三个样本：真实结构与无操作往返', () => {
  it.each(['采购文件.doc', '测试分包合同.doc', '测试请假单.docx'])('%s 的导入、保存和原件完整性', async (name) => {
    const path = join(directory!, name);
    const input = new Uint8Array(readFileSync(path));
    const hash = createHash('sha256').update(input).digest('hex');
    const result = name.endsWith('.doc') ? await convertLegacyDoc(input, name) : { name, bytes: input, conversionWarnings: [] };
    const zip = await JSZip.loadAsync(result.bytes);
    const xml = await zip.file('word/document.xml')!.async('text');
    const dom = new DOMParser().parseFromString(xml, 'application/xml');
    const count = (tag: string) => dom.getElementsByTagNameNS('*', tag).length;
    expect(count('rFonts')).toBeGreaterThan(0);
    expect(count('pgSz')).toBeGreaterThan(0);
    expect(count('pgMar')).toBeGreaterThan(0);
    const doc = await parseDocx(result.bytes, { fileName: result.name });
    // 功能：真实旧模板关闭孤行控制，转换与解析均须保留；否则合同页尾文字会被整段提前移走；时间：2026-11-05；作者：lq。
    if (name === '测试分包合同.doc') {
      expect(xml).toContain('<w:widowControl w:val="0"');
      const body = doc.blocks.find((block) => block.type === 'paragraph' && block.runs.some((run) => run.text?.includes('依照')));
      expect(body?.type === 'paragraph' && body.style.widowControl).toBe(false);
      expect(zip.file('word/fontTable.xml')).not.toBeNull();
    }
    // 功能：采购模板首章编号保留中文计数，并继承其关闭网格的原始设置；时间：2026-11-05；作者：lq。
    if (name === '采购文件.doc') {
      const heading = doc.blocks.find((block) => block.type === 'paragraph' && block.list?.marker.startsWith('第'));
      expect(heading?.type === 'paragraph' && heading.list?.marker).toBe('第一章');
      expect(heading?.type === 'paragraph' && heading.style.snapToGrid).toBe(false);
      // 功能：真实采购页码在页脚文本框中，保留字段和源居中锚定；时间：2026-11-05；作者：lq。
      expect(doc.footer?.pageNumbers).toBe(true);
      expect(doc.footer?.textBoxes?.[0].anchor?.hAlign).toBe('center');
      const page = doc.footer?.textBoxes?.[0].blocks[0];
      expect(page?.type === 'paragraph' && page.runs.some((run) => run.field === 'PAGE')).toBe(true);
      // 功能：采购首表按原件保留 99.96% 首选宽度与 autofit 模式；时间：2026-11-05；作者：lq。
      const table = doc.blocks.find((block) => block.type === 'table');
      expect(table?.type === 'table' && table.widthPct).toBe(99.96);
      expect(table?.type === 'table' && table.layout).toBe('autofit');
    }
    expect(doc.source).toBeDefined();
    expect(await serializeDocx(doc, { sourceBytes: result.bytes, name: result.name })).toEqual(result.bytes);
    if (name.endsWith('.docx')) {
      const table = doc.blocks.find((block) => block.type === 'table');
      expect(table?.columnWidthsPx.map((width) => Math.round(width * 15))).toEqual([2755, 1596, 1276, 3421]);
      expect(table?.borders?.top?.style).toBe('single');
      expect(table?.borders?.insideH?.style).toBe('single');
    }
    // 功能：真实样本上修改一个普通文字 run，确认其他包部件逐字保持且编辑结果可重新打开；时间：2026-11-03；作者：lq。
    const paragraph = doc.blocks.find((block) => block.type === 'paragraph' && block.runs.some((run) => run.text?.trim()));
    if (!paragraph || paragraph.type !== 'paragraph') throw new Error('样本缺少可验证的普通段落');
    const run = paragraph.runs.find((item) => item.text?.trim())!;
    run.text += ' [Mini Office 验证]';
    const edited = await serializeDocx(doc, { sourceBytes: result.bytes, name: result.name });
    const editedZip = await JSZip.loadAsync(edited);
    for (const [part, entry] of Object.entries(zip.files)) {
      if (entry.dir || part === 'word/document.xml') continue;
      expect(await editedZip.file(part)!.async('uint8array'), `${name}: ${part}`).toEqual(await entry.async('uint8array'));
    }
    const reopened = await parseDocx(edited, { fileName: result.name });
    expect(reopened.blocks.some((block) => block.type === 'paragraph' && block.runs.some((item) => item.text?.includes('[Mini Office 验证]')))).toBe(true);
    expect(createHash('sha256').update(readFileSync(path)).digest('hex')).toBe(hash);
    console.log(JSON.stringify({ name, tables: count('tbl'), fontRuns: count('rFonts'), sections: count('sectPr'), noOpByteExact: true, warnings: result.conversionWarnings }));
  });
});
