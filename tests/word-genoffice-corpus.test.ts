/** 功能：固定公开语料的来源、结构及保存往返，避免仅检查 XML 标签造成虚假兼容；时间：2026-11-03；作者：lq。 */
import { beforeAll, describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import manifest from '../third_party/genoffice/fixtures/manifest.json';
import { convertLegacyDoc } from '../packages/platform-adapter/src/legacy-doc';
import { parseDocx } from '../packages/editor-word/src/parse';
import { serializeDocx } from '../packages/editor-word/src/serialize';
// 功能：校验纯图片语料的效果外扩在渲染与重建保存时保持；时间：2026-11-04；作者：lq。
import { renderRunsHtml } from '../packages/editor-word/src/html';
import type { DocBlock, DocParagraph, DocModel } from '@mini-office/document-model';

const directory = resolve('third_party/genoffice/fixtures');
const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');

/** 功能：递归遍历真实表格单元格中的段落，保持正文顺序；时间：2026-11-03；作者：lq。 */
function paragraphs(blocks: DocBlock[]): DocParagraph[] {
  return blocks.flatMap((block): DocParagraph[] => block.type === 'paragraph' ? [block]
    : block.type === 'table' ? block.rows.flatMap((row) => row.cells.flatMap((cell) => paragraphs(cell.blocks))) : []);
}
const text = (paragraph: DocParagraph) => paragraph.runs.map((run) => run.tab ? '\t' : run.text ?? '').join('');

/** 功能：按纯文本导出的行列规则展平表格，不混淆单元格分隔与正文 Tab；时间：2026-11-03；作者：lq。 */
function exportedText(blocks: DocBlock[]): string {
  return blocks.map((block) => block.type === 'paragraph' ? text(block) : block.type === 'table'
    ? block.rows.map((row) => row.cells.map((cell) => exportedText(cell.blocks)).join('\t')).join('\n') : '').join('\n');
}

// 功能：jsdom 没有浏览器对象 URL 接口，显式补齐以验证图片解析而非环境缺失；时间：2026-11-03；作者：lq。
beforeAll(() => {
  if (!URL.createObjectURL) URL.createObjectURL = () => 'blob:genoffice-corpus';
  if (!URL.revokeObjectURL) URL.revokeObjectURL = () => {};
});

/** 功能：通过产品解析器读取相同固定语料，旧 DOC 先走正式转换通道；时间：2026-11-03；作者：lq。 */
async function open(id: string): Promise<{ doc: DocModel; bytes: Uint8Array; input: Uint8Array; name: string }> {
  const fixture = manifest.fixtures.find((item) => item.id === id)!;
  const input = new Uint8Array(readFileSync(resolve(directory, fixture.file)));
  const result = fixture.format === 'doc' ? await convertLegacyDoc(input, fixture.file) : { bytes: input, name: `${id}.docx` };
  return { input, bytes: result.bytes, name: result.name, doc: await parseDocx(result.bytes, { fileName: result.name }) };
}

describe('GenOffice 固定公开语料', () => {
  // 功能：等长替换公开 DOC 内存副本的表格边框属性，验证宽度单位、布局开关及段落操作码隔离；时间：2026-11-05；作者：lq。
  it.each([
    { unit: 2, value: 4998, type: 'pct', autofit: true },
    { unit: 3, value: 8640, type: 'dxa', autofit: false },
    { unit: 1, value: 0, type: 'auto', autofit: true },
  ])('DOC 表格首选 $type 宽度和 autofit=$autofit 保留至转换包', async ({ unit, value, type, autofit }) => {
    const input = Buffer.from(readFileSync(resolve(directory, 'table.doc')));
    const signature = Buffer.from([0x12, 0xd6, 30, 0, 0, 0, 0xff]);
    const replacement = Buffer.from([
      0x14, 0xf6, unit, value & 255, value >> 8,
      0x15, 0x36, Number(autofit),
      0x14, 0xa4, 120, 0,
      0x15, 0xc6, 3, 0, 0, 0,
      ...Array.from({ length: 5 }, () => [0x03, 0x24, 0]).flat(),
    ]);
    expect(replacement.length).toBe(33);
    let offset = 0;
    let rows = 0;
    while ((offset = input.indexOf(signature, offset)) >= 0) {
      replacement.copy(input, offset);
      offset += replacement.length;
      rows++;
    }
    expect(rows).toBe(2);
    const converted = await convertLegacyDoc(new Uint8Array(input), 'width.doc');
    const xml = await (await JSZip.loadAsync(converted.bytes)).file('word/document.xml')!.async('text');
    expect(xml).toContain(`<w:tblW w:w="${value}" w:type="${type}"/>`);
    expect(xml).toContain(`<w:tblLayout w:type="${autofit ? 'autofit' : 'fixed'}"/>`);
    const doc = await parseDocx(converted.bytes);
    const table = doc.blocks.find((block) => block.type === 'table');
    expect(table?.layout).toBe(autofit ? 'autofit' : 'fixed');
    expect(table?.widthPct).toBe(type === 'pct' ? value / 50 : undefined);
    expect(table?.columnWidthsPx).toEqual([192, 192, 192]);
  });

  // 功能：平移公开 DOC 两行的全部原始边界，验证绝对缩进被保留且不改变列宽；时间：2026-11-05；作者：lq。
  it('DOC 原始表格边界换算为首格文字缩进，保存后保持', async () => {
    const input = Buffer.from(readFileSync(resolve(directory, 'table.doc')));
    const signature = Buffer.from([0x08, 0xd6, 0x46, 0x00, 0x03, 0x94, 0xff, 0xd4, 0x0a]);
    let offset = 0;
    let rows = 0;
    while ((offset = input.indexOf(signature, offset)) >= 0) {
      for (let column = 0; column < 4; column++) {
        const position = offset + 5 + column * 2;
        input.writeInt16LE(input.readInt16LE(position) + 360, position);
      }
      offset += signature.length;
      rows++;
    }
    expect(rows).toBe(2);
    const converted = await convertLegacyDoc(new Uint8Array(input), 'indent.doc');
    const xml = await (await JSZip.loadAsync(converted.bytes)).file('word/document.xml')!.async('text');
    expect(xml).toContain('<w:tblInd w:w="360" w:type="dxa"/>');
    const doc = await parseDocx(converted.bytes);
    const table = doc.blocks.find((block) => block.type === 'table');
    expect(table?.indentPx).toBe(24);
    expect(table?.columnWidthsPx).toEqual([192, 192, 192]);
    if (!table) throw new Error('公开 DOC 未重建出表格');
    table.rows[0].cells[0].backgroundColor = '#123456';
    const saved = await serializeDocx(doc, { sourceBytes: converted.bytes, name: 'indent.docx' });
    expect((await parseDocx(saved)).blocks.find((block) => block.type === 'table')?.indentPx).toBe(24);
  });

  // 功能：只修改公开 DOC 内存副本的对齐位并保留其他 TC80 位，验证二进制解码、转换及产品解析的整条路径；时间：2026-11-05；作者：lq。
  it('DOC TC80 的上中下对齐经过转换和 DOCX 保存回读', async () => {
    const input = Buffer.from(readFileSync(resolve(directory, 'table.doc')));
    const signature = Buffer.from([0x08, 0xd6, 0x46, 0x00, 0x03, 0x94, 0xff, 0xd4, 0x0a]);
    let offset = 0;
    let rows = 0;
    while ((offset = input.indexOf(signature, offset)) >= 0) {
      const tcStart = offset + 4 + 1 + 2 * (3 + 1);
      for (let column = 0; column < 3; column++) {
        const position = tcStart + column * 20;
        input.writeUInt16LE((input.readUInt16LE(position) & ~0x0180) | (column << 7), position);
      }
      offset += signature.length;
      rows++;
    }
    expect(rows).toBe(2);
    const converted = await convertLegacyDoc(new Uint8Array(input), 'alignment.doc');
    const doc = await parseDocx(converted.bytes);
    const table = doc.blocks.find((block) => block.type === 'table');
    if (!table || table.type !== 'table') throw new Error('公开 DOC 未重建出表格');
    expect(table.rows.map((row) => row.cells.map((cell) => cell.verticalAlign)))
      .toEqual([['top', 'middle', 'bottom'], ['top', 'middle', 'bottom']]);
    // 功能：公开 DOC 明确声明上下 0、左右 108 twips；不仅验证显示默认值，还检查转换包实际写入；时间：2026-11-05；作者：lq。
    const convertedXml = await (await JSZip.loadAsync(converted.bytes)).file('word/document.xml')!.async('text');
    expect(convertedXml).toContain('<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>');
    expect(table.cellMarginPx).toEqual({ top: 0, left: 7.2, bottom: 0, right: 7.2 });
    table.rows[0].cells[0].backgroundColor = '#123456';
    const saved = await serializeDocx(doc, { sourceBytes: converted.bytes, name: 'alignment.docx' });
    const reopened = (await parseDocx(saved)).blocks.find((block) => block.type === 'table');
    expect(reopened?.type === 'table' && reopened.rows[0].cells.map((cell) => cell.verticalAlign))
      .toEqual(['top', 'middle', 'bottom']);
    expect(reopened?.type === 'table' && reopened.cellMarginPx).toEqual({ top: 0, left: 7.2, bottom: 0, right: 7.2 });
  });

  it('语料 ID 唯一，生成器有独立哈希，Office 基准不冒充已验收', () => {
    expect(manifest.commit).toBe('e1f50e90d246f3ecc6427d2a0af90f04f37810ec');
    expect(new Set(manifest.fixtures.map((item) => item.id)).size).toBe(manifest.fixtures.length);
    expect(hash(readFileSync(resolve(directory, manifest.generatorHelper.file)))).toBe(manifest.generatorHelper.sha256);
    for (const fixture of manifest.fixtures) {
      expect(fixture.upstreamPath).toBeTruthy();
      expect(fixture.coverage.length).toBeGreaterThan(0);
      expect(fixture.officeBaseline).toBe('pending');
      if ('generator' in fixture.provenance) {
        expect(hash(readFileSync(resolve(directory, fixture.provenance.generator!)))).toBe(fixture.provenance.generatorSha256);
      }
      if ('expectedText' in fixture.provenance) {
        const expected = fixture.provenance.expectedText!;
        expect(hash(readFileSync(resolve(directory, expected.file)))).toBe(expected.sha256);
      }
    }
  });

  it.each(manifest.fixtures)('$id：哈希、无操作字节一致、小改动与未编辑部件保留', async (fixture) => {
    const { input, bytes, name, doc } = await open(fixture.id);
    expect(hash(input)).toBe(fixture.sha256);
    expect(doc.source).toBeDefined();
    expect(await serializeDocx(doc, { sourceBytes: bytes, name })).toEqual(bytes);
    const paragraph = paragraphs(doc.blocks).find((item) => item.runs.some((run) => run.text?.trim()));
    if (!paragraph) throw new Error(`${fixture.id}: 缺少可编辑文字`);
    paragraph.runs.find((run) => run.text?.trim())!.text += ' [lq corpus]';
    const saved = await serializeDocx(doc, { sourceBytes: bytes, name });
    const [original, edited] = await Promise.all([JSZip.loadAsync(bytes), JSZip.loadAsync(saved)]);
    const parts = (zip: JSZip) => Object.values(zip.files).filter((entry) => !entry.dir).map((entry) => entry.name).sort();
    expect(parts(edited)).toEqual(parts(original));
    for (const [part, entry] of Object.entries(original.files)) {
      if (entry.dir || part === 'word/document.xml') continue;
      expect(await edited.file(part)!.async('uint8array'), `${fixture.id}: ${part}`).toEqual(await entry.async('uint8array'));
    }
    const reopened = await parseDocx(saved, { fileName: name });
    expect(paragraphs(reopened.blocks).some((item) => text(item).includes('[lq corpus]'))).toBe(true);
    expect(hash(readFileSync(resolve(directory, fixture.file)))).toBe(fixture.sha256);
  });

  it.each(['formatting', 'table', 'page-setup', 'text-basic'])('%s：DOC 正文对齐上游纯文本导出', async (id) => {
    const { doc } = await open(id);
    const expected = readFileSync(resolve(directory, `${id}.txt`), 'utf8').replace(/^\uFEFF/, '').replace(/\r\n/g, '\n')
      .replace(/\\u([0-9a-f]{4})/gi, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)));
    expect(exportedText(doc.blocks).trimEnd()).toBe(expected.trimEnd());
  });

  it('保留段内换行自身的 11pt/28pt 字号及相邻 14pt 文本', async () => {
    const { doc } = await open('br-only-line-size');
    const items = paragraphs(doc.blocks);
    expect(items).toHaveLength(2);
    expect(items.map((item) => item.runs.find((run) => run.lineBreak)?.style.fontSizePt)).toEqual([11, 28]);
    expect(items.map((item) => item.runs.find((run) => run.text)?.style.fontSizePt)).toEqual([14, 14]);
  });

  it('右对齐点线制表位不随段落缩进改写源坐标', async () => {
    const { doc } = await open('tab-stops-column-edge');
    const items = paragraphs(doc.blocks).slice(0, 5);
    expect(items.map((item) => item.style.indentLeftPx ?? 0)).toEqual([0, 16, 32, 0, 16]);
    for (const item of items) {
      expect(text(item).split('\t')).toHaveLength(2);
      expect(item.style.tabStops).toHaveLength(1);
      expect(item.style.tabStops![0]).toMatchObject({ alignment: 'right', leader: 'dot' });
      expect(item.style.tabStops![0].posPx).toBeCloseTo(9746 / 15, 5);
    }
  });

  it('网格与 exact/atLeast/auto 行距保持各自语义', async () => {
    const { doc } = await open('grid-line-rules');
    const items = paragraphs(doc.blocks);
    expect(items).toHaveLength(8);
    expect(doc.page.docGrid).toMatchObject({ type: 'lines', linePitchPx: 24 });
    expect(items[1].style.lineHeight).toBe(1.5);
    expect(items[2].style.lineHeightExactPx).toBe(20);
    expect(items[3].style.lineHeightAtLeastPx).toBeCloseTo(200 / 15, 5);
    expect(items[4].style.lineHeightAtLeastPx).toBe(0);
    expect(items[6].style.snapToGrid).toBe(false);
    expect(items[7].style.lineHeight).toBe(2);
  });

  it('纯图片段落保留每一张图、不同高度及真实媒体关系', async () => {
    const { doc } = await open('para-picture-only-line');
    const images = paragraphs(doc.blocks).flatMap((item) => item.runs.flatMap((run) => run.image ? [run.image] : []));
    expect(images).toHaveLength(18);
    expect(images.every((image) => image.widthPx === 160 && image.relId === 'rIdImg0')).toBe(true);
    expect(images.map((image) => image.heightPx)).toContain(40 * 4 / 3);
    expect(images.filter((image) => image.heightPx === 80)).toHaveLength(17);
    // 功能：原件的两幅图片各含 8pt 下方效果占位，不能当成零间距图片；时间：2026-11-04；作者：lq。
    const effects = images.filter((image) => image.effectExtent);
    expect(effects).toHaveLength(2);
    expect(effects[0].effectExtent?.bottom).toBeCloseTo(32 / 3, 5);
  });

  // 功能：修改图片尺寸并强制重建正文，验证四边效果占位不会只在无操作透传时保留；时间：2026-11-04；作者：lq。
  it('效果外扩随图片尺寸编辑保存并可重新读取', async () => {
    const { doc, bytes, name } = await open('para-picture-only-line');
    const image = paragraphs(doc.blocks).flatMap((item) => item.runs).find((run) => run.image?.effectExtent)!.image!;
    image.widthPx += 1;
    image.effectExtent = { left: 2, top: 3, right: 4, bottom: 32 / 3 };
    expect(renderRunsHtml([{ image, style: {} }])).toContain('margin:3px 4px 10.666666666666666px 2px');
    const saved = await serializeDocx({ ...doc, source: undefined }, { sourceBytes: bytes, name });
    const reopened = await parseDocx(saved, { fileName: name });
    const changed = paragraphs(reopened.blocks).flatMap((item) => item.runs).find((run) => run.image?.widthPx === 161)!.image!;
    expect(changed.effectExtent).toEqual(image.effectExtent);
  });

  it('固定表格保持 2×3 单元格文字及原始 2880 twips 列宽', async () => {
    const { doc } = await open('table');
    const table = doc.blocks.find((block) => block.type === 'table');
    expect(table?.columnWidthsPx).toEqual([192, 192, 192]);
    expect(table?.rows.map((row) => row.cells.map((cell) => paragraphs(cell.blocks).map(text).join(''))))
      .toEqual([['C11', 'C12', 'C13'], ['C21', 'C22', 'C23']]);
  });
});
