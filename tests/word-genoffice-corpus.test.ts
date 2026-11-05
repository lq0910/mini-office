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
