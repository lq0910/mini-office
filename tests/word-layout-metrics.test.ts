/**
 * 功能：验证 P1 度量层与分页规则——行高按字体真实度量换算、页首丢弃段前距、keepNext 链同页。
 * 时间：2026-11-02；作者：lq
 */
import { beforeEach, describe, expect, it } from 'vitest';
import type { DocBlock, DocParagraph, ParagraphStyle } from '@mini-office/document-model';
import {
  lineHeightPx,
  paragraphFontSizePt,
  // 功能：回归按实际 run 字体和字号度量段落，避免文档默认字体干扰；时间：2026-11-05；作者：lq。
  paragraphLineMetrics,
  resetFontMetrics,
  singleLineHeightPx,
} from '../packages/editor-word/src/metrics';
import { paginateBlocks, type ParagraphLayout } from '../packages/editor-word/src/pagination';
// 功能：验证段落样式上下文间距按真实相邻块折叠；时间：2026-11-05；作者：lq。
import { effectiveParagraphSpacing } from '../packages/editor-word/src/paragraph-spacing';

function paragraph(id: string, style: ParagraphStyle = {}, runs: DocParagraph['runs'] = []): DocParagraph {
  return { type: 'paragraph', id, runs, style };
}

function heightsOf(blocks: DocBlock[], height: number): Map<string, number> {
  return new Map(blocks.map((block) => [block.id, height]));
}

beforeEach(() => resetFontMetrics());

describe('字体度量与行距换算', () => {
  it('倍数行距按「单行高度 × 倍数」计算，而不是「字号 × 倍数」', () => {
    const single = singleLineHeightPx('宋体', 12);
    const line = lineHeightPx({ lineHeight: 1.5 }, '宋体', 12);
    expect(line).toBeCloseTo(single * 1.5, 5);
    // CSS 直接写 1.5 时只有 12pt×1.5 = 24px；Word 的实际行高必须更高。
    expect(line).toBeGreaterThan((12 * 96) / 72 * 1.5);
  });

  it('固定值与最小值行距按 Word 语义换算', () => {
    const single = singleLineHeightPx('宋体', 12);
    expect(lineHeightPx({ lineHeightExactPx: 20 }, '宋体', 12)).toBe(20);
    expect(lineHeightPx({ lineHeightAtLeastPx: 10 }, '宋体', 12)).toBeCloseTo(single, 5);
    expect(lineHeightPx({ lineHeightAtLeastPx: 30 }, '宋体', 12)).toBe(30);
  });

  it('同一字体同一字号命中缓存，结果稳定', () => {
    const first = singleLineHeightPx('微软雅黑', 10.5);
    const second = singleLineHeightPx('微软雅黑', 10.5);
    expect(second).toBe(first);
    expect(first).toBeGreaterThan(0);
  });

  it('启用文档网格时单倍行高取网格行距', () => {
    const pitch = 24;
    expect(lineHeightPx({ lineHeight: 1 }, '宋体', 10.5, pitch)).toBeCloseTo(pitch, 5);
    expect(lineHeightPx({ lineHeight: 1.5 }, '宋体', 10.5, pitch)).toBeCloseTo(pitch * 1.5, 5);
    // 固定行距优先级最高，网格不覆盖显式设置。
    expect(lineHeightPx({ lineHeightExactPx: 30 }, '宋体', 10.5, pitch)).toBe(30);
  });

  // 功能：回归 GenOffice Word 网格探针中的大字号跨格、低倍数下限与零最小值例外；时间：2026-11-03；作者：lq。
  it('大字号占多格，低倍数不压低自然行高，最小值零不吸附', () => {
    expect(lineHeightPx({ lineHeight: 1 }, '宋体', 22, 24)).toBe(48);
    expect(lineHeightPx({ lineHeight: 0.5 }, '宋体', 10.5, 24)).toBe(24);
    expect(lineHeightPx({ lineHeightAtLeastPx: 0 }, '宋体', 10.5, 24)).toBe(singleLineHeightPx('宋体', 10.5));
  });

  it('段落行高字号取段内最大值', () => {
    expect(paragraphFontSizePt([{ style: { fontSizePt: 10.5 } }, { style: { fontSizePt: 22 } }], 10.5)).toBe(22);
    expect(paragraphFontSizePt([], 10.5)).toBe(10.5);
    expect(paragraphFontSizePt(undefined, 12)).toBe(12);
  });

  // 功能：中文 run 和空段落的有效字体应决定网格占格，小字号不得被文档默认字号抬高；时间：2026-11-05；作者：lq。
  it('使用实际 run 字体跨格，默认字体仅补未设置的值', () => {
    const chinese = paragraphLineMetrics([{ style: { fontFamily: '微软雅黑', fontSizePt: 12 } }], { lineHeight: 1.5 }, 'Calibri', 10, 20.8);
    expect(chinese.heightPx).toBe(41.6);
    expect(chinese.family).toBe('微软雅黑');
    expect(paragraphLineMetrics([{ style: { fontFamily: 'Calibri', fontSizePt: 8 } }], {}, '宋体', 20).sizePt).toBe(8);
    expect(paragraphLineMetrics([], {}, 'Calibri', 10).family).toBe('Calibri');
  });

  // 功能：字体自然高度也参与比较，不能只取字号最大的 run 并套用另一字体度量；时间：2026-11-05；作者：lq。
  it('不同字体混排按最高行盒计算，而非只取最大字号', () => {
    const mixed = paragraphLineMetrics([{ style: { fontFamily: '宋体', fontSizePt: 12 } }, { style: { fontFamily: 'Calibri', fontSizePt: 13 } }], {}, 'Calibri', 10);
    expect(mixed.family).toBe('宋体');
    expect(mixed.heightPx).toBeCloseTo(singleLineHeightPx('宋体', 12), 5);
  });
});

describe('分页规则', () => {
  // 功能：同样式上下文间距影响渲染与分页；异样式及显式关闭仍保留声明间距；时间：2026-11-05；作者：lq。
  it('同样式相邻段落折叠段前后距并同步分页高度', () => {
    const first = { ...paragraph('ctx-a', { spaceBeforePx: 12, spaceAfterPx: 8, contextualSpacing: true }), styleId: 'Body' };
    const second = { ...paragraph('ctx-b', { spaceBeforePx: 12, spaceAfterPx: 8, contextualSpacing: true }), styleId: 'Body' };
    const blocks: DocBlock[] = [first, second];
    expect(effectiveParagraphSpacing(blocks, 0)).toEqual({ beforePx: 12, afterPx: 0 });
    expect(effectiveParagraphSpacing(blocks, 1)).toEqual({ beforePx: 0, afterPx: 8 });
    // 隐藏测量高为 52px、48px；首段置于页首移除 12px 段前距后，合计 88px，85px 页面必须分两页。
    const pages = paginateBlocks(blocks, new Map([['ctx-a', 52], ['ctx-b', 48]]), 85);
    expect(pages).toHaveLength(2);

    const different = { ...second, id: 'ctx-c', styleId: 'Other' };
    expect(effectiveParagraphSpacing([first, different], 0)).toEqual({ beforePx: 12, afterPx: 8 });
    const explicitlyDisabled = { ...second, id: 'ctx-off', style: { ...second.style, contextualSpacing: false } };
    expect(effectiveParagraphSpacing([first, explicitlyDisabled], 0)?.afterPx).toBe(0);
    expect(effectiveParagraphSpacing([first, explicitlyDisabled], 1)?.beforePx).toBe(12);
  });

  it('页首丢弃段前距，不把整段挤到下一页', () => {
    const first = paragraph('a', { spaceBeforePx: 20 });
    const second = paragraph('b');
    const blocks: DocBlock[] = [first, second];
    // 实测高度含 padding：a = 20(段前) + 40(正文)，b = 60；正文区 100。
    const heights = new Map([
      ['a', 60],
      ['b', 60],
    ]);
    const pages = paginateBlocks(blocks, heights, 100);
    expect(pages).toHaveLength(1);
    expect(pages[0].blocks.map((block) => block.id)).toEqual(['a', 'b']);
  });

  it('keepNext 链整体放不下时一起换页，标题不落在页尾', () => {
    const lead = paragraph('x');
    const heading = paragraph('a', { keepNext: true });
    const body = paragraph('b');
    const blocks: DocBlock[] = [lead, heading, body];
    const heights = heightsOf(blocks, 30);
    heights.set('x', 60);
    // x=60，标题+正文各 30：不留链时页 1 只放得下 x+a，正文被挤到第 2 页。
    const pages = paginateBlocks(blocks, heights, 100);
    expect(pages).toHaveLength(2);
    expect(pages[0].blocks.map((block) => block.id)).toEqual(['x']);
    expect(pages[1].blocks.map((block) => block.id)).toEqual(['a', 'b']);
  });

  // 功能：当前页剩余空间可放标题但放不下标题加正文时，使用页首段前距规则计算整链高度；时间：2026-11-05；作者：lq。
  it('keepNext 链按新页折叠段前距后仍能整体排下', () => {
    const lead = paragraph('lead');
    const heading = paragraph('heading', { keepNext: true, spaceBeforePx: 20 });
    const body = paragraph('body');
    const blocks = [lead, heading, body];
    const heights = new Map([['lead', 35], ['heading', 65], ['body', 40]]);
    const pages = paginateBlocks(blocks, heights, 100);
    expect(pages).toHaveLength(2);
    expect(pages[0].blocks.map((block) => block.id)).toEqual(['lead']);
    expect(pages[1].blocks.map((block) => block.id)).toEqual(['heading', 'body']);
  });

  it('keepNext 链本身超过一页时不做整体移动，避免死循环或空页', () => {
    const heading = paragraph('a', { keepNext: true });
    const body = paragraph('b');
    const blocks: DocBlock[] = [heading, body];
    const heights = new Map([
      ['a', 80],
      ['b', 80],
    ]);
    const pages = paginateBlocks(blocks, heights, 100);
    expect(pages.length).toBe(2);
    expect(pages[0].blocks.map((block) => block.id)).toEqual(['a']);
  });

  it('硬分页符仍然产生独立页', () => {
    const first = paragraph('a');
    const second = paragraph('b');
    const blocks: DocBlock[] = [first, { type: 'pageBreak', id: 'br' }, second];
    const heights = heightsOf([first, second], 20);
    const pages = paginateBlocks(blocks, heights, 100);
    expect(pages).toHaveLength(2);
    expect(pages[0].blocks.map((block) => block.id)).toEqual(['a']);
    expect(pages[1].blocks.map((block) => block.id)).toEqual(['b']);
  });
});

/** 功能：构造 n 行的段落行盒；时间：2026-11-03；作者：lq。 */
function lineBoxes(count: number, lineHeight = 20): ParagraphLayout {
  return {
    lineTops: Array.from({ length: count }, (_, index) => index * lineHeight),
    height: count * lineHeight,
  };
}

describe('段落行级切分', () => {
  it('长段落按行切开，不再整段移到下一页', () => {
    const lead = paragraph('lead');
    const long = paragraph('p');
    const blocks: DocBlock[] = [lead, long];
    const heights = new Map([
      ['lead', 40],
      ['p', 120],
    ]);
    const layouts = new Map([['p', lineBoxes(6)]]);
    const pages = paginateBlocks(blocks, heights, 100, undefined, layouts);

    expect(pages).toHaveLength(2);
    expect(pages[0].blocks.map((block) => block.id)).toEqual(['lead', 'p']);
    expect(pages[1].blocks.map((block) => block.id)).toEqual(['p-part-1']);
    // 首片露出前 3 行（60px），续页从第 3 行开始、上移 60px。
    expect(pages[0].splits?.p).toMatchObject({ sourceId: 'p', start: 0, end: 3, offset: 0, height: 60 });
    expect(pages[1].splits?.['p-part-1']).toMatchObject({ sourceId: 'p', start: 3, end: 6, offset: 60, height: 60 });
    // 续页片段只读，避免编辑落到被裁剪的不可见行。
    expect(pages[1].locked).toContain('p-part-1');
    expect(pages[0].locked).toBeUndefined();
  });

  it('首片保留段前距，续页丢弃段前距', () => {
    const lead = paragraph('lead');
    const long = paragraph('p', { spaceBeforePx: 10 });
    const blocks: DocBlock[] = [lead, long];
    const heights = new Map([
      ['lead', 40],
      ['p', 130],
    ]);
    const layouts = new Map([['p', lineBoxes(6)]]);
    const pages = paginateBlocks(blocks, heights, 100, undefined, layouts);

    // 段前距 10 占掉首片预算，首片只放得下 2 行。
    expect(pages[0].splits?.p).toMatchObject({ start: 0, end: 2, offset: 0, height: 40, paddingTop: 10 });
    expect(pages[1].splits?.['p-part-1']).toMatchObject({ start: 2, end: 6, offset: 40, height: 80, paddingTop: 0 });
  });

  it('页尾只剩一行时整段移到下一页（孤行控制）', () => {
    const lead = paragraph('lead');
    const long = paragraph('p');
    const blocks: DocBlock[] = [lead, long];
    const heights = new Map([
      ['lead', 90],
      ['p', 80],
    ]);
    const layouts = new Map([['p', lineBoxes(4)]]);
    const pages = paginateBlocks(blocks, heights, 100, undefined, layouts);

    // 只剩 10px，一行都放不下 → 换页；换页后整段放得下，按普通段落处理，不产生裁剪窗口。
    expect(pages).toHaveLength(2);
    expect(pages[0].blocks.map((block) => block.id)).toEqual(['lead']);
    expect(pages[1].blocks.map((block) => block.id)).toEqual(['p']);
    expect(pages[1].splits).toBeUndefined();
  });

  it('续页开头只剩一行时本页少放一行（寡行控制）', () => {
    const long = paragraph('p');
    const blocks: DocBlock[] = [long];
    const heights = new Map([['p', 100]]);
    const layouts = new Map([['p', lineBoxes(5)]]);
    const pages = paginateBlocks(blocks, heights, 90, undefined, layouts);

    // 90px 能放 4 行，但续页只剩 1 行 → 本页只放 3 行，续页 2 行。
    expect(pages[0].splits?.p).toMatchObject({ start: 0, end: 3, height: 60 });
    expect(pages[1].splits?.['p-part-1']).toMatchObject({ start: 3, end: 5, height: 40 });
  });

  it('keepLines 段落整体移到下一页，不被拆开', () => {
    const lead = paragraph('lead');
    const clause = paragraph('p', { keepLines: true });
    const blocks: DocBlock[] = [lead, clause];
    const heights = new Map([
      ['lead', 40],
      ['p', 80],
    ]);
    const layouts = new Map([['p', lineBoxes(4)]]);
    const pages = paginateBlocks(blocks, heights, 100, undefined, layouts);

    expect(pages).toHaveLength(2);
    expect(pages[0].blocks.map((block) => block.id)).toEqual(['lead']);
    expect(pages[1].blocks.map((block) => block.id)).toEqual(['p']);
  });

  it('关闭孤行控制时允许页尾单行', () => {
    const long = paragraph('p', { widowControl: false });
    const blocks: DocBlock[] = [long];
    const heights = new Map([['p', 100]]);
    const layouts = new Map([['p', lineBoxes(5)]]);
    const pages = paginateBlocks(blocks, heights, 90, undefined, layouts);

    // 关闭后不再回退一行：本页放满 4 行，续页 1 行。
    expect(pages[0].splits?.p).toMatchObject({ start: 0, end: 4, height: 80 });
    expect(pages[1].splits?.['p-part-1']).toMatchObject({ start: 4, end: 5, height: 20 });
  });

  it('没有行盒信息时保持整段移动，行为不变', () => {
    const lead = paragraph('lead');
    const long = paragraph('p');
    const blocks: DocBlock[] = [lead, long];
    const heights = new Map([
      ['lead', 40],
      ['p', 120],
    ]);
    const pages = paginateBlocks(blocks, heights, 100);
    expect(pages).toHaveLength(2);
    expect(pages[0].blocks.map((block) => block.id)).toEqual(['lead']);
    expect(pages[1].blocks.map((block) => block.id)).toEqual(['p']);
  });
});
