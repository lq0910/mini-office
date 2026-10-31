/**
 * 功能：分页保真——keepNext 链（与下段同页）与孤寡控制（跨页断点两侧至少两行）；
 * 时间：2026-10-31；作者：hou0。
 */
import { describe, expect, it } from 'vitest';
import { createBlankDoc, type DocModel } from '@mini-office/document-model';
import { applyWidowControl, paginateBlocks } from '@mini-office/editor-word';

function para(id: string, text: string, style: Record<string, unknown> = {}): any {
  return { type: 'paragraph', id, style, runs: [{ text, style: {} }] };
}

describe('keepNext 链', () => {
  it('标题放得下但表格放不下时，标题随表格推到新页', () => {
    const doc: DocModel = { ...createBlankDoc(), blocks: [
      para('p', '前置段落'),
      para('t', '表格标题', { keepNext: true }),
      { type: 'table', id: 'tbl', columnWidthsPx: [100], rows: [{ id: 'r', cells: [{ id: 'c', blocks: [], colSpan: 1, rowSpan: 1 }] }] },
    ] };
    // 页高 60：前置段落(20)+标题(20)放得下，但标题的 keepNext 要求表格(200)同页放不下 → 标题推走
    const heights = new Map([['p', 20], ['t', 20], ['tbl', 200]]);
    const pages = paginateBlocks(doc.blocks, heights, 60);
    expect(pages).toHaveLength(3);
    expect(pages[0].blocks.map((b) => b.id)).toEqual(['p']);
    expect(pages[1].blocks.map((b) => b.id)).toEqual(['t']);
    expect(pages[2].blocks.map((b) => b.id)).toEqual(['tbl']);
  });

  it('标题在页首、表格放不下时无处可推，标题留在本页', () => {
    const doc: DocModel = { ...createBlankDoc(), blocks: [
      para('t', '表格标题', { keepNext: true }),
      { type: 'table', id: 'tbl', columnWidthsPx: [100], rows: [{ id: 'r', cells: [{ id: 'c', blocks: [], colSpan: 1, rowSpan: 1 }] }] },
    ] };
    const heights = new Map([['t', 20], ['tbl', 200]]);
    const pages = paginateBlocks(doc.blocks, heights, 60);
    expect(pages).toHaveLength(2);
    expect(pages[0].blocks.map((b) => b.id)).toEqual(['t']);
    expect(pages[1].blocks.map((b) => b.id)).toEqual(['tbl']);
  });

  it('无 keepNext 时标题仍可留在页尾', () => {
    const doc: DocModel = { ...createBlankDoc(), blocks: [
      para('t', '普通段落'),
      { type: 'table', id: 'tbl', columnWidthsPx: [100], rows: [{ id: 'r', cells: [{ id: 'c', blocks: [], colSpan: 1, rowSpan: 1 }] }] },
    ] };
    const heights = new Map([['t', 20], ['tbl', 200]]);
    const pages = paginateBlocks(doc.blocks, heights, 60);
    // 标题占 20，表格放不下 → 换页；但标题没有 keepNext，不连带
    expect(pages).toHaveLength(2);
    expect(pages[0].blocks.map((b) => b.id)).toEqual(['t']);
    expect(pages[1].blocks.map((b) => b.id)).toEqual(['tbl']);
  });

  it('keepNext 链超过整页时放弃推页，避免死循环', () => {
    const blocks = [
      para('a', 'a', { keepNext: true }),
      para('b', 'b', { keepNext: true }),
      para('c', 'c', { keepNext: true }),
      para('d', 'd'),
    ];
    const heights = new Map([['a', 300], ['b', 300], ['c', 300], ['d', 20]]);
    const pages = paginateBlocks(blocks, heights, 100);
    expect(pages.length).toBeGreaterThanOrEqual(2);
  });
});

describe('孤寡控制', () => {
  const LH = 20;
  // 线性测量：每个模型偏移 20px（即按"1 offset = 1 行"近似），含 5px padding
  const make = (totalLines: number) => ({
    measure: (offset: number) => offset * LH + 5,
    totalHeight: totalLines * LH + 10,
  });

  it('头部只剩一行：整段移到下页（返回 null）', () => {
    const { measure, totalHeight } = make(5);
    // headEnd=1 → 头部 1 行
    expect(applyWidowControl({ headEnd: 1, totalHeight, measure, lineHeightPx: LH, padY: 10, maxHeight: 200 })).toBeNull();
  });

  it('尾部只剩一行：断点前移，让尾部保留两行', () => {
    const { measure, totalHeight } = make(5);
    // headEnd=4 → 尾部 1 行 → 调整到 headEnd=3（尾部 2 行）
    expect(applyWidowControl({ headEnd: 4, totalHeight, measure, lineHeightPx: LH, padY: 10, maxHeight: 200 })).toBe(3);
  });

  it('两侧各两行以上：断点不变', () => {
    const { measure, totalHeight } = make(6);
    expect(applyWidowControl({ headEnd: 3, totalHeight, measure, lineHeightPx: LH, padY: 10, maxHeight: 200 })).toBe(3);
  });

  it('总行数不足三行：不做控制', () => {
    const { measure, totalHeight } = make(2);
    expect(applyWidowControl({ headEnd: 1, totalHeight, measure, lineHeightPx: LH, padY: 10, maxHeight: 200 })).toBe(1);
  });

  it('段内强制分页符：不参与孤寡调整', () => {
    const { measure, totalHeight } = make(5);
    expect(applyWidowControl({ headEnd: 1, totalHeight, measure, lineHeightPx: LH, padY: 10, maxHeight: 200, forced: true })).toBe(1);
  });
});
