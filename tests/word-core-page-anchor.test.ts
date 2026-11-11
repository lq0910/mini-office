/** 功能：覆盖跨页表格尾部无文字切点，禁止把未定位正文误当留白；时间：2026-11-11；作者：lq。 */
import { describe, expect, it } from 'vitest';
import { resolveWordPageAnchor, wordTableRowRects } from '../packages/editor-word/src/core-page-anchor';
import type { DomLineRectsFn } from '../packages/word-core/src/upstream/apps/docs/src/renderer/pagination-lines';

describe('Word 换页锚点', () => {
  // 功能：模拟首列下一行文字先于后列续行出现在 DOM 中，主锚点仍应属于计划行；时间：2026-11-11；作者：lq。
  it('多列/合并格按计划行与视觉位置选择续行，不进入相邻行', () => {
    const table = document.createElement('table');
    table.innerHTML = '<tbody><tr><td rowspan="2"><img></td><td><img></td></tr><tr><td><img></td></tr><tr><td><img></td><td><img></td></tr></tbody>';
    const images = Array.from(table.querySelectorAll('img'));
    const rectsOf: DomLineRectsFn = () => [
      { node: images[3], offset: 900, bottom: 920, top: 900, left: 0 },
      { node: images[2], offset: 430, bottom: 450, top: 430, left: 80 },
      { node: images[0], offset: 420, bottom: 440, top: 420, left: 0 },
    ];
    const filtered = wordTableRowRects(table, 1, rectsOf);
    expect(filtered(table, 1).map(line => line.node)).toEqual([images[0], images[2]]);
    expect(resolveWordPageAnchor(table, 410, 1000, 1, filtered)).toEqual({ node: images[0], charOffset: 0 });
    expect(resolveWordPageAnchor(table, 425, 1000, 1, filtered)).toEqual({ node: images[2], charOffset: 0 });
  });
  // 功能：嵌套表格文字不会被外层行过滤器丢失，继续交给原有覆盖层处理；时间：2026-11-11；作者：lq。
  it('计划行保留嵌套表格内的真实文字锚点', () => {
    const table = document.createElement('table');
    table.innerHTML = '<tbody><tr><td><table><tbody><tr><td><img></td></tr></tbody></table></td></tr></tbody>';
    const image = table.querySelector('img')!;
    const rectsOf: DomLineRectsFn = () => [{ node: image, offset: 50, bottom: 70, top: 50, left: 0 }];
    expect(wordTableRowRects(table, 0, rectsOf)(table, 1)).toHaveLength(1);
  });
  /** 功能：用确定行几何覆盖真实样本表格最后字形与块末边框之间的切点；时间：2026-11-11；作者：lq。 */
  function fixture() {
    const element = document.createElement('table');
    const image = document.createElement('img'); element.append(image);
    const rectsOf: DomLineRectsFn = () => [{ offset: 2429.359375, bottom: 2441.359375, left: 0, top: 2429.359375, node: image }];
    return { element, image, rectsOf };
  }
  it('表格尾部留白使用真实块末，不要求不存在的文字锚点', () => {
    const { element, rectsOf } = fixture();
    expect(resolveWordPageAnchor(element, 2454.31875, 2455.609375, 1, rectsOf)).toBe('after-block');
  });
  it('WebKit 切点穿过末行字形时使用该整行的真实开头，补偿量不受 2px 容差限制', () => {
    const { element, image, rectsOf } = fixture();
    for (const offset of [2435, 2439.75]) expect(resolveWordPageAnchor(element, offset, 2455.609375, 1, rectsOf)).toEqual({node:image,charOffset:0,layoutOffset:2429.359375});
  });
  it('真实后续行和匹配行优先于块末留白', () => {
    const { element, image, rectsOf } = fixture();
    for (const offset of [2400, 2429.359375]) expect(resolveWordPageAnchor(element, offset, 2455.609375, 1, rectsOf)).toEqual({ node: image, charOffset: 0 });
  });
  it('超出块范围、非法坐标或无测量行时继续拒绝猜测', () => {
    const { element, rectsOf } = fixture();
    for (const offset of [-1, 2500, NaN, Infinity]) expect(resolveWordPageAnchor(element, offset, 2455.609375, 1, rectsOf)).toBeNull();
    expect(resolveWordPageAnchor(element, 10, 20, 1, () => [])).toBeNull();
  });
});
