/** 功能：覆盖跨页合并格长文字的完整行放置与页缝列网格，防止遮字和扩列；时间：2026-11-11；作者：lq。 */
import { describe, expect, it } from 'vitest';
import { wordCellLineShifts } from '../packages/word-core/src/layout/cell-line-placement';
import { wordTableGapSpans } from '../packages/editor-word/src/core-table-gaps';

describe('跨页合并单元格', () => {
  it('长文字优先保留原位置，遇到页脚把完整行移到下一段正文', () => {
    const lines = [{ top: 80, bottom: 90 }, { top: 95, bottom: 105 }, { top: 110, bottom: 120 }];
    expect(wordCellLineShifts(lines, [{ top: 0, bottom: 100 }, { top: 160, bottom: 200 }])).toEqual([0, 65, 65]);
  });
  it('居中后余下空间不足时从首片段放置，所有文字仍在单元格内', () => {
    const lines = Array.from({ length: 5 }, (_, index) => ({ top: 70 + index * 15, bottom: 80 + index * 15 }));
    const shifts = wordCellLineShifts(lines, [{ top: 0, bottom: 45 }, { top: 100, bottom: 140 }]);
    expect(shifts).toEqual([-70, -70, -70, -15, -15]);
  });
  it('容纳不了完整文字时拒绝返回会覆盖邻格的部分结果', () => {
    expect(wordCellLineShifts([{ top: 10, bottom: 50 }], [{ top: 0, bottom: 20 }, { top: 100, bottom: 120 }])).toBeUndefined();
  });
  it('页缝跳过前两列的合并格，不再插入全宽 td 撑出额外列', () => {
    const table = document.createElement('table');
    table.innerHTML = '<tbody><tr><td rowspan="3"></td><td rowspan="3"></td><td></td><td></td></tr><tr><td></td><td></td></tr><tr><td></td><td></td></tr></tbody>';
    expect(wordTableGapSpans(table, 1)).toEqual([2]);
    expect(wordTableGapSpans(table, 3)).toEqual([4]);
  });
  it('中间列合并时保留两侧独立空列段，并读取来源 rowspan', () => {
    const table = document.createElement('table');
    table.innerHTML = '<tbody><tr><td></td><td rowspan="4" data-base-rowspan="2"></td><td colspan="2"></td></tr><tr><td></td><td colspan="2"></td></tr></tbody>';
    expect(wordTableGapSpans(table, 1)).toEqual([1, 2]);
    expect(wordTableGapSpans(table, 2)).toEqual([4]);
  });
});
