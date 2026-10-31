/**
 * 功能：Word 编辑交互——多行粘贴拆成独立段落（含表格单元格内）、单元格内段首退格不跨单元格；
 * 时间：2026-10-30；作者：hou0。
 */
import { describe, expect, it } from 'vitest';
import { createBlankDoc, type DocModel, type DocTable } from '@mini-office/document-model';
import { mergeIntoPreviousParagraph, pasteMultiline } from '@mini-office/editor-word';

function cellParagraph(id: string, text: string): any {
  return { type: 'paragraph', id, style: {}, runs: [{ text, style: {} }] };
}

function tableWith(id: string, cellBlocks: any[]): DocModel {
  const table: DocTable = {
    type: 'table',
    id: 't1',
    columnWidthsPx: [100],
    rows: [{ id: 'r1', cells: [{ id: 'c1', blocks: cellBlocks, colSpan: 1, rowSpan: 1 }] }],
  };
  return { ...createBlankDoc(), blocks: [table] };
}

describe('多行粘贴', () => {
  it('按换行拆成独立段落，继承段落样式与列表，选区被粘贴内容替换', () => {
    const doc = createBlankDoc();
    doc.blocks = [
      { type: 'paragraph', id: 'p1', style: { alignment: 'center' }, list: { level: 0, marker: '1.', ordered: true }, runs: [{ text: '甲乙', style: {} }] },
    ];
    const anchor = doc.blocks[0];
    if (anchor.type !== 'paragraph') throw new Error('paragraph expected');
    const result = pasteMultiline(doc.blocks, anchor, 1, 2, '第一行\n第二行\n第三行', { fontFamily: '宋体' }, (prefix: string) => `${prefix}-1`);
    expect(result).toBeTruthy();
    const blocks = result!.blocks;
    expect(blocks).toHaveLength(3);
    const [first, second, third] = blocks as any[];
    expect(first.runs.map((r: any) => r.text).join('')).toBe('甲第一行');
    expect(second.runs.map((r: any) => r.text).join('')).toBe('第二行');
    expect(third.runs.map((r: any) => r.text).join('')).toBe('第三行');
    expect(second.style.alignment).toBe('center');
    expect(second.list?.ordered).toBe(true);
    expect(result!.lastParagraphId).toBe(third.id);
    expect(result!.caretOffset).toBe(3);
  });

  it('表格单元格内粘贴多行仍在单元格内拆段', () => {
    const doc = tableWith('t1', [cellParagraph('cp1', '单元')]);
    const anchor = (doc.blocks[0] as DocTable).rows[0].cells[0].blocks[0];
    if (anchor.type !== 'paragraph') throw new Error('paragraph expected');
    const result = pasteMultiline(doc.blocks, anchor, 2, 2, 'A\nB', {}, () => 'np1');
    expect(result).toBeTruthy();
    const cellBlocks = (result!.blocks[0] as DocTable).rows[0].cells[0].blocks;
    expect(cellBlocks).toHaveLength(2);
    expect((cellBlocks[0] as any).runs.map((r: any) => r.text).join('')).toBe('单元A');
    expect((cellBlocks[1] as any).runs.map((r: any) => r.text).join('')).toBe('B');
  });
});

describe('单元格内退格合并', () => {
  it('并回同一单元格的上一段落，不越出单元格', () => {
    const doc = tableWith('t1', [cellParagraph('cp1', '第一段'), cellParagraph('cp2', '第二段')]);
    const merged = mergeIntoPreviousParagraph(doc.blocks, 'cp2');
    expect(merged).toBeTruthy();
    const cellBlocks = (merged!.blocks[0] as DocTable).rows[0].cells[0].blocks;
    expect(cellBlocks).toHaveLength(1);
    expect((cellBlocks[0] as any).runs.map((r: any) => r.text).join('')).toBe('第一段第二段');
    expect(merged!.caretParagraphId).toBe('cp1');
    expect(merged!.caretOffset).toBe(3);
  });

  it('单元格内首段不可并出单元格', () => {
    const doc = tableWith('t1', [cellParagraph('cp1', '内容')]);
    expect(mergeIntoPreviousParagraph(doc.blocks, 'cp1')).toBeNull();
  });
});
