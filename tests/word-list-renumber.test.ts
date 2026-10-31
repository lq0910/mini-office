/**
 * 功能：列表重编号——分段后编号递增（不再继承源段落的静态标记）、
 * numFmt="none" 的无编号列表不渲染标记、回车后也不会出现凭空数字；时间：2026-10-31。
 */
import { describe, expect, it } from 'vitest';
import { createBlankDoc, type DocModel } from '@mini-office/document-model';
import { renumberListMarkers, renderParagraphHtml } from '@mini-office/editor-word';

function numberedDoc(markers: string[], withDefs: boolean): DocModel {
  const doc = createBlankDoc();
  doc.blocks = markers.map((marker, index) => ({
    type: 'paragraph' as const,
    id: `p${index}`,
    style: {},
    list: { numId: 'n1', level: 0, marker, ordered: true },
    runs: [{ text: `内容${index}`, style: {} }],
  }));
  if (withDefs) {
    doc.numberingLevels = { n1: [{ ilvl: 0, numFmt: 'decimal', lvlText: '%1、', start: 1 }] };
  }
  return doc;
}

describe('列表重编号', () => {
  it('分段后标记相同的段落重编号为递增序列（十进制 + 中文顿号）', () => {
    const doc = numberedDoc(['1、', '1、', '1、'], true);
    const renumbered = renumberListMarkers(doc);
    const markers = (renumbered.blocks as any[]).map((b) => b.list.marker);
    expect(markers).toEqual(['1、', '2、', '3、']);
  });

  it('无编号定义时按原标记的数字增量兜底', () => {
    const doc = numberedDoc(['1、', '1、'], false);
    const markers = (renumberListMarkers(doc).blocks as any[]).map((b) => b.list.marker);
    expect(markers).toEqual(['1、', '2、']);
  });

  it('中文序号与字母序号按 numFmt 格式化', () => {
    const doc = createBlankDoc();
    doc.numberingLevels = {
      cn: [{ ilvl: 0, numFmt: 'chineseCounting', lvlText: '%1、', start: 1 }],
      ab: [{ ilvl: 0, numFmt: 'upperLetter', lvlText: '%1.', start: 1 }],
    };
    doc.blocks = [
      { type: 'paragraph', id: 'a', style: {}, list: { numId: 'cn', level: 0, marker: '一、', ordered: true }, runs: [{ text: 'x', style: {} }] },
      { type: 'paragraph', id: 'b', style: {}, list: { numId: 'cn', level: 0, marker: '一、', ordered: true }, runs: [{ text: 'y', style: {} }] },
      { type: 'paragraph', id: 'c', style: {}, list: { numId: 'ab', level: 0, marker: 'A.', ordered: true }, runs: [{ text: 'z', style: {} }] },
      { type: 'paragraph', id: 'd', style: {}, list: { numId: 'ab', level: 0, marker: 'A.', ordered: true }, runs: [{ text: 'w', style: {} }] },
    ];
    const markers = (renumberListMarkers(doc).blocks as any[]).map((b) => b.list.marker);
    expect(markers).toEqual(['一、', '二、', 'A.', 'B.']);
  });

  it('numFmt=none 的无编号列表：标记清空且渲染不出现数字', () => {
    const doc = createBlankDoc();
    doc.numberingLevels = { none: [{ ilvl: 0, numFmt: 'none', lvlText: '', start: 1 }] };
    doc.blocks = [
      { type: 'paragraph', id: 'a', style: {}, list: { numId: 'none', level: 0, marker: '1', ordered: false }, runs: [{ text: '无编号内容', style: {} }] },
    ];
    const renumbered = renumberListMarkers(doc);
    const paragraph = renumbered.blocks[0] as any;
    expect(paragraph.list.marker).toBe('');
    const html = renderParagraphHtml(paragraph);
    expect(html).not.toContain('data-list-marker');
    expect(html).not.toContain('>1<');
  });

  it('无列表的文档不做任何改动', () => {
    const doc = createBlankDoc();
    const renumbered = renumberListMarkers(doc);
    expect(renumbered.blocks).toEqual(doc.blocks);
  });
});
