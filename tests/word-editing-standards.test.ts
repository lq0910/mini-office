/**
 * 功能：Word 编辑器交互标准——系统字体枚举与下拉合并、正文 Tab 插入真制表符（选区替换）、
 * Shift+Tab 不插入内容；时间：2026-10-30；作者：hou0。
 */
import { describe, expect, it } from 'vitest';
import { insertTabAt, mergeFontChoices } from '@mini-office/editor-word';

describe('系统字体', () => {
  it('下拉合并：默认 → 当前 → 文档默认 → 系统字体 → 内置，去重且当前字体始终可见', () => {
    const choices = mergeFontChoices(
      ['微软雅黑', '华文中宋', 'Consolas'],
      '华文中宋',
      '仿宋',
    );
    expect(choices[0]).toBe('默认');
    expect(choices[1]).toBe('华文中宋');
    expect(choices.indexOf('仿宋')).toBeGreaterThan(1);
    expect(choices).toContain('微软雅黑');
    expect(choices).toContain('Consolas');
    expect(choices).toContain('宋体');
    // 去重：系统字体与内置列表重叠时只出现一次
    const deduped = mergeFontChoices(['宋体', 'SimSun'], undefined, '宋体');
    expect(deduped.filter((font) => font === '宋体')).toHaveLength(1);
  });

  it('空系统列表时仍提供内置字体，且当前值不丢', () => {
    const choices = mergeFontChoices([], '方正舒体');
    expect(choices).toContain('方正舒体');
    expect(choices).toContain('Times New Roman');
  });
});

describe('正文 Tab', () => {
  const style = { fontFamily: '宋体', fontSizePt: 10.5 };
  const text = (t: string): any => ({ text: t, style: {} });

  it('折叠光标：在插入点插入一个 tab run，前后 run 保留', () => {
    const runs = insertTabAt([text('甲'), text('乙')], 1, 1, style);
    expect(runs).toHaveLength(3);
    expect(runs[0].text).toBe('甲');
    expect(runs[1].tab).toBe(true);
    expect(runs[1].style).toMatchObject({ fontFamily: '宋体' });
    expect(runs[2].text).toBe('乙');
  });

  it('选区：[start,end) 被替换为一个 tab，边界处的 run 正确截断', () => {
    const runs = insertTabAt([text('abcdef')], 2, 5, style);
    expect(runs).toHaveLength(3);
    expect(runs[0].text).toBe('ab');
    expect(runs[1].tab).toBe(true);
    expect(runs[2].text).toBe('f');
  });

  it('选区覆盖原子 run（tab/图片）时整体替换', () => {
    const atomic = { text: '', style: {}, tab: true };
    const runs = insertTabAt([text('a'), atomic, text('c')], 0, 3, style);
    expect(runs).toHaveLength(1);
    expect(runs[0].tab).toBe(true);
  });

  it('越界偏移被钳制，不抛错', () => {
    const runs = insertTabAt([text('abc')], 10, 20, style);
    expect(runs).toHaveLength(2);
    expect(runs[0].text).toBe('abc');
    expect(runs[1].tab).toBe(true);
  });
});
