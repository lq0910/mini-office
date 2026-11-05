import { describe, expect, it } from 'vitest';
import { fontStackFor, availableFontChoices } from '@mini-office/design-system';

describe('fontStackFor', () => {
  it('西文字体走随包度量兼容字体', () => {
    expect(fontStackFor('Calibri')).toContain('Carlito');
    expect(fontStackFor('Cambria')).toContain('Caladea');
    expect(fontStackFor('Times New Roman')).toContain('Liberation Serif');
  });

  it('中文字体按字形风格回退，而不是一律变成黑体', () => {
    const song = fontStackFor('宋体');
    const fangsong = fontStackFor('仿宋');
    const kai = fontStackFor('楷体');
    const hei = fontStackFor('黑体');
    expect(song).toContain('Songti SC');
    expect(fangsong).toContain('FangSong');
    expect(kai).toContain('Kaiti');
    expect(hei).toContain('PingFang SC');
    // 宋体/仿宋/楷体都不应回退到以 PingFang SC 打头的黑体栈
    expect(song.startsWith('"PingFang SC"')).toBe(false);
    expect(fangsong.startsWith('"PingFang SC"')).toBe(false);
    expect(kai.startsWith('"PingFang SC"')).toBe(false);
  });

  it('华文/方正等别名与对应的基础字体等价', () => {
    expect(fontStackFor('华文中宋')).toContain('STZhongsong');
    expect(fontStackFor('华文仿宋')).toContain('STFangsong');
    expect(fontStackFor('华文楷体')).toContain('STKaiti');
    expect(fontStackFor('方正小标宋简体')).toContain('Songti SC');
    expect(fontStackFor('仿宋_GB2312')).toContain('FangSong');
    expect(fontStackFor('楷体_GB2312')).toContain('Kaiti');
  });

  it('表里没有的厂商字体按名称关键字推断回退栈，并保留原字体名在栈首', () => {
    const unknownSong = fontStackFor('方正兰亭宋');
    expect(unknownSong.startsWith('"方正兰亭宋"')).toBe(true);
    expect(unknownSong).toContain('Songti SC');

    const unknownKai = fontStackFor('汉仪楷体简');
    expect(unknownKai).toContain('Kaiti');

    const unknownHei = fontStackFor('思源黑体 Medium');
    expect(unknownHei).toContain('PingFang SC');
  });

  it('空值与默认值走文档默认字体栈', () => {
    expect(fontStackFor(undefined)).toContain('Songti SC');
    expect(fontStackFor('默认')).toContain('Songti SC');
    expect(fontStackFor('   ')).toContain('Songti SC');
  });
});

describe('availableFontChoices', () => {
  it('优先使用平台枚举到的真实字体清单', () => {
    const list = availableFontChoices('宋体', ['PingFang SC', 'Songti SC', 'Helvetica']);
    expect(list[0]).toBe('默认');
    expect(list).toContain('Songti SC');
    expect(list).toContain('Helvetica');
    // 当前字体即使不在清单里也要保留，否则下拉会显示空白
    expect(list).toContain('宋体');
  });

  it('没有清单时退回探测结果，且不会重复列出当前字体', () => {
    const list = availableFontChoices('PingFang SC', []);
    expect(list.filter((name) => name === 'PingFang SC')).toHaveLength(1);
  });

  // 功能：复现同一字体不同字重造成重复选项和内部字体混入，正常外文字体仍可选择；时间：2026-11-03；作者：lq。
  it('隐藏内部和异常名称，并合并字体族而非删除外文字体', () => {
    expect(availableFontChoices(undefined, [
      '.Hiragino Kaku Gothic Interface', 'Verdana', 'Verdana', ' verdana ',
      '宋体', '宋体', 'دمشق', 'दमिश्क', '', '默认', 'Broken\u0000Font', 'Broken\ufffdFont',
    ])).toEqual(['默认', 'Verdana', '宋体', 'دمشق', 'दमिश्क']);
  });

  // 功能：名称归一化去重后仍须保留受控 select 的精确当前值，选择字体不改变原文档的族名；时间：2026-11-03；作者：lq。
  it('当前字体大小写或 Unicode 形式不同也只保留一项且仍能选中', () => {
    expect(availableFontChoices('verdana', ['Verdana', 'Verdana'])).toEqual(['默认', 'verdana']);
    expect(availableFontChoices('Cafe\u0301', ['Café', 'Cafe\u0301'])).toEqual(['默认', 'Cafe\u0301']);
  });

  // 功能：文档已使用的内部或缺失字体仍显示当前值，不在打开时自动替换成默认字体；时间：2026-11-03；作者：lq。
  it('只为当前文档保留内部字体或尚未安装的字体', () => {
    expect(availableFontChoices('.Hiragino Kaku Gothic Interface', [
      '.Hiragino Kaku Gothic Interface', '.OtherInternalFont', 'Arial',
    ])).toEqual(['默认', 'Arial', '.Hiragino Kaku Gothic Interface']);
    expect(availableFontChoices('方正小标宋简体', ['Arial'])).toEqual(['默认', 'Arial', '方正小标宋简体']);
  });
});
