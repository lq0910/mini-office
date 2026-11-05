/**
 * 功能：为常用 Office 中文字体提供 Word 单倍行高，避免拿浏览器字形边界直接当作 Word 行距。
 * 时间：2026-11-05；作者：lq。
 * 数据依据：GenOffice e1f50e90 packages/font-metrics/word-faces.json 与 src/resolve.ts，Apache-2.0，许可见 third_party/genoffice。
 * 仅包含五个字体族的设计度量，不包含字体文件；未知字体仍使用实际测量。
 */
const FONT_FACTORS: Readonly<Record<string, number>> = {
  fangsong: 1.3, 仿宋: 1.3,
  simhei: 1.3, 黑体: 1.3,
  simsun: 1.3, 宋体: 1.3,
  kaiti: 1.3, 楷体: 1.3,
  'microsoft yahei': 1.7157, 微软雅黑: 1.7157,
  // 功能：确实安装旧 GB 字体时仍按该字体族度量；缺失替代由解析器根据已安装字体决定；时间：2026-11-05；作者：lq。
  fangsong_gb2312: 1.3, 仿宋_gb2312: 1.3,
  kaiti_gb2312: 1.3, 楷体_gb2312: 1.3,
};

/** 功能：显式替代已由导入阶段确认时按替代字体度量，其他情况使用主字体的设计值；时间：2026-11-05；作者：lq。 */
export function wordFontLineFactor(family?: string, alternateFamily?: string): number | undefined {
  const key = (alternateFamily ?? family)?.normalize('NFKC').trim().toLowerCase();
  return key ? FONT_FACTORS[key] : undefined;
}
