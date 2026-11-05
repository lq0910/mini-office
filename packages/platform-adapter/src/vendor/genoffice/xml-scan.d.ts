/** 功能：声明原始正文 XML 区间扫描接口，用于局部保存；时间：2026-11-03；作者：lq。 */
export function scanBody(xml: string): {
  elements: { name: string; start: number; end: number }[];
  bodyContentStart: number;
  bodyContentEnd: number;
};
