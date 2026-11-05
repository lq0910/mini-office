/** 功能：为视觉基准校验器提供类型约束，浏览器脚本保持无编译依赖；时间：2026-11-03；作者：lq。 */
export interface OfficeBaseline {
  schemaVersion: 1;
  application: 'Microsoft Word' | 'WPS Writer';
  version: string;
  operatingSystem: string;
  exportedAt: string;
  exportMethod: string;
  dpi: 96;
  fonts: string[];
  fixtureId: string;
  sourceSha256: string;
  pdfFile: string;
  pdfSha256: string;
  pages: { file: string; sha256: string }[];
}
export function baselinePath(root: string, file: string): string;
export function validateOfficeBaseline(value: unknown, fixture: { id: string; sha256: string }): OfficeBaseline;
