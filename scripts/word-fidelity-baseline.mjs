/** 功能：验证外部 Office 基准的身份、源文件哈希和图像清单，拒绝以浏览器截图冒充 Office；时间：2026-11-03；作者：lq。 */
import { resolve, relative, isAbsolute } from 'node:path';

/** 功能：约束基准图像只读路径，拒绝绝对路径和上级目录跳转；时间：2026-11-03；作者：lq。 */
export function baselinePath(root, file) {
  if (typeof file !== 'string' || !file || isAbsolute(file) || file.includes('\\')) throw new Error('基准路径必须是相对路径');
  const path = resolve(root, file);
  const inside = relative(resolve(root), path);
  if (inside === '..' || inside.startsWith('../') || isAbsolute(inside)) throw new Error('基准路径越界');
  return path;
}

/** 功能：严格校验环境与文档绑定，缺失或陈旧基准不可计为通过；时间：2026-11-03；作者：lq。 */
export function validateOfficeBaseline(value, fixture) {
  const fail = (message) => { throw new Error(`${fixture.id}: ${message}`); };
  if (!value || value.schemaVersion !== 1) fail('不支持的基准格式');
  if (!['Microsoft Word', 'WPS Writer'].includes(value.application)) fail('基准必须明确来自 Microsoft Word 或 WPS Writer');
  for (const key of ['version', 'operatingSystem', 'exportedAt', 'exportMethod']) {
    if (typeof value[key] !== 'string' || !value[key].trim()) fail(`缺少 ${key}`);
  }
  if (!Number.isFinite(Date.parse(value.exportedAt))) fail('exportedAt 日期无效');
  if (value.dpi !== 96) fail('基准必须按 96 DPI 导出，禁止隐式缩放');
  if (!Array.isArray(value.fonts) || !value.fonts.length || value.fonts.some((font) => typeof font !== 'string' || !font.trim())) fail('缺少实际字体环境');
  if (value.fixtureId !== fixture.id || value.sourceSha256 !== fixture.sha256) fail('基准与源文件版本不匹配');
  if (typeof value.pdfFile !== 'string' || !/\.pdf$/i.test(value.pdfFile)) fail('缺少原始 Office PDF 文件');
  baselinePath('/baseline', value.pdfFile);
  if (!/^[a-f0-9]{64}$/.test(value.pdfSha256 ?? '')) fail('缺少原始 Office PDF 的 SHA-256');
  if (!Array.isArray(value.pages) || !value.pages.length) fail('基准没有页面');
  const seen = new Set();
  for (const page of value.pages) {
    if (!page || typeof page.file !== 'string' || !/\.png$/i.test(page.file) || !/^[a-f0-9]{64}$/.test(page.sha256 ?? '')) fail('页面缺少 PNG 或哈希');
    baselinePath('/baseline', page.file);
    if (seen.has(page.file)) fail('基准页面重复');
    seen.add(page.file);
  }
  return value;
}
