/** 功能：按新核心实际保留及产品可操作范围扫描 DOCX，报告不经过旧正文模型；时间：2026-11-11；作者：lq。 */
import type { CapabilityReport, WarningSeverity } from '@mini-office/shared';
import type { OoxmlPackage } from '@mini-office/ooxml';
import { decodeEntities } from '../../word-core/src/upstream/packages/docx-engine/src/parse-xml-text';
import { checkMissingFonts } from '../../word-core/src/upstream/apps/docs/src/renderer/font-check';

/** 功能：浏览器采用真实字体度量，桌面采用系统字体及其别名，避免把通用回退误认成原字体；时间：2026-11-11；作者：lq。 */
export function coreMissingFonts(used: string[], available: string[] = [], embedded: string[] = []): string[] {
  const normalize = (value: string) => value.toLowerCase().replace(/[\s_-]+/g, '');
  const bundled = new Set(embedded.map(normalize));
  const fonts = [...new Set(used)].filter(font => !bundled.has(normalize(font)) && !/^(serif|sans-serif|monospace|system-ui|系统默认)$/i.test(font));
  if (!available.length) return checkMissingFonts(fonts).map(font => font.name);
  const installed = new Set(available.map(normalize));
  return fonts.filter(font => !installed.has(normalize(font)));
}

/** 功能：读取正文、样式、所有页眉页脚和注释区域的真实部件，未使用的字体表条目不制造缺字警告；时间：2026-11-11；作者：lq。 */
export async function analyzeCoreWordCapabilities(pkg: OoxmlPackage, knownFonts: string[] = []): Promise<CapabilityReport> {
  const report: CapabilityReport = { level: 3, warnings: [], features: [], missingFonts: [], scannedAt: Date.now() };
  const parts = pkg.names();
  const xmlParts = parts.filter(name => /^word\/(?:document|styles|numbering|settings|header[^/]*|footer[^/]*|footnotes|endnotes|comments)\.xml$/.test(name));
  const xml = (await Promise.all(xmlParts.map(name => pkg.text(name)))).join('\n');
  const fonts = new Set<string>();
  for (const tag of xml.match(/<w:rFonts\b[^>]*>/g) ?? []) for (const match of tag.matchAll(/w:(?:ascii|hAnsi|eastAsia|cs)=["']([^"']+)["']/g)) fonts.add(decodeEntities(match[1]));
  // 功能：嵌入字体由同源字体加载器采用，不因宿主系统未安装同名字体而误报；时间：2026-11-11；作者：lq。
  const fontTable = pkg.has('word/fontTable.xml') ? await pkg.text('word/fontTable.xml') : '';
  const embedded = [...fontTable.matchAll(/<w:font\b[^>]*w:name=["']([^"']+)["'][^>]*>([\s\S]*?)<\/w:font>/g)].filter(match => /<w:embed(?:Regular|Bold|Italic)/.test(match[2])).map(match => decodeEntities(match[1]));
  report.missingFonts = coreMissingFonts([...fonts], knownFonts, embedded);
  const feature = (code: string, label: string, supported = true, note?: string, severity: WarningSeverity = supported ? 'info' : 'warning') => {
    report.features.push({ code, label, supported, severity, note });
    if (!supported && note) report.warnings.push({ code, message: note, severity });
  };
  feature('word-core', 'Word 文字与段落', true, '由当前 Word 会话编辑并保存，未修改的来源部件保留。');
  if (/<w:tbl\b/.test(xml)) feature('docx.tables', '表格');
  if (/<w:(?:drawing|pict)\b/.test(xml)) feature('docx.images', '图片与图形', true, '支持图片插入、尺寸和环绕；复杂形状保留来源。');
  if (/<w:(?:headerReference|footerReference)\b/.test(xml)) feature('docx.headerFooter', '页眉页脚', true, '普通文字和页码可编辑；复杂区域保留来源。');
  if (/<w:numPr\b/.test(xml)) feature('docx.numbering', '项目符号与编号');
  if (/\bTOC\b/.test(xml)) feature('docx.toc', '目录', true, '普通标题目录可插入更新；复杂域边界受到保护。');
  if (/<w:(?:instrText|fldSimple)\b/.test(xml)) feature('docx.fields', '域与页码', true, '页码及目录按布局更新，其他域缓存保留，不连接外部数据。');
  if (parts.some(name => /(?:^|\/)vbaProject\.bin$/i.test(name))) feature('docx.macro', 'VBA 宏', false, '文档包含 VBA 宏，当前以只读方式打开并保留原文件。', 'blocking');
  if (parts.some(name => /^_xmlsignatures\//.test(name))) feature('docx.signature', '数字签名', false, '文档包含数字签名，编辑会使签名失效，当前以只读方式打开。', 'blocking');
  if (/<w:documentProtection\b[^>]*w:enforcement=["'](?:1|true|on)["']/.test(xml)) feature('docx.protection', '文档编辑保护', false, '文档已启用编辑保护，当前以只读方式打开。', 'blocking');
  if (/<w:(?:ins|del|moveFrom|moveTo|pPrChange|rPrChange|tblPrChange|tcPrChange)\b/.test(xml)) feature('docx.revisions', '修订记录', false, '修订记录由核心保留；当前工具栏尚不提供完整的接受、拒绝与修订审阅操作，建议另存副本。');
  if (parts.includes('word/comments.xml') || /<w:commentReference\b/.test(xml)) feature('docx.comments', '批注', false, '批注部件和引用保留；当前工具栏尚不提供完整批注编辑，建议另存副本。');
  if (/<w:sdt\b/.test(xml)) feature('docx.contentControls', '内容控件', false, '保留内容控件及其文本，暂不提供控件配置、数据绑定和表单交互，建议另存副本。');
  if (parts.some(name => /^word\/embeddings\//.test(name)) || /<w:object\b/.test(xml)) feature('docx.embedded', '嵌入对象', false, '嵌入对象部件保留，当前不能启动外部应用编辑对象，建议另存副本。');
  if (parts.some(name => /^word\/charts\/[^/]+\.xml$/.test(name))) feature('docx.charts', '图表', false, '图表由核心显示并保留；当前工具栏尚未开放完整图表数据及样式编辑，建议另存副本。');
  if (/<w:altChunk\b/.test(xml)) feature('docx.altChunk', '外部内容块', false, '外部内容块保留原部件，当前不展开编辑，建议另存副本。');
  if (/<w:(?:footnoteReference|endnoteReference)\b/.test(xml)) feature('docx.notes', '脚注与尾注', false, '脚注和尾注由核心显示并保留；当前工具栏尚未开放完整注释编辑，建议另存副本。');
  if (/<w:trackRevisions\b(?![^>]*w:val=["'](?:0|false|off)["'])/.test(xml)) feature('docx.trackChangesOn', '修订跟踪设置', false, '来源文档开启修订跟踪；当前工具栏尚未开放完整审阅工作流，建议另存副本。');
  if (report.missingFonts.length) report.warnings.push({ code: 'fonts.missing', message: `缺少 ${report.missingFonts.join('、')} 字体，将采用替代字体，版面可能变化。`, detail: report.missingFonts.join(', '), severity: 'warning' });
  report.level = report.warnings.some(warning => warning.severity === 'blocking') ? 1 : report.warnings.some(warning => warning.severity === 'warning' && warning.code !== 'fonts.missing') ? 2 : 3;
  return report;
}
