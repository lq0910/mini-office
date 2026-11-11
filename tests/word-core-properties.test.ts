/** 功能：验证原生统计、跨部件能力检测及真实工作区只读限制；时间：2026-11-11；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { loadOoxmlPackage } from '@mini-office/ooxml';
import { WordCoreSession, buildBlankDocx } from '../packages/word-core/src/index';
import { coreWordFormatHandler } from '../packages/editor-word/src/core-handler';
import { analyzeCoreWordCapabilities, coreMissingFonts } from '../packages/editor-word/src/core-capabilities';
import { insertWordTable } from '../packages/word-core/src/editing/objects';
const cores: WordCoreSession[] = [];
beforeEach(() => {
  vi.stubGlobal('requestAnimationFrame', (fn: FrameRequestCallback) => setTimeout(() => fn(performance.now()), 0));
  vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id));
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
});
afterEach(() => { for (const core of cores.splice(0)) core.dispose(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });
/** 功能：在真实空白包里添加测试部件，避免 mock 核心保护或扫描结果；时间：2026-11-11；作者：lq。 */
async function fixture(parts: Record<string, string>) {
  const zip = await JSZip.loadAsync(await buildBlankDocx());
  for (const [name, content] of Object.entries(parts)) zip.file(name, content);
  return zip.generateAsync({ type: 'uint8array' });
}
describe('M3 属性与兼容性', () => {
  it('统计从原生文字、单元格及空段读取 Unicode 字符，读取不污染历史', async () => {
    const core = await WordCoreSession.create(); cores.push(core);
    core.editor.commands.insertContent('中文 alpha beta 😀');
    core.editor.commands.insertContent({ type: 'docParagraph' });
    expect(insertWordTable(core.editor, 2, 2)).toBe(true);
    core.editor.commands.insertContent('格内');
    const before = core.state(); const stats = core.statistics();
    expect(stats.words).toBe(7); expect(stats.characters).toBe(14); expect(stats.paragraphs).toBeGreaterThanOrEqual(5);
    expect(core.state()).toEqual(before);
    const saved = await WordCoreSession.open((await core.prepareSave()).bytes); cores.push(saved);
    expect(saved.statistics()).toEqual(stats);
    core.editor.commands.undo(); expect(core.statistics().characters).toBe(12);
  });
  it('检测所有页眉页脚和注释区域使用的字体，系统别名及嵌入字体不会误报', async () => {
    const bytes = await fixture({ 'word/headerExtra.xml': '<w:hdr><w:p><w:r><w:rPr><w:rFonts w:ascii="Header Missing" w:eastAsia="已安装"/></w:rPr><w:t>页眉</w:t></w:r></w:p></w:hdr>',
      'word/footnotes.xml': '<w:footnotes><w:footnote><w:p><w:r><w:rPr><w:rFonts w:ascii="Embedded"/></w:rPr></w:r></w:p></w:footnote></w:footnotes>',
      'word/fontTable.xml': '<w:fonts><w:font w:name="Embedded"><w:embedRegular r:id="f1"/></w:font><w:font w:name="未使用字体"/></w:fonts>' });
    const pkg = await loadOoxmlPackage(bytes);
    const report = await analyzeCoreWordCapabilities(pkg, ['已安装', 'Calibri', '宋体', 'Arial']); pkg.dispose();
    expect(report.missingFonts).toContain('Header Missing'); expect(report.missingFonts).not.toContain('Embedded'); expect(report.missingFonts).not.toContain('未使用字体');
    expect(report.warnings.some(warning => warning.code === 'fonts.missing')).toBe(true);
    expect(coreMissingFonts(['Songti_SC', 'Lost', 'serif', 'Embed'], ['songti sc'], ['Embed'])).toEqual(['Lost']);
  });
  it('宏和数字签名通过真实 handler 打开后事务层只读并拒绝生成保存快照', async () => {
    for (const [path, code] of [['word/vbaProject.bin', 'docx.macro'], ['_xmlsignatures/sig1.xml', 'docx.signature']]) {
      const bytes = await fixture({ [path]: '<payload/>' });
      const opened = await coreWordFormatHandler.open({ name: '限制.docx', bytes });
      const core = opened.wordCore as WordCoreSession; cores.push(core);
      expect(opened.readonly).toBe(true); expect(opened.capability.level).toBe(1);
      expect(opened.capability.warnings).toContainEqual(expect.objectContaining({ code, severity: 'blocking' }));
      const doc = core.editor.state.doc;
      core.editor.view.dispatch(core.editor.state.tr.insertText('不能绕过'));
      expect(core.editor.state.doc).toBe(doc); expect(core.state().dirty).toBe(false);
      await expect(core.prepareSave()).rejects.toThrow('只读');
      expect(core.parsed.internal.originalBytes).toEqual(bytes);
    }
  });
  it('编辑保护由包检测和核心共同执行，不把保护件标为安全可编辑', async () => {
    const bytes = await fixture({ 'word/settings.xml': '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>' });
    const opened = await coreWordFormatHandler.open({ name: '保护.docx', bytes }); cores.push(opened.wordCore as WordCoreSession);
    expect(opened.readonly).toBe(true); expect(opened.capability.features).toContainEqual(expect.objectContaining({ code: 'docx.protection', severity: 'blocking' }));
  });
  it('复杂特性按真实产品范围报告 L2，普通表格和图片不继承旧解析器降级', async () => {
    const bytes = await fixture({ 'word/comments.xml': '<w:comments/>', 'word/embeddings/object.bin': 'data',
      'word/header1.xml': '<w:hdr><w:p><w:ins w:id="1"><w:r><w:t>修订</w:t></w:r></w:ins><w:sdt/></w:p></w:hdr>' });
    const opened = await coreWordFormatHandler.open({ name: '复杂.docx', bytes }); cores.push(opened.wordCore as WordCoreSession);
    expect(opened.readonly).toBe(false); expect(opened.capability.level).toBe(2);
    for (const code of ['docx.revisions', 'docx.comments', 'docx.embedded', 'docx.contentControls']) expect(opened.capability.warnings).toContainEqual(expect.objectContaining({ code, severity: 'warning' }));
    const core = await WordCoreSession.create(); cores.push(core); insertWordTable(core.editor, 2, 2);
    const pkg = await loadOoxmlPackage((await core.prepareSave()).bytes); const report = await analyzeCoreWordCapabilities(pkg); pkg.dispose();
    expect(report.level).toBe(3); expect(report.features).toContainEqual(expect.objectContaining({ code: 'docx.tables', supported: true }));
  });
});
