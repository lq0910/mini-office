/** 功能：以真实核心和 DOCX 回读验证页面设置、分节继承、动态域及历史版本隔离；时间：2026-11-10；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { WordCoreSession, buildBlankDocx } from '../packages/word-core/src/index';
import { patchWordSection, type WordPageLayout } from '../packages/word-core/src/editing/page-layout';
import { withHfEdit, resolveHf, withHfLink, hfLinked } from '../packages/word-core/src/upstream/apps/docs/src/renderer/hf-sections';
import { WordHeaderFooterEditor } from '../packages/word-core/src/editing/header-footer';
import { hfCellParaStyle } from '../packages/word-core/src/upstream/apps/docs/src/renderer/editor/hf-dom';
import { setActiveSubEditor } from '../packages/word-core/src/upstream/apps/docs/src/renderer/editor/active-editor';
// 功能：覆盖区域样式目录与正文历史的统一提交；时间：2026-11-11；作者：lq。
import { applyWordStyle, defineWordStyle, renameWordStyle, wordStyles, wordStyleCatalogOf } from '../packages/word-core/src/editing/styles';
import { readWordList, changeWordListLevel } from '../packages/word-core/src/editing/lists';
import { applyWordRuler } from '../packages/word-core/src/editing/ruler';
import { captureWordFormat } from '../packages/word-core/src/editing/format-painter';
import { insertWordPictures } from '../packages/word-core/src/editing/objects';
import { coreMakeHeaderFooter, coreHeaderFooterText } from '../packages/editor-word/src/WordCoreHeaderFooterDialog';

const sessions: WordCoreSession[] = [];
/** 功能：可复现的两节共享页脚、独立首页和未知节属性样本；时间：2026-11-10；作者：lq。 */
async function open(readonly = false) {
  const zip = await JSZip.loadAsync(await buildBlankDocx());
  const hfRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', `<Relationship Id="rIdFooter" Type="${hfRel}footer" Target="footer1.xml"/><Relationship Id="rIdFirst" Type="${hfRel}footer" Target="footerFirst.xml"/></Relationships>`));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/footerFirst.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'));
  zip.file('word/footer1.xml', '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Shared footer #</w:t></w:r></w:p></w:ftr>');
  zip.file('word/footerFirst.xml', '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>FIRST preserved</w:t></w:r></w:p></w:ftr>');
  const first = '<w:sectPr w:rsidSect="12345678"><w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="first" r:id="rIdFirst"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1200" w:right="1300" w:bottom="1400" w:left="1500" w:header="300" w:footer="400"/><w:cols w:num="2" w:space="720"/><w:titlePg/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>';
  const last = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/><w:pgNumType w:fmt="decimal"/><w:cols w:num="1"/></w:sectPr>';
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:pPr>${first}</w:pPr><w:r><w:t>FIRST body</w:t></w:r></w:p><w:p><w:r><w:t>SECOND body</w:t></w:r></w:p>${last}</w:body></w:document>`);
  const core = await WordCoreSession.open(await zip.generateAsync({ type: 'uint8array' }), { readonly });
  sessions.push(core); return core;
}
/** 功能：仅修改目标分节字段，不重建其他来源设置；时间：2026-11-10；作者：lq。 */
function page(core: WordCoreSession, index: number, patch: WordPageLayout['sections'][string]) {
  const catalog = core.pageLayout(); const id = core.sections()[index].lastBlockIndex;
  return core.commitPageLayout({ ...catalog, sections: { ...catalog.sections, [id]: { ...catalog.sections[id], ...patch } } });
}
/** 功能：页眉页脚原生继承与保存路径一致；时间：2026-11-10；作者：lq。 */
function footer(core: WordCoreSession, section: number, text: string) {
  const next = withHfEdit(core.headerFooterState(), section, 'footer', 'default', coreMakeHeaderFooter(text, 'right', 24))!;
  return core.commitPageLayout({ ...core.pageLayout(), ...next });
}
/** 功能：测试读回真实保存包，避免只验证内存目录；时间：2026-11-10；作者：lq。 */
async function reopen(core: WordCoreSession) { const next = await WordCoreSession.open((await core.prepareSave()).bytes); sessions.push(next); return next; }
beforeEach(() => {
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => setTimeout(() => callback(performance.now()), 0));
  vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id));
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
});
afterEach(() => { for (const core of sessions.splice(0)) core.dispose(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });

describe('M3 页面设置和页眉页脚', () => {
  /** 功能：完整格内格式保存重开及纸面显示，保留未改属性、邻格和共享历史；时间：2026-11-11；作者：lq。 */
  it('页脚布局表格格内完整段落格式保存重开，未知属性和邻格保持', async () => {
    const source = await open(), zip = await JSZip.loadAsync((await source.prepareSave()).bytes);
    const original = '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:ind w:left="0" w:leftChars="0" w:custom="preserved"/><w:suppressOverlap/></w:pPr><w:r><w:t>目标格</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p><w:r><w:t>邻格原文</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:ftr>';
    zip.file('word/footer1.xml', original); const core = await WordCoreSession.open(await zip.generateAsync({ type: 'uint8array' })); sessions.push(core);
    const region = new WordHeaderFooterEditor(core, document.createElement('div'), { section: 0, kind: 'footer', variant: 'default' });
    try {
      region.editor.commands.setTextSelection(3);
      region.editor.commands.updateAttributes('docParagraph', { indentRight: 360, keepNext: true, keepLines: true, widowControl: false, shadingFill: 'ddeeff', borders: 'b', borderLines: JSON.stringify({ b: { val: 'double', color: '006699', szPt: 1.5 } }) });
      const reopened = await reopen(core), value = resolveHf(reopened.headerFooterState(), 0, 'footer', 'default').value!;
      expect(value.paras![0].cells![0].paraProps![0]).toMatchObject({ indentRight: 360, keepNext: true, keepLines: true, widowControl: false, shadingFill: 'ddeeff', borderLines: { b: { val: 'double' } } });
      expect(value.paras![0].cells![1].paraProps?.[0]?.indentRight).not.toBe(360);
      const saved = await JSZip.loadAsync((await core.prepareSave()).bytes); const xml = await saved.file('word/footer1.xml')!.async('string'); expect(xml).toContain('<w:suppressOverlap/>');
      expect(xml).toContain('<w:spacing w:before="0" w:after="0"/>'); expect(xml).toContain('w:leftChars="0" w:custom="preserved"');
      expect(hfCellParaStyle(value.paras![0].cells![0].paraProps![0])).toMatchObject({ marginInlineEnd: '24.0px', backgroundColor: '#ddeeff', borderBottom: '2px double #006699' });
      core.editor.commands.undo(); await expect((await JSZip.loadAsync((await core.prepareSave()).bytes)).file('word/footer1.xml')!.async('string')).resolves.toBe(original);
    } finally { region.dispose(); }
  });
  /** 功能：区域列表编号由主会话分配，保存编号资源和原生身份，撤销/重做与正文隔离；时间：2026-11-11；作者：lq。 */
  it('页脚编号、层级和取消列表均共享主历史并能保存重开', async () => {
    const core = await open();
    const region = new WordHeaderFooterEditor(core, document.createElement('div'), { section: 1, kind: 'footer', variant: 'default' });
    try {
      setActiveSubEditor(region.editor);
      expect(core.toggleList('ordered')).toBe(true);
      expect(readWordList(region.editor)).toMatchObject({ kind: 'ordered', level: 0 });
      expect(changeWordListLevel(region.editor, 1)).toBe(true);
      expect(readWordList(region.editor).level).toBe(1);
      const saved = await reopen(core);
      const marker = resolveHf(saved.headerFooterState(), 1, 'footer', 'default').value?.paras?.[0].listMarker;
      expect(marker).toMatchObject({ kind: 'ordered', ilvl: 1, text: '1.' });
      expect(saved.parsed.numbering.has(marker!.numId!)).toBe(true);
      expect(core.editor.getText()).toBe('FIRST body\n\nSECOND body');
      expect(core.toggleList('ordered', region.editor)).toBe(true);
      expect(readWordList(region.editor).kind).toBeUndefined();
      core.editor.commands.undo(); expect(readWordList(region.editor).level).toBe(1);
      core.editor.commands.undo(); expect(readWordList(region.editor).level).toBe(0);
      core.editor.commands.undo(); expect(readWordList(region.editor).kind).toBeUndefined();
      const undone = await reopen(core);
      expect(resolveHf(undone.headerFooterState(), 1, 'footer', 'default').value?.paras?.[0].listMarker).toBeUndefined();
      expect([...undone.parsed.numbering.keys()]).toEqual([...core.parsed.numbering.keys()]);
    } finally { region.dispose(); }
  });
  /** 功能：区域组合输入即使失去活动焦点也拒绝保存，完成后正常保存；时间：2026-11-11；作者：lq。 */
  it('保存检查全部已登记区域的组合输入并在退出后恢复', async () => {
    const core = await open();
    const region = new WordHeaderFooterEditor(core, document.createElement('div'), { section: 1, kind: 'footer', variant: 'default' });
    try {
      Object.defineProperty(region.editor.view, 'composing', { configurable: true, get: () => true });
      setActiveSubEditor(null);
      await expect(core.prepareSave()).rejects.toThrow('中文输入尚未确认');
      Object.defineProperty(region.editor.view, 'composing', { configurable: true, get: () => false });
      await expect(core.prepareSave()).resolves.toHaveProperty('bytes');
      expect(core.state().dirty).toBe(false);
    } finally { region.dispose(); }
    await expect(core.prepareSave()).resolves.toHaveProperty('bytes');
  });
  /** 功能：子区域创建/套用和目录修改均一次进入主历史，重开保留定义和部件引用；时间：2026-11-11；作者：lq。 */
  it('页脚创建样式与内容原子撤销、正文复用且单独改名不重写原部件', async () => {
    const core = await open();
    const region = new WordHeaderFooterEditor(core, document.createElement('div'), { section: 1, kind: 'footer', variant: 'default' });
    try {
      const snapshot = captureWordFormat(core.editor, core.formattingSource());
      snapshot.text.bold = true; snapshot.text.sizeHalfPoints = 36; snapshot.paragraph.align = 'right';
      const definition = defineWordStyle('RegionStyle', '区域样式', snapshot);
      expect(applyWordStyle(region.editor, core.parsed, 'RegionStyle', definition)).toBe(true);
      expect(wordStyles(core.editor.state.doc, core.parsed).has('RegionStyle')).toBe(true);
      expect(region.editor.state.doc.firstChild!.attrs.styleId).toBe('RegionStyle');
      const originalIndent = region.editor.state.doc.firstChild!.attrs.indentLeft;
      expect(applyWordRuler(region.editor, core.formattingSource(), { indentLeft: 720 })).toBe(true);
      region.editor.commands.undo();
      expect(region.editor.state.doc.firstChild!.attrs.indentLeft).toBe(originalIndent);
      expect(region.editor.state.doc.firstChild!.attrs.styleId).toBe('RegionStyle');
      core.editor.commands.undo();
      expect(wordStyles(core.editor.state.doc, core.parsed).has('RegionStyle')).toBe(false);
      expect(region.editor.state.doc.firstChild!.attrs.styleId).not.toBe('RegionStyle');
      region.editor.commands.redo();
      expect(wordStyleCatalogOf(region.editor.state.doc)).toEqual(wordStyleCatalogOf(core.editor.state.doc));
      expect(applyWordStyle(core.editor, core.parsed, 'RegionStyle')).toBe(true);
      const saved = await reopen(core);
      expect(saved.editor.state.doc.firstChild!.attrs.styleId).toBe('RegionStyle');
      expect(saved.parsed.styles.get('RegionStyle')?.display).toMatchObject({ bold: true, sizeHalfPoints: 36, align: 'right' });
      expect(resolveHf(saved.headerFooterState(), 1, 'footer', 'default').value?.paras?.[0].styleId).toBe('RegionStyle');
      const before = JSON.stringify(core.pageLayout());
      expect(renameWordStyle(region.editor, core.parsed, 'RegionStyle', '区域改名')).toBe(true);
      expect(JSON.stringify(core.pageLayout())).toBe(before);
      expect(wordStyles(core.editor.state.doc, core.parsed).get('RegionStyle')?.name).toBe('区域改名');
      core.editor.commands.undo();
      expect(wordStyles(core.editor.state.doc, core.parsed).get('RegionStyle')?.name).toBe('区域样式');
      expect(region.editor.state.doc.firstChild!.attrs.styleId).toBe('RegionStyle');
    } finally { region.dispose(); }
  });

  /** 功能：变体设置跟随各节及文档范围，关闭重开保留原首页部件；时间：2026-11-11；作者：lq。 */
  it('首页按节、奇偶按文档保存并共享历史，不删除隐藏变体', async () => {
    const core = await open();
    page(core, 0, { titlePg: false }); page(core, 1, { titlePg: true });
    core.commitPageLayout({ ...core.pageLayout(), evenAndOddHeaders: true });
    expect(core.sections().map(section => section.titlePg)).toEqual([false, true]);
    expect(core.evenAndOddHeaders()).toBe(true);
    const saved = await reopen(core);
    expect(saved.sections().map(section => section.titlePg)).toEqual([false, true]);
    expect(saved.evenAndOddHeaders()).toBe(true);
    const zip = await JSZip.loadAsync((await core.prepareSave()).bytes);
    expect(await zip.file('word/footerFirst.xml')!.async('string')).toContain('FIRST preserved');
    core.editor.commands.undo(); expect(core.evenAndOddHeaders()).toBe(false);
    core.editor.commands.undo(); expect(core.sections()[1].titlePg).toBe(false);
    core.editor.commands.undo(); expect(core.sections()[0].titlePg).toBe(true);
  });
  /** 功能：解除继承必须生成独立部件，原始所有者、隐藏变体和原始字符属性均不受影响；时间：2026-11-11；作者：lq。 */
  it('解除继承克隆原部件，编辑独立副本和再次链接都可撤销保存重开', async () => {
    const core = await open();
    core.commitPageLayout({ ...core.pageLayout(), ...withHfLink(core.headerFooterState(), 1, 'footer', 'default', false)! });
    expect(hfLinked(core.headerFooterState(), 1, 'footer', 'default')).toBe(false);
    let saved = await reopen(core);
    expect(saved.sections()[1].footerRefs.default).not.toBe(saved.sections()[0].footerRefs.default);
    expect(resolveHf(saved.headerFooterState(), 1, 'footer', 'default').value?.text).toBe('Shared footer #');
    footer(core, 1, 'Independent {SECTIONPAGES}');
    expect(resolveHf(core.headerFooterState(), 0, 'footer', 'default').value?.text).toBe('Shared footer #');
    saved = await reopen(core);
    expect(coreHeaderFooterText(resolveHf(saved.headerFooterState(), 1, 'footer', 'default').value)).toBe('Independent {SECTIONPAGES}');
    expect(resolveHf(saved.headerFooterState(), 0, 'footer', 'default').value?.text).toBe('Shared footer #');
    core.commitPageLayout({ ...core.pageLayout(), ...withHfLink(core.headerFooterState(), 1, 'footer', 'default', true)! });
    saved = await reopen(core);
    expect(saved.sections()[1].footerRefs.default).toBeUndefined();
    expect(resolveHf(saved.headerFooterState(), 1, 'footer', 'default').value?.text).toBe('Shared footer #');
    core.editor.commands.undo(); saved = await reopen(core);
    expect(coreHeaderFooterText(resolveHf(saved.headerFooterState(), 1, 'footer', 'default').value)).toBe('Independent {SECTIONPAGES}');
  });
  /** 功能：区域是主历史的实时投影，保持打开时保存、主撤销及区域重做不得丢失或重复记录；时间：2026-11-11；作者：lq。 */
  it('富文本区域直接输入和格式命令即时保存，只有主历史且与其他会话隔离', async () => {
    const core = await open(); const other = await open();
    const region = new WordHeaderFooterEditor(core, document.createElement('div'), { section: 1, kind: 'footer', variant: 'default' });
    try {
      region.editor.commands.focus('end');
      region.editor.commands.setTextSelection(region.editor.state.doc.content.size - 1); setActiveSubEditor(region.editor);
      region.editor.commands.insertContent('中文');
      region.editor.commands.insertContent('续写');
      expect(core.commandEditor() === region.editor).toBe(true); expect(other.commandEditor() === other.editor).toBe(true);
      expect(coreHeaderFooterText(resolveHf(core.headerFooterState(), 0, 'footer', 'default').value)).toContain('中文续写');
      const saved = await reopen(core);
      expect(coreHeaderFooterText(resolveHf(saved.headerFooterState(), 1, 'footer', 'default').value)).toContain('中文续写');
      core.editor.commands.undo();
      expect(region.editor.getText()).toBe('Shared footer #');
      region.editor.commands.redo(); expect(region.editor.getText()).toContain('中文续写');
      region.editor.commands.undo(); expect(region.editor.getText()).toBe('Shared footer #');
      expect(core.editor.getText()).toBe('FIRST body\n\nSECOND body');
    } finally { region.dispose(); }
  });
  /** 功能：图片及包内/外关系在解除继承时完整复制，后续行内新图和原图可共同保存；时间：2026-11-11；作者：lq。 */
  it('独立页脚克隆子目录图片关系，保留来源字节并支持区域新插图', async () => {
    const original = await open(); const zip = await JSZip.loadAsync((await original.prepareSave()).bytes);
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jBf8AAAAASUVORK5CYII=';
    zip.file('word/media/footer.png', png, { base64: true });
    const xml = '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:p><w:r><w:drawing><wp:inline><wp:extent cx="190500" cy="190500"/><wp:docPr id="73" name="原图"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill><pic:spPr><a:xfrm><a:ext cx="190500" cy="190500"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t>图旁原文</w:t></w:r></w:p></w:ftr>';
    zip.file('word/parts/footer.xml', xml); zip.remove('word/footer1.xml');
    zip.file('word/_rels/document.xml.rels', (await zip.file('word/_rels/document.xml.rels')!.async('string')).replace('Target="footer1.xml"', 'Target="parts/footer.xml"'));
    zip.file('word/parts/_rels/footer.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/footer.png"/></Relationships>');
    const core = await WordCoreSession.open(await zip.generateAsync({ type: 'uint8array' })); sessions.push(core);
    core.commitPageLayout({ ...core.pageLayout(), ...withHfLink(core.headerFooterState(), 1, 'footer', 'default', false)! });
    const untouched = await JSZip.loadAsync((await core.prepareSave()).bytes);
    const copyPath = Object.keys(untouched.files).find(path => /^word\/footer\d+\.xml$/.test(path))!;
    expect(await untouched.file(copyPath)!.async('string')).toBe(xml);
    expect(await untouched.file(copyPath.replace('word/', 'word/_rels/') + '.rels')!.async('string')).toContain('Target="/word/media/footer.png"');
    const region = new WordHeaderFooterEditor(core, document.createElement('div'), { section: 1, kind: 'footer', variant: 'default' });
    try {
      region.editor.commands.setTextSelection(region.editor.state.doc.content.size - 1);
      expect(insertWordPictures(region.editor, [{ base64: png, dataUrl: `data:image/png;base64,${png}`, mime: 'image/png', widthPx: 10, heightPx: 10, label: '新图' }])).toBe(true);
      const saved = await JSZip.loadAsync((await core.prepareSave()).bytes);
      expect(await saved.file('word/parts/footer.xml')!.async('string')).toBe(xml);
      expect(await saved.file(copyPath)!.async('string')).toContain('图旁原文');
      const reopened = await reopen(core); expect(resolveHf(reopened.headerFooterState(), 1, 'footer', 'default').value?.paras?.[0].runs.filter(run => run.image).length).toBe(2);
    } finally { region.dispose(); }
  });
  it('只修改目标节的方向尺寸及一侧边距，保留原栏、网格、引用和未知 XML', async () => {
    const core = await open(); const original = core.sections();
    page(core, 0, { pageWidth: 16838, pageHeight: 11906, orientation: 'landscape', marginLeft: 900 });
    const saved = await reopen(core);
    expect(saved.sections()[0].settings).toMatchObject({ pageWidth: 16838, pageHeight: 11906, orientation: 'landscape', marginLeft: 900, columns: 2, marginTop: 1200, headerDist: 300 });
    expect(saved.sections()[0].sectPrXml).toContain('w:rsidSect="12345678"');
    expect(saved.sections()[0].sectPrXml).toContain('<w:docGrid w:type="lines" w:linePitch="360"/>');
    expect(saved.sections()[1].sectPrXml).toBe(original[1].sectPrXml);
    expect(saved.editor.state.doc.textContent).toBe('FIRST bodySECOND body');
    expect(core.parsed.blocks.map(b => b.originalXml)).not.toEqual(saved.parsed.blocks.map(b => b.originalXml));
    expect(core.sections()[0].headerRefs).toEqual(original[0].headerRefs);
  });
  it('修改带分节的正文后，生成段落仍保存目标页面设置，尾节设置互不覆盖', async () => {
    const core = await open(); core.editor.commands.insertContent('EDIT ');
    page(core, 0, { marginRight: 960, headerDist: 567 });
    page(core, 1, { pageWidth: 16838, pageHeight: 11906, orientation: 'landscape', footerDist: 280, pageNumberStart: 5 });
    const saved = await reopen(core);
    expect(saved.editor.getText()).toContain('EDIT');
    expect(saved.sections()[0].settings).toMatchObject({ marginRight: 960, headerDist: 567 });
    expect(saved.sections()[1].settings).toMatchObject({ pageWidth: 16838, pageHeight: 11906, orientation: 'landscape', footerDist: 280 });
    expect(saved.sections()[1].pageNumberStart).toBe(5);
  });
  it('正文、页面和页脚共享历史，逐步撤销重做与保存保持一致', async () => {
    const core = await open(); const before = core.editor.getJSON();
    core.editor.commands.insertContent('EDIT');
    page(core, 0, { marginLeft: 800 }); footer(core, 0, '第 {PAGE} 页 / 共 {NUMPAGES} 页 #');
    core.editor.commands.undo(); expect(coreHeaderFooterText(resolveHf(core.headerFooterState(), 0, 'footer', 'default').value)).toBe('Shared footer #');
    core.editor.commands.undo(); expect(core.sections()[0].settings.marginLeft).toBe(1500);
    core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before);
    core.editor.commands.redo(); core.editor.commands.redo(); core.editor.commands.redo();
    const saved = await reopen(core); expect(saved.sections()[0].settings.marginLeft).toBe(800);
    expect(coreHeaderFooterText(resolveHf(saved.headerFooterState(), 0, 'footer', 'default').value)).toBe('第 {PAGE} 页 / 共 {NUMPAGES} 页 #');
  });
  it('继承页脚编辑写回所有者，首页部件保持字节，动态域和多行对齐持久化', async () => {
    const core = await open(); expect(resolveHf(core.headerFooterState(), 1, 'footer', 'default').owner).toBe(0);
    footer(core, 1, '合同 #\n第 {PAGE} 页 / 共 {NUMPAGES} 页');
    const prepared = await core.prepareSave(); const zip = await JSZip.loadAsync(prepared.bytes);
    expect(await zip.file('word/footerFirst.xml')!.async('string')).toContain('FIRST preserved');
    expect(await zip.file('word/footerFirst.xml')!.async('string')).toBe(await (await JSZip.loadAsync(core.parsed.internal.originalBytes)).file('word/footerFirst.xml')!.async('string'));
    const saved = await reopen(core);
    for (const index of [0,1]) {
      const value = resolveHf(saved.headerFooterState(), index, 'footer', 'default').value!;
      expect(coreHeaderFooterText(value)).toBe('合同 #\n第 {PAGE} 页 / 共 {NUMPAGES} 页');
      expect(value.paras).toHaveLength(2); expect(value.paras![0]).toMatchObject({ align: 'right', runs: [expect.objectContaining({ sizeHalfPoints: 24 })] });
    }
  });
  it('保存捕获页面和页脚同一版本，保存期间继续改动不会污染已签发字节', async () => {
    const core = await open(); page(core, 1, { marginLeft: 800 }); footer(core, 0, '旧 {PAGE}');
    const saving = core.prepareSave(); page(core, 1, { marginLeft: 1000 }); footer(core, 0, '新 {PAGE}');
    const prepared = await saving; core.acknowledgeSave(prepared); expect(core.state().dirty).toBe(true);
    const captured = await WordCoreSession.open(prepared.bytes); sessions.push(captured);
    expect(captured.sections()[1].settings.marginLeft).toBe(800);
    expect(coreHeaderFooterText(resolveHf(captured.headerFooterState(), 0, 'footer', 'default').value)).toBe('旧 {PAGE}');
    expect(core.sections()[1].settings.marginLeft).toBe(1000);
  });
  it('拒绝无效页面尺寸、起始页码、只读与组合输入，目录和历史不变', async () => {
    const core = await open(); const original = core.state().revision;
    for (const patch of [{ marginLeft: NaN }, { pageWidth: 0 }, { marginTop: 16000 }, { pageNumberStart: -1 }, { footerDist: 17000 }]) expect(() => page(core, 0, patch)).toThrow();
    expect(core.state().revision).toBe(original);
    Object.defineProperty(core.editor.view, 'composing', { configurable: true, get: () => true });
    expect(page(core, 0, { marginLeft: 800 })).toBe(false); expect(core.state().revision).toBe(original);
    const readonly = await open(true); expect(() => page(readonly, 0, { marginLeft: 800 })).toThrow('只读'); expect(readonly.state().dirty).toBe(false);
  });
  it('单节空白文档可新增页眉页码、距边及清空，撤销后来源字节完整', async () => {
    const core = await WordCoreSession.create(); sessions.push(core);
    const before = await core.prepareSave();
    page(core, 0, { headerDist: 285, footerDist: 570 });
    const edits = withHfEdit(core.headerFooterState(), 0, 'header', 'default', coreMakeHeaderFooter('页眉 {PAGE} / {NUMPAGES}', 'center', 21))!;
    core.commitPageLayout({ ...core.pageLayout(), ...edits });
    const saved = await reopen(core); expect(saved.sections()[0].settings).toMatchObject({ headerDist: 285, footerDist: 570 });
    expect(coreHeaderFooterText(resolveHf(saved.headerFooterState(), 0, 'header', 'default').value)).toBe('页眉 {PAGE} / {NUMPAGES}');
    const clear = withHfEdit(core.headerFooterState(), 0, 'header', 'default', coreMakeHeaderFooter('', 'center', 21))!;
    core.commitPageLayout({ ...core.pageLayout(), ...clear });
    const cleared = await reopen(core); expect(coreHeaderFooterText(resolveHf(cleared.headerFooterState(), 0, 'header', 'default').value)).toBe('');
    core.editor.commands.undo(); core.editor.commands.undo(); core.editor.commands.undo();
    expect((await core.prepareSave()).bytes).toEqual(before.bytes);
  });
  it('装订线和固定负边距保存回读保持等效尺寸，拒绝小于装订线的值', async () => {
    const core = await open();
    const zip = await JSZip.loadAsync(core.parsed.internal.originalBytes);
    const original = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', original.replace('w:top="1200"', 'w:top="-1200"').replace('w:left="1500"', 'w:left="1500" w:gutter="240"'));
    const input = await WordCoreSession.open(await zip.generateAsync({ type: 'uint8array' })); sessions.push(input);
    expect(input.sections()[0].settings).toMatchObject({ marginLeft: 1740, marginTopFixed: true });
    page(input, 0, { marginLeft: 1000, marginTop: 900, marginBottom: 800 });
    const saved = await reopen(input);
    expect(saved.sections()[0].settings).toMatchObject({ marginLeft: 1000, marginTop: 900, marginTopFixed: true, marginBottom: 800, gutter: 240 });
    expect(saved.sections()[0].sectPrXml).toContain('w:left="760"');
    expect(saved.sections()[0].sectPrXml).toContain('w:top="-900"');
    expect(() => page(input, 0, { marginLeft: 100 })).toThrow('装订线');
  });
  it('过长页眉页脚在提交前拒绝，避免覆盖正文并破坏历史', async () => {
    const core = await open(); const before = core.editor.getJSON();
    expect(() => footer(core, 0, '超长页脚\n'.repeat(150))).toThrow('占用空间过大');
    expect(core.editor.getJSON()).toEqual(before);
  });
  it('补写缺失标签保持 CT_SectPr 顺序、单引号未知属性与原页码格式', () => {
    const xml = patchWordSection('<w:sectPr><w:headerReference r:id="x"/><w:pgMar w:left=\'120\' w:gutter=\'50\'/><w:pgNumType w:fmt="lowerRoman"/><w:cols/></w:sectPr>', { pageWidth: 11906, pageHeight: 16838, marginLeft: 800, pageNumberStart: 5 });
    expect(xml.indexOf('<w:headerReference')).toBeLessThan(xml.indexOf('<w:pgSz'));
    expect(xml.indexOf('<w:pgSz')).toBeLessThan(xml.indexOf('<w:pgMar'));
    expect(xml).toContain('w:gutter=\'50\''); expect(xml).not.toContain("w:left='120'"); expect(xml).toContain('w:fmt="lowerRoman"');
  });
});
