/** 功能：以真实核心事务和 DOCX 重开验证样式定义历史、未知 XML 与格式刷；时间：2026-11-10；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { PPR_CHILD_ORDER, RPR_CHILD_ORDER, splitXmlChildren } from '../packages/word-core/src/upstream/packages/docx-engine/src/generate';
import { WordCoreSession, buildBlankDocx } from '../packages/word-core/src/index';
import { applyWordStyle, applyWordFormatPainter, captureWordFormat, defineWordStyle, deleteWordStyle, renameWordStyle, importWordStyles, wordStyles, toggleWordInlineFormat, CellSelection, type WordFormatSnapshot } from '../packages/word-core/src/editing/index';
import { coreToolbarFormatting } from '../packages/editor-word/src/core-formatting';
import { coreStyleDefinition } from '../packages/editor-word/src/WordCoreStyles';
import { BUILTIN_STYLES } from '../packages/editor-word/src/styles';

const sessions: WordCoreSession[] = [];
/** 功能：保留导入样式未知属性、字符语言和超链接的可复现 DOCX；时间：2026-11-10；作者：lq。 */
async function open(readonly = false, body?: string) {
  const zip = await JSZip.loadAsync(await buildBlankDocx());
  let styles = await zip.file('word/styles.xml')!.async('string');
  styles = styles.replace('</w:styles>', '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Imported"><w:name w:val="来源样式"/><w:basedOn w:val="Normal"/><w:rsid w:val="12345678"/><w:pPr><w:keepNext/><w:tabs><w:tab w:val="right" w:pos="1440"/></w:tabs></w:pPr><w:rPr><w:b/><w:lang w:val="en-GB"/><w:sz w:val="30"/></w:rPr></w:style></w:styles>');
  zip.file('word/styles.xml', styles);
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
    <w:p><w:pPr><w:pStyle w:val="Imported"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="宋体"/><w:sz w:val="32"/><w:color w:val="D92D20"/><w:lang w:val="zh-CN"/></w:rPr><w:t>源文字</w:t></w:r></w:p>
    <w:p><w:hyperlink r:id="rId998"><w:r><w:rPr><w:i/><w:lang w:val="en-US"/><w:noProof/></w:rPr><w:t>Target</w:t></w:r></w:hyperlink></w:p><w:p><w:r><w:t>第三段</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  if (body) zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rId998" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/target" TargetMode="External"/></Relationships>'));
  const core = await WordCoreSession.open(await zip.generateAsync({ type: 'uint8array' }), { readonly });
  sessions.push(core); return core;
}
/** 功能：只读取当前生成的 DOCX，不用断言替代真实保存器；时间：2026-11-10；作者：lq。 */
async function xml(core: WordCoreSession, path = 'word/styles.xml') { return (await JSZip.loadAsync((await core.prepareSave()).bytes)).file(path)!.async('string'); }
/** 功能：选中第二段并保留第一、三段作为未目标对照；时间：2026-11-10；作者：lq。 */
function selectTarget(core: WordCoreSession) { const start = core.editor.state.doc.child(0).nodeSize + 1; core.editor.commands.setTextSelection({ from: start, to: start + 6 }); }
beforeEach(() => {
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => setTimeout(() => callback(performance.now()), 0));
  vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id));
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
});
afterEach(() => { for (const core of sessions.splice(0)) core.dispose(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });

describe('M3 样式与格式刷事务', () => {
  /** 功能：删除被不同部件引用的样式只重定向合法样式引用，保留原格式、普通同名文本和未知 XML；时间：2026-11-11；作者：lq。 */
  it('导入样式安全删除覆盖页眉脚注编号派生关系且撤销恢复原字节', async () => {
    const original = await open(); const zip = await JSZip.loadAsync((await original.prepareSave()).bytes);
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const part = `<w:hdr xmlns:w="${W}" xmlns:x="urn:unknown"><w:p><w:pPr><w:pStyle w:val="Imported"/><w:contextualSpacing/></w:pPr><w:r><w:t>Imported</w:t></w:r></w:p><x:style x:val="Imported"/></w:hdr>`;
    zip.file('word/header-dependency.xml', part);
    zip.file('word/styles.xml', (await zip.file('word/styles.xml')!.async('string')).replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="Dependent"><w:name w:val="继承"/><w:basedOn w:val="Imported"/><w:next w:val="Imported"/></w:style></w:styles>'));
    zip.file('word/footnotes.xml', `<w:footnotes xmlns:w="${W}"><w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="Imported"/></w:pPr><w:r><w:t>注释</w:t></w:r></w:p></w:footnote></w:footnotes>`);
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const core = await WordCoreSession.open(bytes); sessions.push(core);
    expect(deleteWordStyle(core.editor, core.parsed, 'Imported')).toBe(true);
    const alias = core.editor.state.doc.firstChild!.attrs.styleId;
    const saved = await JSZip.loadAsync((await core.prepareSave()).bytes);
    expect(await saved.file('word/header-dependency.xml')!.async('string')).toBe(part.replace('w:val="Imported"', `w:val="${alias}"`));
    expect(await saved.file('word/footnotes.xml')!.async('string')).toContain(`w:pStyle w:val="${alias}"`);
    expect(await saved.file('word/styles.xml')!.async('string')).toContain(`<w:basedOn w:val="${alias}"/><w:next w:val="${alias}"/>`);
    expect(deleteWordStyle(core.editor, core.parsed, 'Normal')).toBe(false);
    const reopened = await WordCoreSession.open((await core.prepareSave()).bytes); sessions.push(reopened);
    expect(reopened.parsed.styles.get('Dependent')?.display).toMatchObject({ bold: true, sizeHalfPoints: 30 });
    core.editor.commands.undo();
    const undone = (await core.prepareSave()).bytes;
    expect(undone.length === bytes.length && undone.every((value, index) => value === bytes[index])).toBe(true);
  });
  /** 功能：原内置引用/列表样式的边框和编号必须在正文、定义、撤销及重开保持一致；时间：2026-11-11；作者：lq。 */
  it('引用边框和列表预设通过原生定义/编号资源保存，重开后可再次应用', async () => {
    const core = await open(); selectTarget(core);
    const quote = coreStyleDefinition(BUILTIN_STYLES.find(item => item.id === 'IntenseQuote')!, core);
    expect(applyWordStyle(core.editor, core.parsed, 'IntenseQuote', quote)).toBe(true);
    expect(core.editor.state.doc.child(1).attrs.borders).toBe('tb');
    let reopened = await WordCoreSession.open((await core.prepareSave()).bytes); sessions.push(reopened);
    expect(reopened.parsed.styles.get('IntenseQuote')?.display?.borderSides).toMatchObject({ t: { color: '98A2B3', szPt: 1.125 }, b: { color: '98A2B3' } });
    selectTarget(core);
    const list = coreStyleDefinition(BUILTIN_STYLES.find(item => item.id === 'MoListNumber')!, core);
    expect(applyWordStyle(core.editor, core.parsed, 'MoListNumber', list)).toBe(true);
    expect(core.editor.state.doc.child(1).type.name).toBe('docListItem');
    const numId = core.editor.state.doc.child(1).attrs.numId;
    reopened = await WordCoreSession.open((await core.prepareSave()).bytes); sessions.push(reopened);
    expect(reopened.parsed.numbering.has(numId)).toBe(true);
    expect(reopened.parsed.styles.get('MoListNumber')?.numPr).toMatchObject({ numId, ilvl: 0 });
    selectTarget(reopened); expect(applyWordStyle(reopened.editor, reopened.parsed, 'MoListNumber')).toBe(true);
    expect(reopened.editor.state.doc.child(1).attrs.numId).toBe(numId);
    core.editor.commands.undo(); expect(core.editor.state.doc.child(1).type.name).toBe('docParagraph');
    expect(await xml(core, 'word/numbering.xml')).not.toContain(`w:numId="${numId}"`);
  });
  /** 功能：高级字符效果按已建模字段复制，目标语言/未知属性保留，段落双线和制表位一起持久化；时间：2026-11-11；作者：lq。 */
  it('格式刷保留双线边框、制表位、字符间距/上标/阴影和大写效果', async () => {
    const body = '<w:p><w:pPr><w:pBdr><w:bottom w:val="double" w:color="336699" w:sz="16"/></w:pBdr><w:tabs><w:tab w:val="right" w:pos="2000"/></w:tabs></w:pPr><w:r><w:rPr><w:caps/><w:spacing w:val="40"/><w:vertAlign w:val="superscript"/><w:shadow/><w:bdr w:val="single" w:sz="8" w:color="FF0000"/></w:rPr><w:t>源效果</w:t></w:r></w:p><w:p><w:r><w:rPr><w:lang w:val="en-US"/><w:noProof/></w:rPr><w:t>Target</w:t></w:r></w:p>';
    const core = await open(false, body); core.editor.commands.setTextSelection(1);
    const f = captureWordFormat(core.editor, core.formattingSource());
    expect(f.effects).toMatchObject({ charSpacingTwips: 40, vertAlign: 'superscript', caps: 'all', textEffect: 'shadow' });
    selectTarget(core); expect(applyWordFormatPainter(core.editor, f)).toBe(true);
    const saved = await xml(core, 'word/document.xml');
    const target = saved.match(/<w:p>(?:(?!<w:p>)[\s\S])*?Target[\s\S]*?<\/w:p>/)![0];
    for (const item of ['w:val="double"', 'w:val="2000"', 'w:val="superscript"', '<w:shadow', '<w:caps', 'w:val="en-US"', '<w:noProof']) {
      if (item === 'w:val="2000"') expect(target).toContain('w:pos="2000"'); else expect(target).toContain(item);
    }
    const reopened = await WordCoreSession.open((await core.prepareSave()).bytes); sessions.push(reopened);
    expect(JSON.parse(reopened.editor.state.doc.child(1).attrs.borderLines).b.val).toBe('double');
    expect(captureWordFormat(reopened.editor, reopened.formattingSource()).effects?.caps).toBe('all');
  });
  /** 功能：样式继承的下划线/删除线可局部显式关闭，保存重开和撤销不再次出现；时间：2026-11-11；作者：lq。 */
  it('继承下划线和删除线允许局部取消并保存显式关闭值', async () => {
    const core = await open(); selectTarget(core);
    const f = captureWordFormat(core.editor, core.formattingSource()); f.text.underline = true; f.text.strike = true;
    applyWordStyle(core.editor, core.parsed, 'Decorated', defineWordStyle('Decorated', '装饰文字', f));
    toggleWordInlineFormat(core.editor, core.formattingSource(), 'underline'); toggleWordInlineFormat(core.editor, core.formattingSource(), 'strike');
    expect(coreToolbarFormatting(core).activeStyle).toMatchObject({ underline: false, strike: false });
    const saved = await xml(core, 'word/document.xml'); expect(saved).toContain('<w:u w:val="none"/>'); expect(saved).toContain('<w:strike w:val="0"/>');
    const reopened = await WordCoreSession.open((await core.prepareSave()).bytes); sessions.push(reopened); selectTarget(reopened);
    expect(coreToolbarFormatting(reopened).activeStyle).toMatchObject({ underline: false, strike: false });
    core.editor.commands.undo(); expect(coreToolbarFormatting(core).activeStyle.strike).toBe(true);
    core.editor.commands.undo(); expect(coreToolbarFormatting(core).activeStyle.underline).toBe(true);
  });
  it('新建、应用、改名、删除均可逐步撤销；撤销定义后保存不泄露新样式', async () => {
    const core = await open();
    selectTarget(core);
    const original = core.editor.getJSON();
    const f = captureWordFormat(core.editor, core.formattingSource());
    expect(applyWordStyle(core.editor, core.parsed, 'MoTest', defineWordStyle('MoTest', '合同', f))).toBe(true);
    expect(core.state().dirty).toBe(true);
    expect(core.editor.state.doc.child(1).attrs.styleId).toBe('MoTest');
    expect(core.editor.commands.undo()).toBe(true);
    expect(core.editor.getJSON()).toEqual(original);
    expect(await xml(core)).not.toContain('MoTest');
    expect(core.editor.commands.redo()).toBe(true);
    expect(renameWordStyle(core.editor, core.parsed, 'MoTest', '合同正文')).toBe(true);
    expect(wordStyles(core.editor.state.doc, core.parsed).get('MoTest')!.name).toBe('合同正文');
    core.editor.commands.undo();
    expect(wordStyles(core.editor.state.doc, core.parsed).get('MoTest')!.name).toBe('合同');
    core.editor.commands.redo();
    const beforeDelete = core.editor.state.doc.child(1);
    expect(deleteWordStyle(core.editor, core.parsed, 'MoTest')).toBe(true);
    expect(core.editor.state.doc.child(1).attrs.styleId).toBeNull();
    expect(core.editor.state.doc.child(1).content.eq(beforeDelete.content)).toBe(true);
    expect(await xml(core)).not.toContain('MoTest');
    core.editor.commands.undo();
    expect(core.editor.state.doc.child(1).eq(beforeDelete)).toBe(true);
    expect(await xml(core)).toContain('合同正文');
  });

  it('仅导入定义的空正文事务可撤销，并隔离保存期间的重命名版本', async () => {
    const core = await WordCoreSession.create(); sessions.push(core);
    const f = captureWordFormat(core.editor, core.formattingSource());
    expect(importWordStyles(core.editor, [defineWordStyle('MoOne', '一', f), defineWordStyle('MoTwo', '二', f)])).toBe(true);
    const saving = core.prepareSave();
    renameWordStyle(core.editor, core.parsed, 'MoOne', '新一');
    const saved = await saving;
    const savedXml = await (await JSZip.loadAsync(saved.bytes)).file('word/styles.xml')!.async('string');
    expect(savedXml).toContain('w:val="一"'); expect(savedXml).not.toContain('新一');
    core.acknowledgeSave(saved); expect(core.state().dirty).toBe(true);
    core.editor.commands.undo(); core.editor.commands.undo();
    expect(wordStyles(core.editor.state.doc, core.parsed).has('MoOne')).toBe(false);
    expect(wordStyles(core.editor.state.doc, core.parsed).has('MoTwo')).toBe(false);
    expect(core.editor.state.doc.textContent).toBe('');
  });

  it('自定义固定行距、高亮、字号及标题语义保存重开后仍可重复套用', async () => {
    const core = await open(); selectTarget(core);
    const f: WordFormatSnapshot = { text: { font: '宋体', fontAscii: 'Arial', sizeHalfPoints: 36, bold: true, italic: false, color: '#006699', highlight: '#BADA55' },
      paragraph: { align: 'center', indentFirstLine: -240, indentLeft: 480, lineRule: 'exact', lineRawTwips: 480, lineSpacing: null, spaceBefore: 120, spaceAfter: 60, widowControl: false } };
    applyWordStyle(core.editor, core.parsed, 'MoHeading', defineWordStyle('MoHeading', '合同标题', f, 2));
    expect(core.editor.state.doc.child(1).type.name).toBe('docHeading');
    const reopened = await WordCoreSession.open((await core.prepareSave()).bytes); sessions.push(reopened);
    const styleXml = (await xml(core)).match(/<w:style\b[^>]*w:styleId="MoHeading"[\s\S]*?<\/w:style>/)![0];
    for (const [tag, order] of [['pPr', PPR_CHILD_ORDER], ['rPr', RPR_CHILD_ORDER]] as const) {
      const inner = new RegExp(`<w:${tag}>([\\s\\S]*?)</w:${tag}>`).exec(styleXml)![1];
      const ranks = splitXmlChildren(inner).map(item => order.indexOf(item.name));
      expect(ranks).toEqual([...ranks].sort((a, b) => a - b));
    }
    const style = reopened.parsed.styles.get('MoHeading')!;
    expect(style).toMatchObject({ name: '合同标题', headingLevel: 2, display: { lineRule: 'exact', lineRawTwips: 480, shading: 'BADA55', sizeHalfPoints: 36, widowControl: false } });
    const start = reopened.editor.state.doc.child(0).nodeSize + reopened.editor.state.doc.child(1).nodeSize + 1;
    reopened.editor.commands.setTextSelection(start);
    expect(applyWordStyle(reopened.editor, reopened.parsed, 'MoHeading')).toBe(true);
    expect(coreToolbarFormatting(reopened)).toMatchObject({ activeStyle: { fontSizePt: 18, bold: true, highlight: '#BADA55' }, paragraph: { alignment: 'center', lineHeightExactPx: 32, firstLineIndentPx: -16 } });
    const docXml = await xml(reopened, 'word/document.xml');
    expect(docXml).not.toContain('https:');
    expect(docXml).toContain('w:val="MoHeading"');
  });

  it('导入定义改名保留原未知样式 XML，重开和撤销均一致，删除保留格式基底且可撤销', async () => {
    const core = await open();
    const original = (await xml(core)).match(/<w:style\b[^>]*w:styleId="Imported"[\s\S]*?<\/w:style>/)![0];
    expect(renameWordStyle(core.editor, core.parsed, 'Imported', '来源新名')).toBe(true);
    const renamed = (await xml(core)).match(/<w:style\b[^>]*w:styleId="Imported"[\s\S]*?<\/w:style>/)![0];
    expect(renamed.replace('来源新名', '来源样式')).toBe(original);
    const revision = core.state().revision;
    expect(deleteWordStyle(core.editor, core.parsed, 'Imported')).toBe(true);
    expect(core.state().revision).toBeGreaterThan(revision);
    expect(wordStyles(core.editor.state.doc, core.parsed).has('Imported')).toBe(false);
    const alias = core.editor.state.doc.firstChild!.attrs.styleId;
    expect(alias.startsWith('MoRetained_')).toBe(true);
    expect(await xml(core)).not.toContain('w:styleId="Imported"');
    expect(await xml(core)).toContain(`w:styleId="${alias}"`);
    const reopened = await WordCoreSession.open((await core.prepareSave()).bytes); sessions.push(reopened);
    expect(reopened.parsed.styles.get(alias)?.display).toMatchObject({ bold: true, sizeHalfPoints: 30 });
    expect(reopened.parsed.styles.get(alias)?.semiHidden).toBe(true);
    core.editor.commands.undo(); expect(await xml(core)).toContain(renamed);
    core.editor.commands.undo(); expect(await xml(core)).toContain(original);
  });

  it('格式刷复制不改 dirty，局部应用保留超链接、语言及未知属性，单步撤销并可保存重开', async () => {
    const core = await open(); core.editor.commands.setTextSelection(1);
    const f = captureWordFormat(core.editor, core.formattingSource());
    expect(core.state()).toMatchObject({ revision: 0, dirty: false, canUndo: false });
    selectTarget(core);
    const start = core.editor.state.selection.from;
    core.editor.commands.setTextSelection({ from: start + 1, to: start + 5 });
    const original = core.editor.getJSON();
    expect(applyWordFormatPainter(core.editor, f)).toBe(true);
    expect(coreToolbarFormatting(core)).toMatchObject({ activeStyle: { fontFamily: '宋体', fontSizePt: 16, bold: true, italic: false, color: '#D92D20' } });
    expect(core.editor.state.doc.child(1).textContent).toBe('Target');
    const saved = await core.prepareSave();
    const zip = await JSZip.loadAsync(saved.bytes);
    const text = await zip.file('word/document.xml')!.async('string');
    expect(text).toContain('w:val="en-US"'); expect(text).toContain('<w:noProof'); expect(text).toContain('<w:hyperlink');
    expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain('https://example.com/target');
    expect(core.editor.commands.undo()).toBe(true); expect(core.editor.getJSON()).toEqual(original);
    core.editor.commands.redo();
    const reopened = await WordCoreSession.open(saved.bytes); sessions.push(reopened);
    reopened.editor.commands.setTextSelection({ from: start + 1, to: start + 5 });
    expect(coreToolbarFormatting(reopened).activeStyle).toMatchObject({ fontSizePt: 16, bold: true, italic: false, color: '#D92D20' });
  });

  it('预设标题能回到正文，目标末端下一段保持原样，只读和组合输入拒绝写入', async () => {
    const core = await open(); selectTarget(core);
    const next = core.editor.state.doc.child(2);
    const definition = coreStyleDefinition(BUILTIN_STYLES.find(item => item.id === 'Heading1')!, core);
    expect(applyWordStyle(core.editor, core.parsed, 'Heading1', definition)).toBe(true);
    expect(core.editor.state.doc.child(1).type.name).toBe('docHeading');
    expect(core.editor.state.doc.child(2).eq(next)).toBe(true);
    expect(applyWordStyle(core.editor, core.parsed, 'Normal')).toBe(true);
    expect(core.editor.state.doc.child(1).type.name).toBe('docParagraph');
    const readonly = await open(true); const before = readonly.editor.getJSON();
    expect(applyWordStyle(readonly.editor, readonly.parsed, 'Heading1', definition)).toBe(false);
    expect(renameWordStyle(readonly.editor, readonly.parsed, 'Imported', '不能改')).toBe(false);
    expect(importWordStyles(readonly.editor, [definition])).toBe(false);
    expect(applyWordFormatPainter(readonly.editor, captureWordFormat(core.editor, core.formattingSource()))).toBe(false);
    expect(readonly.editor.getJSON()).toEqual(before);
    Object.defineProperty(core.editor.view, 'composing', { configurable: true, value: true });
    expect(applyWordStyle(core.editor, core.parsed, 'Heading1', definition)).toBe(false);
    expect(renameWordStyle(core.editor, core.parsed, 'Imported', '不能改')).toBe(false);
  });

  it('表格矩形多格选区的样式与格式刷只改目标列，保留表格及表后正文', async () => {
    const cell = (text: string) => `<w:tc><w:tcPr/><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const core = await open(false, `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell('甲')}${cell('乙')}</w:tr><w:tr>${cell('丙')}${cell('丁')}</w:tr></w:tbl><w:p><w:r><w:t>表后</w:t></w:r></w:p>`);
    const cells: number[] = [];
    core.editor.state.doc.descendants((node, pos) => { if (node.type.name === 'docTableCell') cells.push(pos); });
    const snapshot = captureWordFormat(core.editor, core.formattingSource());
    snapshot.text.sizeHalfPoints = 40; snapshot.paragraph.align = 'center';
    core.editor.view.dispatch(core.editor.state.tr.setSelection(CellSelection.create(core.editor.state.doc, cells[0], cells[2])));
    const original = core.editor.getJSON();
    expect(applyWordStyle(core.editor, core.parsed, 'MoCells', defineWordStyle('MoCells', '表格文字', snapshot))).toBe(true);
    const paragraphs: Array<{ text: string; styleId: string | null; align: string | null }> = [];
    core.editor.state.doc.descendants(node => { if (node.isTextblock) paragraphs.push({ text: node.textContent, styleId: node.attrs.styleId, align: node.attrs.align }); });
    expect(paragraphs.filter(p => p.styleId === 'MoCells').map(p => p.text)).toEqual(['甲', '丙']);
    core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(original);
    expect(applyWordFormatPainter(core.editor, snapshot)).toBe(true);
    expect(core.editor.state.doc.lastChild!.attrs.align).not.toBe('center');
    const reopened = await WordCoreSession.open((await core.prepareSave()).bytes); sessions.push(reopened);
    const centered: string[] = [];
    reopened.editor.state.doc.descendants(node => { if (node.isTextblock && node.attrs.align === 'center') centered.push(node.textContent); });
    expect(centered).toEqual(['甲', '丙']);
    expect(reopened.editor.state.doc.firstChild!.childCount).toBe(2);
    expect(reopened.editor.state.doc.lastChild!.textContent).toBe('表后');
  });
});
