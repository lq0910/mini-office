/** 功能：验证新核心标尺/制表位、目录域、原生历史及保存快照隔离；时间：2026-11-11；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { WordCoreSession, buildBlankDocx, generateTocFieldXml } from '../packages/word-core/src/index';
import { applyWordRuler, wordRulerTabs } from '../packages/word-core/src/editing/ruler';
import { insertWordToc, wordHeadings, wordTocSnapshot, followWordToc } from '../packages/word-core/src/editing/toc';
const sessions: WordCoreSession[] = [];
/** 功能：包含重复中文标题、继承/直接/清除制表位和未知段落属性的真实包；时间：2026-11-11；作者：lq。 */
async function open(readonly = false, prefix = '') {
  const zip = await JSZip.loadAsync(await buildBlankDocx());
  zip.file('word/styles.xml', (await zip.file('word/styles.xml')!.async('string')).replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="Tabs"><w:name w:val="Tabs"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="1440"/><w:tab w:val="left" w:pos="2400"/></w:tabs></w:pPr></w:style></w:styles>'));
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${prefix}<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>重复标题</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Tabs"/><w:tabs><w:tab w:val="clear" w:pos="2400"/><w:tab w:val="bar" w:pos="3200"/></w:tabs><w:suppressLineNumbers/></w:pPr><w:r><w:rPr><w:lang w:val="zh-CN"/></w:rPr><w:t>内容 &amp; 正文</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>重复标题</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const core = await WordCoreSession.open(await zip.generateAsync({ type: 'uint8array' }), { readonly }); sessions.push(core); return core;
}
/** 功能：回读真实 DOCX 序列化字节，不只断言内存对象；时间：2026-11-11；作者：lq。 */
async function reopen(core: WordCoreSession) { const next = await WordCoreSession.open((await core.prepareSave()).bytes); sessions.push(next); return next; }
/** 功能：检查本次保存的原生正文 XML；时间：2026-11-11；作者：lq。 */
async function xml(core: WordCoreSession) { return (await JSZip.loadAsync((await core.prepareSave()).bytes)).file('word/document.xml')!.async('string'); }
beforeEach(() => {
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => setTimeout(() => callback(performance.now()), 0));
  vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id));
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
});
afterEach(() => { for (const core of sessions.splice(0)) core.dispose(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });

describe('M3 标尺和目录', () => {
  it('同格多段标尺保存重开，跨格选区拒绝修改且邻格不受影响', async () => {
    const table = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>格一</w:t></w:r></w:p><w:p><w:r><w:t>续段</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p><w:r><w:t>格二</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const core = await open(false, table);
    const positions: number[] = []; core.editor.state.doc.descendants((node, pos) => { if (node.type.name === 'docParagraph' && ['格一','续段','格二'].includes(node.textContent)) positions.push(pos); });
    core.editor.commands.setTextSelection({ from: positions[0]+1, to: positions[1]+2 });
    expect(applyWordRuler(core.editor, core.parsed, { indentLeft: 240, tabs: [{ pos: 900, val: 'right' }] })).toBe(true);
    expect(core.editor.state.doc.nodeAt(positions[2])!.attrs.indentLeft).not.toBe(240);
    const saved = await reopen(core); const cell = saved.editor.state.doc.firstChild!.firstChild!.firstChild!;
    expect(cell.child(0).attrs.indentLeft).toBe(240); expect(cell.child(1).attrs.indentLeft).toBe(240);
    expect(JSON.parse(cell.child(0).attrs.tabStops)).toContainEqual({ pos: 900, val: 'right' });
    core.editor.commands.undo(); expect(core.editor.state.doc.nodeAt(positions[0])!.attrs.indentLeft).not.toBe(240);
    const revision = core.state().revision; core.editor.commands.setTextSelection({ from: positions[0]+1, to: positions[2]+2 });
    expect(applyWordRuler(core.editor, core.parsed, { indentLeft: 600 })).toBe(false); expect(core.state().revision).toBe(revision);
  });
  it('缩进与继承制表位单步历史，保存保留语言、未知段落属性和隐藏 bar', async () => {
    const core = await open(); const start = core.editor.state.doc.child(0).nodeSize+1; core.editor.commands.setTextSelection(start);
    const original = core.editor.getJSON();
    expect(wordRulerTabs(core.editor.state.selection.$from.parent.attrs, core.parsed)).toEqual([expect.objectContaining({ pos: 1440, val: 'right', leader: 'dot' }), { pos: 3200, val: 'bar' }]);
    expect(applyWordRuler(core.editor, core.parsed, { indentLeft: 720, indentRight: 360, indentFirstLine: -360, tabs: [{ pos: 1800, val: 'decimal', leader: 'underscore' }] })).toBe(true);
    expect(core.editor.commands.undo()).toBe(true); expect(core.editor.getJSON()).toEqual(original);
    core.editor.commands.redo(); const saved = await reopen(core);
    const para = saved.editor.state.doc.child(1); expect(para.attrs).toMatchObject({ indentLeft: 720, indentRight: 360, indentFirstLine: -360 });
    const tabs = JSON.parse(para.attrs.tabStops); expect(tabs).toContainEqual({ pos: 1440, val: 'clear' }); expect(tabs).toContainEqual({ pos: 2400, val: 'clear' }); expect(tabs).toContainEqual({ pos: 3200, val: 'bar' });
    expect(wordRulerTabs(para.attrs, saved.parsed)).toContainEqual({ pos: 1800, val: 'decimal', leader: 'underscore' });
    expect(await xml(core)).toContain('<w:suppressLineNumbers/>'); expect(await xml(core)).toContain('zh-CN');
    expect(saved.editor.state.doc.textContent).toBe(core.editor.state.doc.textContent);
  });
  it('多段缩进保持各自字符属性，空/非法参数、只读和组合输入不改历史', async () => {
    const core = await open(); core.editor.commands.selectAll();
    expect(applyWordRuler(core.editor, core.parsed, { indentLeft: 555 })).toBe(true);
    for (let i=0; i<3; i++) expect(core.editor.state.doc.child(i).attrs.indentLeft).toBe(555);
    const revision = core.state().revision;
    expect(applyWordRuler(core.editor, core.parsed, { indentLeft: NaN })).toBe(false);
    expect(applyWordRuler(core.editor, core.parsed, { tabs: [{ pos: -2, val: 'left' }] })).toBe(false);
    Object.defineProperty(core.editor.view, 'composing', { configurable: true, get: () => true });
    expect(applyWordRuler(core.editor, core.parsed, { indentRight: 444 })).toBe(false); expect(core.state().revision).toBe(revision);
    const locked = await open(true); expect(applyWordRuler(locked.editor, locked.parsed, { indentLeft: 222 })).toBe(false); expect(locked.state().dirty).toBe(false);
  });
  it('真实 TOC 域插入/更新一步撤销，重复标题顺序、三级过滤及重开保持一致', async () => {
    const core = await open(); const before = core.editor.getJSON();
    expect(insertWordToc(core.editor, core.parsed)).toBe(true); expect(core.editor.state.doc.child(0).attrs.fieldDisplay.left).toBe('重复标题');
    expect(core.editor.state.doc.child(1).attrs.fieldDisplay.level).toBe(2);
    expect(await xml(core)).toContain(' TOC '); expect(await xml(core)).toContain('w:dirty="true"');
    core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before); core.editor.commands.redo();
    const heading = wordHeadings(core.editor.state.doc, core.parsed.styles)[0]; core.editor.commands.insertContentAt(heading.pos+1, '修改');
    const beforeUpdate = core.editor.getJSON(); insertWordToc(core.editor, core.parsed, true);
    expect(core.editor.state.doc.child(0).attrs.fieldDisplay.left).toBe('修改重复标题');
    core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(beforeUpdate); core.editor.commands.redo();
    const saved = await reopen(core); expect(saved.editor.state.doc.child(0).attrs.fieldDisplay.left).toBe('修改重复标题');
    expect(() => insertWordToc(saved.editor, saved.parsed)).toThrow('已有目录');
    expect(insertWordToc(saved.editor, saved.parsed, true)).toBe(true);
  });
  /** 功能：复现旧恢复快照丢失 TOC 域/样式并带显式下划线，显示接入同源目录且未编辑保存保持原字节；时间：2026-11-14；作者：lq。 */
  it('旧转换的连续目录链接补充显示模型并可跳转，不改来源 XML', async () => {
    const entry = (anchor: string, title: string, page: string) => `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9000" w:leader="dot"/></w:tabs></w:pPr><w:hyperlink w:anchor="${anchor}"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:rFonts w:eastAsia="仿宋"/><w:sz w:val="28"/><w:b/><w:u w:val="single"/></w:rPr><w:t>${title}</w:t><w:tab/><w:t>${page}</w:t></w:r></w:hyperlink></w:p>`;
    const prefix = entry('_TocOld1', '重复标题', '3') + entry('_TocOld2', '末级标题', '9');
    const core = await open(false, prefix);
    expect(core.editor.state.doc.firstChild!.type.name).toBe('docProtected');
    expect(core.editor.state.doc.firstChild!.attrs.fieldDisplay).toMatchObject({ kind: 'tocLine', left: '重复标题', right: '3', anchor: '_TocOld1' });
    expect(core.editor.state.doc.firstChild!.attrs.fieldDisplay.runStyleId).toBeUndefined();
    expect(followWordToc(core.editor, core.parsed, 0)).toBe(true);
    expect(core.state().dirty).toBe(false);
    expect((await core.prepareSave()).bytes).toEqual(core.parsed.internal.originalBytes);
    expect(await xml(core)).toContain(prefix);
    const single = await open(false, entry('_TocOnly', '普通书签链接', '3'));
    expect(single.editor.state.doc.firstChild!.type.name).toBe('docParagraph');
  });
  it('实际分页缓存按重复标题各自的位置进入快照，不增加 revision/dirty 或污染并发保存', async () => {
    const core = await open(); insertWordToc(core.editor, core.parsed); const revision = core.state().revision;
    const headings = wordHeadings(core.editor.state.doc, core.parsed.styles);
    core.recordHeadingPages(new Map([[headings[0].pos, 'iv'], [headings[1].pos, 'vi']])); expect(core.state().revision).toBe(revision);
    const saving = core.prepareSave(); core.editor.commands.insertContentAt(headings[0].pos+1, '新');
    const prepared = await saving; const saved = await WordCoreSession.open(prepared.bytes); sessions.push(saved);
    expect(saved.editor.state.doc.child(0).attrs.fieldDisplay.right).toBe('iv'); expect(saved.editor.state.doc.child(1).attrs.fieldDisplay.right).toBe('vi');
    core.acknowledgeSave(prepared); expect(core.state().dirty).toBe(true);
    const stale = await reopen(core); expect(stale.editor.state.doc.child(0).attrs.fieldDisplay.right ?? '').toBe('');
  });
  it('未编辑导入目录时测量不改原始字节，正文编辑后才更新缓存', async () => {
    const prefix = generateTocFieldXml([{ text: '重复标题', level: 1, pageNo: 99 }]).join('');
    const core = await open(false, prefix); const original = (await core.prepareSave()).bytes;
    const headings = wordHeadings(core.editor.state.doc, core.parsed.styles); core.recordHeadingPages(new Map([[headings[0].pos, '3']]));
    expect(core.state().dirty).toBe(false); expect((await core.prepareSave()).bytes).toEqual(original);
    core.editor.commands.insertContentAt(headings[1].pos+1, '变化');
    const changed = wordHeadings(core.editor.state.doc, core.parsed.styles); core.recordHeadingPages(new Map([[changed[0].pos, '3']]));
    expect((await reopen(core)).editor.state.doc.firstChild!.attrs.fieldDisplay.right).toBe('3');
  });
  it('无标题、重复插入、非完整导入目录、只读和输入法不会部分修改文档', async () => {
    const blank = await WordCoreSession.create(); sessions.push(blank); expect(() => insertWordToc(blank.editor, blank.parsed)).toThrow('标题'); expect(blank.state().dirty).toBe(false);
    const core = await open(); insertWordToc(core.editor, core.parsed); const revision = core.state().revision;
    expect(() => insertWordToc(core.editor, core.parsed)).toThrow('已有目录'); expect(core.state().revision).toBe(revision);
    Object.defineProperty(core.editor.view, 'composing', { configurable: true, get: () => true }); expect(insertWordToc(core.editor, core.parsed, true)).toBe(false);
    const locked = await open(true); expect(insertWordToc(locked.editor, locked.parsed)).toBe(false);
    const complex = await open(false, '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>复杂缓存</w:t></w:r></w:p>');
    expect(() => insertWordToc(complex.editor, complex.parsed, true)).toThrow(); expect(complex.state().dirty).toBe(false);
  });
});

/** 功能：验证目录缓存按标题级别区分同名项，以及区域安全保护与导航不改历史；时间：2026-11-11；作者：lq。 */
describe('M3 目录边界与导航', () => {
  // 功能：旧 DOC 章标题与前句粘连时定位标题字符，重复或仅正文引用不误跳；时间：2026-11-11；作者：lq。
  it('粘连章标题按句末完整章名定位，歧义目标拒绝猜测', async () => {
    const title = '第五章 供应商资格证明及相关文件要求';
    const core = await open(false, generateTocFieldXml([{ text: title, level: 1, pageNo: 99 }]).join(''));
    const paragraph = { type: 'docParagraph', content: [{ type: 'text', text: '前文必须满足。' + title }] };
    core.editor.commands.insertContentAt(core.editor.state.doc.content.size, paragraph);
    const revision = core.state().revision;
    expect(followWordToc(core.editor, core.parsed, 0)).toBe(true);
    expect(core.editor.state.selection.$from.parent.textBetween(core.editor.state.selection.$from.parentOffset, core.editor.state.selection.$from.parent.content.size)).toBe(title);
    expect(core.state().revision).toBe(revision);
    core.editor.commands.insertContentAt(core.editor.state.doc.content.size, paragraph);
    expect(followWordToc(core.editor, core.parsed, 0)).toBe(false);
  });
  // 功能：旧 DOC 转换后标题未携带样式或书签时按独立标题文字导航；时间：2026-11-11；作者：lq。
  it('旧文档缺失书签时匹配正文标题，列表章节号与项目前缀可以不同', async () => {
    const prefix = generateTocFieldXml([{ text: '第一章 竞争性磋商公告', level: 1, pageNo: 99 }]).join('');
    const core = await open(false, prefix);
    const pos = core.editor.state.doc.content.size;
    core.editor.commands.insertContentAt(pos, { type: 'docParagraph', content: [{ type: 'text', text: '项目名称竞争性磋商公告' }] });
    const revision = core.state().revision;
    expect(followWordToc(core.editor, core.parsed, 0)).toBe(true);
    expect(core.editor.state.selection.$from.parent.textContent).toBe('项目名称竞争性磋商公告');
    expect(core.state().revision).toBe(revision);
  });
  // 功能：导入 TOC 书签可以指向无标题样式的段落，必须优先于同名标题并保持正文原样；时间：2026-11-11；作者：lq。
  it('导入目录优先定位隐藏书签，目标无需标题样式或相同标题文字', async () => {
    const core = await open(); insertWordToc(core.editor, core.parsed);
    const tr = core.editor.state.tr;
    tr.setNodeMarkup(0, undefined, { ...tr.doc.child(0).attrs, fieldDisplay: { ...tr.doc.child(0).attrs.fieldDisplay, anchor: '_Toc2026' } });
    const last = tr.doc.content.size - tr.doc.lastChild!.nodeSize;
    tr.setNodeMarkup(last, undefined, { ...tr.doc.lastChild!.attrs, hiddenBookmarks: ['_Toc2026'] });
    core.editor.view.dispatch(tr);
    const revision = core.state().revision;
    expect(followWordToc(core.editor, core.parsed, 0)).toBe(true);
    expect(core.editor.state.selection.$from.before(1)).toBe(last);
    expect(core.state().revision).toBe(revision);
  });
  it('同名四级标题不会错配三级目录的页码，修饰键导航保持文档版本', async () => {
    const core = await open(); insertWordToc(core.editor, core.parsed);
    const headings = wordHeadings(core.editor.state.doc, core.parsed.styles);
    const snapshot = wordTocSnapshot(core.editor.getJSON(), [{ text: '重复标题', level: 4, pos: 100 }, ...headings], new Map([[100,'99'], [headings[0].pos,'iv'], [headings[1].pos,'vi']]));
    expect((snapshot.content![0].attrs!.fieldDisplay as {right:string}).right).toBe('iv');
    expect((snapshot.content![1].attrs!.fieldDisplay as {right:string}).right).toBe('vi');
    const revision = core.state().revision; expect(followWordToc(core.editor, core.parsed, 1)).toBe(true);
    expect(core.editor.state.selection.$from.parent.attrs.level).toBe(2); expect(core.state().revision).toBe(revision);
  });
  it('目录域边界外的正文与跨区域书签不会随更新丢失', async () => {
    const xml = generateTocFieldXml([{ text: '重复标题', level: 1, pageNo: 1 }])[0];
    const prefix = xml.replace('<w:r><w:fldChar', '<w:r><w:t>不能删除的正文</w:t></w:r><w:r><w:fldChar');
    const core = await open(false, prefix); const original = core.editor.getJSON();
    expect(() => insertWordToc(core.editor, core.parsed, true)).toThrow(); expect(core.editor.getJSON()).toEqual(original); expect(core.state().dirty).toBe(false);
    const bookmark = await open(false, xml.replace('<w:p>', '<w:p><w:bookmarkStart w:id="123" w:name="keep"/>'));
    expect(() => insertWordToc(bookmark.editor, bookmark.parsed, true)).toThrow(); expect(bookmark.state().dirty).toBe(false);
  });
});
