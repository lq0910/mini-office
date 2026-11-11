/** 功能：以真实嵌套文本框视图验证即时保存、主历史和样式投影；时间：2026-11-11；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { WordCoreSession, buildBlankDocx } from '../packages/word-core/src/index';
import { getActiveSubEditor, setActiveSubEditor } from '../packages/word-core/src/upstream/apps/docs/src/renderer/editor/active-editor';
import { applyWordStyle, defineWordStyle, wordStyles, renameWordStyle, readWordList, changeWordListLevel } from '../packages/word-core/src/editing/index';
import { insertWordMergeField, wordMergeFields } from '../packages/word-core/src/editing/mail-merge';
const sessions: WordCoreSession[] = [];
/** 功能：公共 VML 文本框夹具保留段落未知属性和无关正文；时间：2026-11-11；作者：lq。 */
async function open(bytes?: Uint8Array) {
  if (!bytes) {
    const zip = await JSZip.loadAsync(await buildBlankDocx());
    zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:body><w:p><w:r><w:t>正文保持</w:t></w:r></w:p><w:p><w:r><w:pict><v:shape id="box" style="position:absolute;margin-left:10pt;margin-top:10pt;width:250pt;height:100pt"><v:textbox><w:txbxContent><w:p w:rsidR="12345678"><w:pPr><w:keepNext/><w:spacing w:after="120"/><w:suppressOverlap/></w:pPr><w:r><w:rPr><w:lang w:val="zh-CN"/></w:rPr><w:t>文本框原文</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
    bytes = await zip.generateAsync({ type: 'uint8array' });
  }
  const core = await WordCoreSession.open(bytes); sessions.push(core);
  document.body.append(core.editor.view.dom);
  await new Promise(resolve => setTimeout(resolve, 20));
  const dom = core.editor.view.dom.querySelector<HTMLElement>('.doc-textbox-editor');
  if (!dom) throw new Error('真实文本框未挂载');
  dom.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
  dom.focus();
  const sub = getActiveSubEditor(); if (!sub) throw new Error('文本框未激活');
  return { core, sub };
}
/** 功能：读取真实存盘 XML，不用视图断言代替序列化；时间：2026-11-11；作者：lq。 */
async function xml(core: WordCoreSession) { return (await JSZip.loadAsync((await core.prepareSave()).bytes)).file('word/document.xml')!.async('string'); }
beforeEach(() => {
  vi.stubGlobal('requestAnimationFrame', (cb: FrameRequestCallback) => setTimeout(() => cb(performance.now()), 0));
  vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id));
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
  // 功能：jsdom 仅补齐选择滚动所需几何接口，真实布局另由浏览器验证；时间：2026-11-11；作者：lq。
  Object.defineProperty(Range.prototype, 'getClientRects', { configurable: true, value: () => [] });
  Object.defineProperty(Range.prototype, 'getBoundingClientRect', { configurable: true, value: () => new DOMRect() });
});
afterEach(() => { setActiveSubEditor(null); for (const core of sessions.splice(0)) core.dispose(); document.body.replaceChildren(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });
describe('M3 真实文本框受控投影', () => {
  it('文本框新合并字段即时反映在主摘要，删除和撤销不沿用来源缓存', async () => {
    const { core, sub } = await open(); expect(insertWordMergeField(sub, '框内姓名')).toBe(true);
    expect(wordMergeFields(core.editor.state.doc, core.parsed).fields).toContain('框内姓名');
    const reopened = await open((await core.prepareSave()).bytes); reopened.sub.commands.selectAll(); reopened.sub.commands.deleteSelection();
    expect(wordMergeFields(reopened.core.editor.state.doc, reopened.core.parsed).fields).not.toContain('框内姓名');
    reopened.sub.commands.undo(); expect(wordMergeFields(reopened.core.editor.state.doc, reopened.core.parsed).fields).toContain('框内姓名');
  });
  it('文本框新图片跨保存快照仍有媒体和关系，重复保存及重开不丢失', async () => {
    const { core, sub } = await open();
    const dataUrl = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6+AAAAABJRU5ErkJggg==';
    sub.commands.insertContent({ type: 'docInlineImage', attrs: { dataUrl, widthPx: 40, heightPx: 20 } });
    sub.commands.undo(); expect(sub.state.doc.firstChild!.firstChild?.type.name).not.toBe('docInlineImage');
    sub.commands.redo(); expect(sub.state.doc.firstChild!.firstChild?.type.name).toBe('docInlineImage');
    for (let attempt = 0; attempt < 2; attempt++) {
      const prepared = await core.prepareSave(), zip = await JSZip.loadAsync(prepared.bytes);
      const saved = await zip.file('word/document.xml')!.async('string');
      expect(saved).not.toContain('aidocs-new-pic-');
      expect(Object.keys(zip.files).filter(path => /^word\/media\/.*\.png$/.test(path))).toHaveLength(1);
      const reopened = await open(prepared.bytes);
      const image = reopened.sub.state.doc.firstChild!.firstChild;
      expect(image?.type.name).toBe('docInlineImage'); expect(image?.attrs).toMatchObject({ widthPx: 40, heightPx: 20 });
      expect(reopened.core.parsed.blocks.find(block => block.textboxes)?.textboxes![0].paras[0].runs.find(run => run.image)?.image?.dataUrl).toBeTruthy();
      const media = Object.keys(zip.files).find(path => /^word\/media\/.*\.png$/.test(path))!;
      expect(await zip.file(media)!.async('base64')).toBe(dataUrl.split(',')[1]);
    }
  });
  it('未失焦内容即时进入保存，工具栏和快捷键共享主历史', async () => {
    const { core, sub } = await open();
    sub.commands.setTextSelection({ from: 1, to: sub.state.doc.content.size - 1 });
    sub.commands.insertContent('立即存盘');
    expect(await xml(core)).toContain('立即存盘');
    expect(core.editor.state.doc.firstChild!.textContent).toBe('正文保持');
    expect(sub.can().undo()).toBe(true); sub.commands.undo();
    expect(sub.getText()).toBe('文本框原文');
    sub.commands.redo(); expect(sub.getText()).toBe('立即存盘');
    const reopened = await open((await core.prepareSave()).bytes); expect(reopened.sub.getText()).toBe('立即存盘');
  });
  it('段落格式、边框及制表位保存重开，并保留未知来源属性', async () => {
    const { core, sub } = await open();
    sub.commands.updateAttributes('docParagraph', { align: 'center', indentRight: 360, indentFirstLine: 240, spaceAfter: 300, keepNext: false, keepLines: true, widowControl: false, tabStops: JSON.stringify([{ pos: 1440, val: 'right' }]), borders: 'b', borderLines: JSON.stringify({ b: { val: 'double', szPt: 1.5, color: '006699' } }) });
    const saved = await xml(core); expect(saved).toContain('w:suppressOverlap'); expect(saved).toContain('w:val="double"'); expect(saved).toContain('w:right="360"');
    const reopened = await open((await core.prepareSave()).bytes);
    expect(reopened.sub.state.doc.firstChild!.attrs).toMatchObject({ align: 'center', indentRight: 360, indentFirstLine: 240, spaceAfter: 300, keepNext: false, keepLines: true, widowControl: false });
    expect(JSON.parse(reopened.sub.state.doc.firstChild!.attrs.tabStops)).toMatchObject([{ pos: 1440, val: 'right' }]);
    core.editor.commands.undo(); expect(sub.state.doc.firstChild!.attrs.align).toBeNull();
  });
  it('文本框创建样式与正文共享目录，改名和撤销投影不会回写旧目录', async () => {
    const opened = await open(), core = opened.core; let sub = opened.sub;
    const definition = defineWordStyle('BoxStyle', '框内样式', { text: { bold: true, sizeHalfPoints: 30 }, paragraph: { align: 'right', keepLines: true } });
    expect(applyWordStyle(sub, core.formattingSource(), 'BoxStyle', definition)).toBe(true); sub = core.commandEditor(); expect(sub.state.doc.firstChild!.attrs.styleId).toBe('BoxStyle');
    expect(wordStyles(core.editor.state.doc, core.parsed).has('BoxStyle')).toBe(true);
    expect(core.editor.state.doc.firstChild!.attrs.styleId).not.toBe('BoxStyle');
    core.editor.commands.undo(); sub = core.commandEditor(); expect(sub.state.doc.firstChild!.attrs.styleId).toBeNull();
    expect(wordStyles(core.editor.state.doc, core.parsed).has('BoxStyle')).toBe(false);
    core.editor.commands.redo(); sub = core.commandEditor(); expect(sub.state.doc.firstChild!.attrs.styleId).toBe('BoxStyle');
    expect(renameWordStyle(core.editor, core.parsed, 'BoxStyle', '全区改名')).toBe(true);
    sub = core.commandEditor(); expect(wordStyles(sub.state.doc, core.parsed).get('BoxStyle')?.name).toBe('全区改名');
    sub.commands.insertContent('追加');
    expect(wordStyles(core.editor.state.doc, core.parsed).get('BoxStyle')?.name).toBe('全区改名');
    const reopened = await open((await core.prepareSave()).bytes); expect(reopened.sub.state.doc.firstChild!.attrs.styleId).toBe('BoxStyle');
  });
  it('列表分配、层级和取消可撤销，未确认输入拒绝保存', async () => {
    const { core, sub } = await open();
    expect(core.toggleList('ordered', sub)).toBe(true); expect(readWordList(sub).kind).toBe('ordered');
    expect(changeWordListLevel(sub, 1)).toBe(true); expect(readWordList(sub).level).toBe(1);
    sub.commands.undo(); expect(readWordList(sub).level).toBe(0);
    expect(core.toggleList('ordered', sub)).toBe(true); expect(readWordList(sub).kind).toBeUndefined();
    sub.commands.undo(); expect(readWordList(sub).kind).toBe('ordered');
    const reopened = await open((await core.prepareSave()).bytes); expect(readWordList(reopened.sub).kind).toBe('ordered');
    setActiveSubEditor(sub); Object.defineProperty(sub.view, 'composing', { configurable: true, value: true });
    await expect(core.prepareSave()).rejects.toThrow('中文输入尚未确认');
    Object.defineProperty(sub.view, 'composing', { configurable: true, value: false });
    await expect(core.prepareSave()).resolves.toHaveProperty('bytes');
  });
});
