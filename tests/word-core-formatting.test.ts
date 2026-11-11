/** 功能：以真实 Word 事务和 DOCX 重开验证工具栏格式迁移及来源保留；时间：2026-11-10；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { WordCoreSession, buildBlankDocx } from '../packages/word-core/src/index';
import { applyWordTextFormat, applyWordParagraphFormat, readWordFormatting, setActiveSubEditor, toggleWordInlineFormat } from '../packages/word-core/src/editing/index';
import { coreFontFormat, coreParagraphFormat, coreToolbarFormatting } from '../packages/editor-word/src/core-formatting';

const sessions: WordCoreSession[] = [];
/** 功能：构造具有不同字号、字体、语言及段落设置的真实包；时间：2026-11-10；作者：lq。 */
async function open(readonly = false) {
  const zip = await JSZip.loadAsync(await buildBlankDocx());
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    <w:p><w:pPr><w:ind w:left="360" w:firstLineChars="200"/><w:spacing w:line="360" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="宋体"/><w:sz w:val="24"/><w:lang w:val="zh-CN"/></w:rPr><w:t>中文</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="黑体"/><w:sz w:val="32"/><w:i/><w:lang w:val="en-US"/></w:rPr><w:t>Latin</w:t></w:r></w:p>
    <w:p><w:pPr><w:jc w:val="right"/><w:spacing w:line="480" w:lineRule="auto"/></w:pPr><w:r><w:t>第二段</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const session = await WordCoreSession.open(await zip.generateAsync({ type: 'uint8array' }), { readonly });
  sessions.push(session);
  return session;
}
beforeEach(() => {
  // 功能：仅补齐 jsdom 的调度/Canvas 缺省，不替换解析、事务及保存实现；时间：2026-11-10；作者：lq。
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => setTimeout(() => callback(performance.now()), 0));
  vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id));
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
});
afterEach(() => { setActiveSubEditor(null); for (const session of sessions.splice(0)) session.dispose(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });

describe('M3 字符与段落工具栏', () => {
  it('混合格式准确回填，改变字号不覆盖各文本片段的字体、语言及斜体', async () => {
    const core = await open();
    core.editor.commands.setTextSelection({ from: 1, to: 8 });
    expect(coreToolbarFormatting(core).mixedText).toEqual(expect.arrayContaining(['fontFamily', 'fontSizePt', 'italic']));
    expect(coreFontFormat(core.editor, { fontSizePt: 18 })).toBe(true);
    const format = coreToolbarFormatting(core);
    expect(format.activeStyle.fontSizePt).toBe(18);
    expect(format.mixedText).toContain('fontFamily');
    const saved = await core.prepareSave();
    const zip = await JSZip.loadAsync(saved.bytes);
    const xml = await zip.file('word/document.xml')!.async('text');
    for (const value of ['宋体', '黑体', 'zh-CN', 'en-US', '<w:i', 'w:val="36"']) expect(xml).toContain(value);
    const reopened = await WordCoreSession.open(saved.bytes); sessions.push(reopened);
    reopened.editor.commands.setTextSelection({ from: 1, to: 8 });
    expect(coreToolbarFormatting(reopened).activeStyle.fontSizePt).toBe(18);
    expect(coreToolbarFormatting(reopened).mixedText).toContain('fontFamily');
  });

  it('颜色和任意高亮色经过原生格式保存重开，取消颜色不伪造黑色', async () => {
    const core = await open();
    core.editor.commands.setTextSelection({ from: 1, to: 3 });
    coreFontFormat(core.editor, { color: '#FF3300', highlight: '#BADA55', fontFamily: '微软雅黑' });
    expect(coreToolbarFormatting(core).activeStyle).toMatchObject({ color: '#FF3300', highlight: '#BADA55', fontFamily: '微软雅黑' });
    const saved = await core.prepareSave();
    const reopened = await WordCoreSession.open(saved.bytes); sessions.push(reopened);
    reopened.editor.commands.setTextSelection({ from: 1, to: 3 });
    expect(coreToolbarFormatting(reopened).activeStyle).toMatchObject({ color: '#FF3300', highlight: '#BADA55', fontFamily: '微软雅黑' });
    coreFontFormat(reopened.editor, { color: undefined, highlight: undefined });
    const xml = await (await JSZip.loadAsync((await reopened.prepareSave()).bytes)).file('word/document.xml')!.async('text');
    expect(xml).not.toContain('FF3300'); expect(xml).not.toContain('BADA55');
  });

  it('多个段落一次设置对齐和精确单位，撤销一步恢复，模式切换不遗留旧行距', async () => {
    const core = await open();
    core.editor.commands.selectAll();
    expect(coreToolbarFormatting(core).mixedParagraph).toContain('alignment');
    coreParagraphFormat(core.editor, { alignment: 'center', lineHeightExactPx: 24, firstLineIndentPx: -16, spaceBeforePx: 8 });
    let format = coreToolbarFormatting(core);
    expect(format.paragraph).toMatchObject({ alignment: 'center', lineHeightExactPx: 24, firstLineIndentPx: -16, spaceBeforePx: 8 });
    expect(core.editor.commands.undo()).toBe(true);
    expect(coreToolbarFormatting(core).mixedParagraph).toContain('alignment');
    expect(core.editor.commands.redo()).toBe(true);
    coreParagraphFormat(core.editor, { lineHeight: 2, lineHeightExactPx: undefined, lineHeightAtLeastPx: undefined });
    format = coreToolbarFormatting(core);
    expect(format.paragraph.lineHeight).toBe(2); expect(format.paragraph.lineHeightExactPx).toBeUndefined();
    const xml = await (await JSZip.loadAsync((await core.prepareSave()).bytes)).file('word/document.xml')!.async('text');
    expect(xml).toContain('w:line="480"'); expect(xml).toContain('w:hanging="240"'); expect(xml).not.toContain('w:firstLineChars="200"');
    const reopened = await WordCoreSession.open((await core.prepareSave()).bytes); sessions.push(reopened);
    reopened.editor.commands.selectAll();
    expect(coreToolbarFormatting(reopened).paragraph).toMatchObject({ firstLineIndentPx: -16, lineHeight: 2 });
  });

  it('选区结束于下一段开头时不改下一段，光标格式只影响后续输入', async () => {
    const core = await open();
    const second = core.editor.state.doc.child(0).nodeSize;
    core.editor.commands.setTextSelection({ from: 1, to: second + 1 });
    applyWordParagraphFormat(core.editor, { align: 'center' });
    expect(core.editor.state.doc.child(1).attrs.align).toBe('right');
    core.editor.commands.setTextSelection(2);
    applyWordTextFormat(core.editor, { sizeHalfPoints: 40 });
    core.editor.commands.insertContent('新');
    expect(readWordFormatting(core.editor, core.parsed).text.value.sizeHalfPoints).toBe(40);
    expect(core.editor.state.doc.child(0).child(0).marks.find(mark => mark.type.name === 'docTextStyle')?.attrs.sizeHalfPoints).toBe(24);
  });

  it('只读会话拒绝格式修改，其他文档子编辑器不串入命令目标', async () => {
    const readonly = await open(true);
    const other = await open();
    setActiveSubEditor(other.editor);
    expect(readonly.commandEditor()).toBe(readonly.editor);
    const before = readonly.editor.getJSON();
    expect(coreFontFormat(readonly.editor, { fontSizePt: 42 })).toBe(false);
    expect(coreParagraphFormat(readonly.editor, { alignment: 'center' })).toBe(false);
    expect(readonly.editor.getJSON()).toEqual(before);
  });

  it('粗斜体切换覆盖显式关闭标记，并按当前活动子编辑器返回历史状态', async () => {
    const core = await open();
    core.editor.commands.setTextSelection({ from: 1, to: 3 });
    applyWordTextFormat(core.editor, { boldOff: true });
    expect(toggleWordInlineFormat(core.editor, core.parsed, 'bold')).toBe(true);
    expect(coreToolbarFormatting(core).activeStyle.bold).toBe(true);
    const reopened = await WordCoreSession.open((await core.prepareSave()).bytes); sessions.push(reopened);
    reopened.editor.commands.setTextSelection({ from: 1, to: 3 });
    expect(coreToolbarFormatting(reopened).activeStyle.bold).toBe(true);
    expect(toggleWordInlineFormat(core.editor, core.parsed, 'bold')).toBe(true);
    expect(coreToolbarFormatting(core).activeStyle.bold).toBe(false);
    const sub = await open();
    core.editor.view.dom.append(sub.editor.view.dom);
    setActiveSubEditor(sub.editor);
    expect(core.commandEditor()).toBe(sub.editor);
    expect(core.state().canUndo).toBe(false);
    sub.editor.commands.insertContent('子编辑器');
    expect(core.state().canUndo).toBe(true);
    setActiveSubEditor(null);
    sub.editor.view.dom.remove();
  });
});
