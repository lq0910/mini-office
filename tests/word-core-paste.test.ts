/** 功能：验证粘贴三模式、延迟重放单步历史、后续修改与文档隔离；时间：2026-11-11；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WordCoreSession } from '../packages/word-core/src/index';
import { applyWordPasteMode, dismissWordPaste, wordPasteRecord } from '../packages/word-core/src/editing/paste';
import { applyWordTextFormat } from '../packages/word-core/src/editing/formatting';
import { insertWordPicture, insertWordTable } from '../packages/word-core/src/editing/objects';
const cores: WordCoreSession[] = [];
async function create() { const core = await WordCoreSession.create(); cores.push(core); return core; }
/** 功能：使用原生 view.pasteHTML 走剪贴板解析和插件路径，独立载荷模拟真实事件数据；时间：2026-11-11；作者：lq。 */
function paste(core: WordCoreSession, html: string, text = '') {
  const event = { clipboardData: { getData: (type: string) => type === 'text/html' ? html : type === 'text/plain' ? text : '' } } as ClipboardEvent;
  core.editor.view.pasteHTML(html, event);
}
const HTML = '<p style="font-family: Georgia; font-size:24pt;color:#cc2200;text-align:right"><strong>来源中文</strong><em>斜体</em></p><p>第二段</p>';
beforeEach(() => { vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null); vi.stubGlobal('requestAnimationFrame', (fn: FrameRequestCallback) => setTimeout(() => fn(performance.now()), 0)); vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id)); });
afterEach(() => { cores.splice(0).forEach(core => core.dispose()); vi.restoreAllMocks(); vi.unstubAllGlobals(); });
describe('M3 核心粘贴', () => {
  it('原生图片粘贴保留可编辑尾段，紧接剪切分别撤销，内部与纯文本粘贴也有独立历史', async () => {
    // 功能：通过真实 schema 序列化原子图片，再走 HTML 剪贴板插件，覆盖无格式浮层时的历史边界；时间：2026-11-11；作者：lq。
    const core = await create();
    const source = await create(); const base64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Zl1sAAAAASUVORK5CYII=';
    insertWordPicture(source.editor, { base64, mime: 'image/png', dataUrl: `data:image/png;base64,${base64}`, widthPx: 1, heightPx: 1, label: '回归图片' });
    let sourcePos = 0; source.editor.state.doc.forEach((node, offset) => { if (node.attrs.blockType === 'image') sourcePos = offset; });
    source.editor.commands.setNodeSelection(sourcePos);
    const native = source.editor.view.serializeForClipboard(source.editor.state.selection.content()).dom;
    const before = core.editor.getJSON(); paste(core, native.innerHTML);
    const pasted = core.editor.getJSON();
    expect(wordPasteRecord(core.editor)).toBeNull();
    expect(core.editor.state.doc.lastChild?.type.name).toBe('docParagraph');
    let pos = -1; core.editor.state.doc.forEach((node, offset) => { if (node.attrs.blockType === 'image') pos = offset; });
    expect(pos).toBeGreaterThanOrEqual(0);
    core.editor.commands.setNodeSelection(pos);
    core.editor.view.dispatch(core.editor.state.tr.deleteSelection().setMeta('uiEvent', 'cut'));
    expect(core.editor.state.doc.childCount).toBeLessThan(pasted.content!.length);
    core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(pasted);
    const reopened = await WordCoreSession.open((await core.prepareSave()).bytes); cores.push(reopened);
    expect(reopened.editor.getJSON().content?.filter(node => node.attrs?.blockType === 'image')).toHaveLength(1);
    core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before);
    for (const html of ['<p data-pm-slice="1 1 []">内部文字</p>', '<p>普通文字</p>']) {
      const sample = await create(); paste(sample, html, '普通文字'); dismissWordPaste(sample.editor); const snapshot = sample.editor.getJSON();
      sample.editor.commands.insertContent('追加'); sample.editor.commands.undo(); expect(sample.editor.getJSON()).toEqual(snapshot);
      sample.editor.commands.undo(); expect(sample.editor.state.doc.textContent).toBe('');
    }
  });
  it('网页来源格式、目标格式和纯文本可反复切换，延迟后仍一次撤销至粘贴前且重做正确', async () => {
    const core = await create(); core.editor.commands.insertContent('前文');
    applyWordTextFormat(core.editor, { sizeHalfPoints: 28, font: 'Arial', color: '246B8E' });
    const before = core.editor.getJSON();
    paste(core, HTML, '来源中文斜体\n第二段');
    expect(core.editor.state.doc.textContent).toBe('前文来源中文斜体第二段');
    expect(wordPasteRecord(core.editor)?.mode).toBe('source');
    const runs = () => core.editor.getJSON().content!.flatMap(node => node.content ?? []).filter(node => node.type === 'text');
    expect(runs().find(node => 'text' in node && node.text === '来源中文')?.marks?.find(mark => mark.type === 'docTextStyle')?.attrs?.sizeHalfPoints).toBe(48);
    const originalNow = Date.now; vi.spyOn(Date, 'now').mockImplementation(() => originalNow() + 10000);
    expect(applyWordPasteMode(core.editor, 'merge')).toBe(true);
    expect(runs().find(node => 'text' in node && node.text === '来源中文')?.marks?.find(mark => mark.type === 'docTextStyle')?.attrs?.sizeHalfPoints).toBe(28);
    expect(runs().find(node => 'text' in node && node.text === '来源中文')?.marks?.some(mark => mark.type === 'bold')).toBe(true);
    expect(applyWordPasteMode(core.editor, 'text')).toBe(true);
    expect(runs().some(node => node.marks?.some(mark => ['bold', 'italic'].includes(mark.type)))).toBe(false);
    expect(core.editor.state.doc.textContent).toBe('前文来源中文斜体第二段');
    expect(applyWordPasteMode(core.editor, 'source')).toBe(true);
    const after = core.editor.getJSON(); core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before);
    core.editor.commands.redo(); expect(core.editor.getJSON()).toEqual(after);
    const reopened = await WordCoreSession.open((await core.prepareSave()).bytes); cores.push(reopened);
    expect(reopened.editor.state.doc.textContent).toBe(core.editor.state.doc.textContent);
  });
  it('后续输入/撤销使浮层失效，普通输入与粘贴分别撤销，不允许旧载荷覆盖新内容', async () => {
    const core = await create(); paste(core, HTML, '来源中文斜体\n第二段'); const pasted = core.editor.getJSON();
    core.editor.commands.insertContent('后续编辑'); expect(wordPasteRecord(core.editor)).toBeNull();
    expect(applyWordPasteMode(core.editor, 'text')).toBe(false);
    core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(pasted); expect(wordPasteRecord(core.editor)).toBeNull();
    core.editor.commands.undo(); expect(core.editor.state.doc.textContent).toBe('');
  });
  it('空段粘贴保留标题，HTML 单独存在可转文本，关闭选项不改版本', async () => {
    const core = await create(); paste(core, '<h2>粘贴标题</h2><p>尾段</p>');
    expect(core.editor.state.doc.firstChild?.type.name).toBe('docHeading');
    expect(applyWordPasteMode(core.editor, 'text')).toBe(true); expect(core.editor.state.doc.firstChild?.type.name).toBe('docParagraph');
    expect(core.editor.state.doc.childCount).toBe(2); expect(core.editor.state.doc.textContent).toBe('粘贴标题尾段');
    const version = core.state().revision; dismissWordPaste(core.editor); expect(core.state().revision).toBe(version);
  });
  it('纯文本采用光标格式、单列表格按行贴入单元格，多个编辑器载荷相互独立', async () => {
    const a = await create(); const b = await create(); paste(a, HTML, '来源中文斜体\n第二段');
    insertWordTable(b.editor, 2, 2); paste(b, '<table><tr><td>甲</td></tr><tr><td>乙</td></tr></table>', '甲\n乙');
    expect(b.editor.state.doc.firstChild?.type.name).toBe('docTable');
    expect(b.editor.state.doc.firstChild?.firstChild?.firstChild?.childCount).toBe(2);
    expect(applyWordPasteMode(a.editor, 'text')).toBe(true);
    expect(b.editor.state.doc.textContent).toBe('甲乙');
    const c = await create(); applyWordTextFormat(c.editor, { sizeHalfPoints: 36, font: 'Arial' });
    c.editor.view.pasteText('普通文字\n第二行', { clipboardData: { getData: (type: string) => type === 'text/plain' ? '普通文字\n第二行' : '' } } as ClipboardEvent);
    expect(c.editor.getJSON().content?.[0].content?.[0].marks?.find(mark => mark.type === 'docTextStyle')?.attrs?.sizeHalfPoints).toBe(36);
    expect(wordPasteRecord(c.editor)).toBeNull();
  });
  it('只读和输入法期间拒绝粘贴与模式切换，内部复制不出现来源浮层', async () => {
    const core = await create(); paste(core, '<p data-pm-slice="1 1 []"><strong>内部</strong></p>', '内部'); expect(wordPasteRecord(core.editor)).toBeNull();
    paste(core, HTML, '来源中文斜体\n第二段'); const before = core.editor.getJSON();
    Object.defineProperty(core.editor.view, 'composing', { configurable: true, get: () => true });
    expect(applyWordPasteMode(core.editor, 'text')).toBe(false); paste(core, '<p>禁止插入</p>'); expect(core.editor.getJSON()).toEqual(before);
    const locked = await WordCoreSession.create({ readonly: true }); cores.push(locked); paste(locked, HTML); expect(locked.state().dirty).toBe(false);
  });
  it('粘贴以标题或表格结尾时覆盖附加尾段事务，切换及撤销不残留内容', async () => {
    for (const html of ['<h2>单标题</h2>', '<table><tr><td><strong>甲</strong></td><td>乙</td></tr></table>']) {
      const core = await create(); const before = core.editor.getJSON(); paste(core, html);
      expect(wordPasteRecord(core.editor)).not.toBeNull();
      expect(applyWordPasteMode(core.editor, 'text')).toBe(true);
      expect(applyWordPasteMode(core.editor, 'source')).toBe(true);
      core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before);
    }
  });
});
