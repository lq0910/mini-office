/** 功能：真实核心验证查找位置、替换历史及编号资源的保存重开；时间：2026-11-10；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { WordCoreSession, buildBlankDocx } from '../packages/word-core/src/index';
import { findWordText, highlightWordSearch, replaceWordText, readWordList, changeWordListLevel, setActiveSubEditor, CellSelection, Editor } from '../packages/word-core/src/editing/index';
import { textboxSubExtensions } from '../packages/word-core/src/upstream/apps/docs/src/renderer/editor/extensions';

const sessions: WordCoreSession[] = [];
const body = '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:b/><w:lang w:val="en-US"/></w:rPr><w:t>Ab</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>c abc</w:t></w:r></w:p><w:p><w:r><w:t>ABC 中文 abc</w:t></w:r></w:p>';
/** 功能：构造来源真实的混合字符包，可模拟无编号部件和已占用编号；时间：2026-11-10；作者：lq。 */
async function open(options: { readonly?: boolean; noNumbering?: boolean; body?: string; conflict?: boolean } = {}) {
  const zip = await JSZip.loadAsync(await buildBlankDocx());
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${options.body ?? body}<w:sectPr/></w:body></w:document>`);
  if (options.noNumbering) {
    zip.remove('word/numbering.xml');
    const rels = 'word/_rels/document.xml.rels';
    zip.file(rels, (await zip.file(rels)!.async('text')).replace(/<Relationship\b[^>]*Type="[^"]*\/numbering"[^>]*\/>/g, ''));
  }
  if (options.conflict) {
    const xml = await zip.file('word/numbering.xml')!.async('text');
    zip.file('word/numbering.xml', xml.replace('</w:numbering>', '<w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num></w:numbering>'));
  }
  const core = await WordCoreSession.open(await zip.generateAsync({ type: 'uint8array' }), { readonly: options.readonly });
  sessions.push(core);
  return core;
}
/** 功能：通过真实 DOCX 重新打开结果，不以内存 JSON 冒充持久化验证；时间：2026-11-10；作者：lq。 */
async function reopen(core: WordCoreSession) {
  const core2 = await WordCoreSession.open((await core.prepareSave()).bytes);
  sessions.push(core2);
  return core2;
}
beforeEach(() => {
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => setTimeout(() => callback(performance.now()), 0));
  vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id));
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
});
afterEach(() => { setActiveSubEditor(null); for (const core of sessions.splice(0)) core.dispose(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });

describe('M3 查找替换迁移', () => {
  it('跨字符格式匹配、中文和大小写查询返回正确原生位置，高亮不修改修订和历史', async () => {
    const core = await open();
    const before = core.state();
    const results = findWordText(core.editor, 'abc');
    expect(results.matches).toHaveLength(4);
    expect(results.matches[0]).toMatchObject({ from: 1, to: 4, text: 'Abc' });
    expect(findWordText(core.editor, 'abc', true).matches).toHaveLength(2);
    expect(findWordText(core.editor, '中文').matches).toHaveLength(1);
    highlightWordSearch(core.editor, results, 0);
    // 功能：跨格式命中会拆成多个装饰 span，检查覆盖文字而非错误假定一个命中一个 DOM；时间：2026-11-10；作者：lq。
    expect([...core.editor.view.dom.querySelectorAll('.search-hit')].map(node => node.textContent).join('')).toBe('AbcabcABCabc');
    expect(core.state()).toEqual(before);
    highlightWordSearch(core.editor, findWordText(core.editor, ''), 0);
    expect(core.editor.view.dom.querySelectorAll('.search-hit')).toHaveLength(0);
  });
  it('全部替换一步撤销重做，继承命中首字符格式和语言，保存重开内容正确', async () => {
    const core = await open();
    const originalMarks = core.editor.state.doc.child(0).child(0).marks;
    expect(replaceWordText(core.editor, findWordText(core.editor, 'abc'), '替换')).toBe(4);
    expect(core.editor.state.doc.textContent).toBe('替换 替换替换 中文 替换');
    const first = core.editor.state.doc.child(0).child(0);
    expect(first.marks.some(mark => mark.type.name === 'bold')).toBe(true);
    expect(first.marks).toEqual(originalMarks);
    core.editor.commands.undo();
    expect(core.editor.state.doc.textContent).toBe('Abc abcABC 中文 abc');
    core.editor.commands.redo();
    expect((await reopen(core)).editor.state.doc.textContent).toBe('替换 替换替换 中文 替换');
    expect(await (await JSZip.loadAsync((await core.prepareSave()).bytes)).file('word/document.xml')!.async('text')).toContain('en-US');
  });
  it('单次替换可删除并拒绝过期位置，只读仍可查找且不能替换', async () => {
    const core = await open();
    const results = findWordText(core.editor, 'abc');
    expect(replaceWordText(core.editor, results, '', 1)).toBe(1);
    expect(replaceWordText(core.editor, results, '旧位置')).toBe(0);
    expect(findWordText(core.editor, 'abc').matches).toHaveLength(3);
    const readonly = await open({ readonly: true });
    expect(findWordText(readonly.editor, 'abc').matches).toHaveLength(4);
    expect(replaceWordText(readonly.editor, findWordText(readonly.editor, 'abc'), '禁改')).toBe(0);
    expect(readonly.state().dirty).toBe(false);
  });
  it('按字面处理符号，不跨图片/字段或段落匹配，表格内可替换', async () => {
    const core = await open({ body: '<w:p><w:r><w:t>^p $1 [a] 中文</w:t><w:br/><w:t>后文</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>表格中文</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' });
    expect(findWordText(core.editor, '^p').matches).toHaveLength(1);
    expect(findWordText(core.editor, '[a]').matches).toHaveLength(1);
    expect(findWordText(core.editor, '文后').matches).toHaveLength(0);
    expect(replaceWordText(core.editor, findWordText(core.editor, '中文'), '$1')).toBe(2);
    expect((await reopen(core)).editor.state.doc.textContent).toContain('表格$1');
  });
  it('精简文本框 schema 按需高亮、替换及撤销，列表命令使用区域 marker', async () => {
    const core = await open();
    const sub = new Editor({ extensions: textboxSubExtensions, content: { type: 'doc', content: [{ type: 'docParagraph', content: [{ type: 'text', text: '框内中文' }] }] } });
    try {
      core.editor.view.dom.append(sub.view.dom);
      setActiveSubEditor(sub);
      const results = findWordText(core.commandEditor(), '中文');
      highlightWordSearch(sub, results, 0);
      expect(sub.view.dom.querySelector('.search-hit-active')?.textContent).toBe('中文');
      expect(replaceWordText(sub, results, '替换')).toBe(1);
      expect(sub.state.doc.textContent).toBe('框内替换');
      expect(sub.commands.undo()).toBe(true);
      expect(sub.state.doc.textContent).toBe('框内中文');
      // 功能：精简 schema 保持段落节点并使用编号身份，不误改正文；时间：2026-11-11；作者：lq。
      expect(core.toggleList('bullet')).toBe(true);
      expect(readWordList(sub)).toMatchObject({ kind: 'bullet', level: 0 });
      expect(sub.state.doc.firstChild!.type.name).toBe('docParagraph');
      expect(sub.commands.undo()).toBe(true);
      expect(readWordList(sub).kind).toBeUndefined();
      expect(core.editor.state.doc.textContent).toBe('Abc abcABC 中文 abc');
    } finally { setActiveSubEditor(null); sub.destroy(); }
  });
});

describe('M3 列表迁移', () => {
  it('同一选区列表切换保留来源、字符及段落格式，单步撤销重做并保存真实编号', async () => {
    const core = await open();
    core.editor.commands.selectAll();
    const before = core.editor.getJSON();
    expect(core.toggleList('ordered')).toBe(true);
    expect(readWordList(core.editor)).toMatchObject({ kind: 'ordered', level: 0, mixed: false });
    const first = core.editor.state.doc.child(0);
    expect(first.attrs.align).toBe('right');
    expect(first.attrs.docxIndex).toBe(before.content![0].attrs!.docxIndex);
    const numId = first.attrs.numId;
    expect(core.editor.state.doc.child(1).attrs.numId).toBe(numId);
    core.editor.commands.undo();
    expect(core.editor.getJSON()).toEqual(before);
    core.editor.commands.redo();
    const reopened = await reopen(core);
    reopened.editor.commands.selectAll();
    expect(readWordList(reopened.editor)).toMatchObject({ kind: 'ordered', level: 0 });
    expect(reopened.parsed.numbering.get(numId)?.levels[0].numFmt).toBe('decimal');
    expect(reopened.editor.state.doc.child(0).child(0).marks.some(mark => mark.type.name === 'bold')).toBe(true);
    expect(reopened.toggleList('ordered')).toBe(true);
    expect((await reopen(reopened)).editor.state.doc.child(0).type.name).toBe('docParagraph');
  });
  it('缺少编号部件时创建合法关系，避开导入编号冲突，未使用定义不写入', async () => {
    for (const options of [{ noNumbering: true }, { conflict: true }]) {
      const core = await open(options);
      core.editor.commands.selectAll(); core.toggleList('bullet');
      const numId = core.editor.state.doc.child(0).attrs.numId;
      expect(Number(numId)).toBeGreaterThan(options.conflict ? 3 : 2);
      const saved = await core.prepareSave();
      const zip = await JSZip.loadAsync(saved.bytes);
      expect(await zip.file('word/numbering.xml')!.async('text')).toContain(`w:numId="${numId}"`);
      expect(await zip.file('word/_rels/document.xml.rels')!.async('text')).toContain('/numbering');
      const reopened = await reopen(core);
      expect(reopened.parsed.numbering.get(numId)?.levels[0].lvlText).toBe('•');
      core.editor.commands.undo();
      const reverted = await JSZip.loadAsync((await core.prepareSave()).bytes);
      expect(await reverted.file('word/numbering.xml')?.async('text') ?? '').not.toContain(`w:numId="${numId}"`);
    }
  });
  it('当前层级、类别切换和相邻续号保存重开一致，增减层级可撤销', async () => {
    const core = await open();
    core.editor.commands.setTextSelection(1); core.toggleList('ordered');
    const numId = core.editor.state.doc.child(0).attrs.numId;
    const next = core.editor.state.doc.child(0).nodeSize + 1;
    core.editor.commands.setTextSelection(next); core.toggleList('ordered');
    expect(core.editor.state.doc.child(1).attrs.numId).toBe(numId);
    expect(changeWordListLevel(core.editor, 1)).toBe(true);
    expect(readWordList(core.editor).level).toBe(1);
    core.editor.commands.undo(); expect(readWordList(core.editor).level).toBe(0);
    core.editor.commands.redo();
    core.toggleList('bullet');
    const reopened = await reopen(core);
    reopened.editor.commands.setTextSelection(next);
    expect(readWordList(reopened.editor)).toMatchObject({ kind: 'bullet', level: 1 });
    expect(reopened.parsed.numbering.get(reopened.editor.state.doc.child(1).attrs.numId)?.levels[1].lvlText).toBe('•');
  });
  it('只读和中文组合输入拒绝列表及替换', async () => {
    const core = await open();
    const readonly = await open({ readonly: true });
    expect(readonly.toggleList('ordered')).toBe(false);
    expect(changeWordListLevel(readonly.editor, 1)).toBe(false);
    Object.defineProperty(core.editor.view, 'composing', { configurable: true, value: true });
    expect(core.toggleList('bullet')).toBe(false);
    expect(replaceWordText(core.editor, findWordText(core.editor, 'abc'), '未确认')).toBe(0);
    expect(core.state().dirty).toBe(false);
  });
  it('表格矩形选区切换列表保留格结构，不误改表后段落', async () => {
    const core = await open({ body: '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>甲格</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p><w:r><w:t>乙格</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>表后</w:t></w:r></w:p>' });
    // 功能：核心会将非法跨格 TextSelection 收回单格，矩形操作必须使用真实 CellSelection；时间：2026-11-10；作者：lq。
    const cells: number[] = [];
    core.editor.state.doc.descendants((node, pos) => { if (node.type.name === 'docTableCell') cells.push(pos); });
    core.editor.view.dispatch(core.editor.state.tr.setSelection(CellSelection.create(core.editor.state.doc, cells[0], cells[1])));
    core.toggleList('ordered');
    expect(core.editor.state.doc.lastChild!.type.name).toBe('docParagraph');
    const reopened = await reopen(core);
    const lists: string[] = [];
    reopened.editor.state.doc.descendants(node => { if (node.type.name === 'docListItem') lists.push(node.textContent); });
    expect(lists).toEqual(['甲格', '乙格']);
    expect(reopened.editor.state.doc.firstChild!.child(0).childCount).toBe(2);
    expect(reopened.editor.state.doc.lastChild!.textContent).toBe('表后');
  });
});
