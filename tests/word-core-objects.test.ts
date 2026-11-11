/** 功能：验证图片和表格真实命令、矩形选区、单步历史及 OOXML 保存重开；时间：2026-11-11；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { WordCoreSession, buildBlankDocx } from '../packages/word-core/src/index';
import { CellSelection } from '../packages/word-core/src/editing/index';
import { applyWordImage, applyWordTableProperties, insertWordPicture, insertWordPictures, insertWordTable, runWordTableAction, wordContentWidth, wordTableState, type WordTableAction } from '../packages/word-core/src/editing/objects';
import { toggleWordInlineFormat } from '../packages/word-core/src/editing/formatting';
import { patchWordCellParagraph } from '../packages/word-core/src/serialization/table-paragraph';

const cores: WordCoreSession[] = [];
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Zl1sAAAAASUVORK5CYII=';
/** 功能：原生会话和真实 DOCX 构建，不以旧 DocModel 代替操作目标；时间：2026-11-11；作者：lq。 */
async function create(readonly = false) { const core = await WordCoreSession.create({ readonly }); cores.push(core); return core; }
/** 功能：保存件重新解析，以序列化结果验证关系和属性；时间：2026-11-11；作者：lq。 */
async function reopen(core: WordCoreSession) { const next = await WordCoreSession.open((await core.prepareSave()).bytes); cores.push(next); return next; }
/** 功能：找到当前网格单元格内部的位置，兼容合并格；时间：2026-11-11；作者：lq。 */
function cell(core: WordCoreSession, row = 0, column = 0) {
  const table = core.editor.state.doc.content.content.find(node => node.type.name === 'docTable')!;
  let at = 0; core.editor.state.doc.forEach((node, pos) => { if (node === table) at = pos; });
  let r = 0; table.forEach((item, offset, index) => { if (index === row) r = at + 1 + offset; });
  let c = 0; table.child(row).forEach((_item, offset, index) => { if (index === column) c = r + 1 + offset; });
  core.editor.commands.setTextSelection(c + 2); return c;
}
/** 功能：检查表格保存的 XML，而非只看运行时摘要；时间：2026-11-11；作者：lq。 */
async function xml(core: WordCoreSession) { return (await JSZip.loadAsync((await core.prepareSave()).bytes)).file('word/document.xml')!.async('string'); }
/** 功能：含未知表格/行/格属性及段落书签的独立公开测试包；时间：2026-11-11；作者：lq。 */
async function imported() {
  const zip = await JSZip.loadAsync(await buildBlankDocx());
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCaption w:val="保留表说明"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${[0, 1].map(row => `<w:tr><w:trPr><w:cantSplit/></w:trPr>${[0, 1].map(col => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:noWrap/><w:hideMark/></w:tcPr><w:p><w:pPr><w:suppressLineNumbers/></w:pPr><w:bookmarkStart w:id="${row * 2 + col}" w:name="keep${row}${col}"/><w:r><w:rPr><w:lang w:val="zh-CN"/></w:rPr><w:t>格${row}${col}</w:t></w:r><w:bookmarkEnd w:id="${row * 2 + col}"/></w:p></w:tc>`).join('')}</w:tr>`).join('')}</w:tbl><w:p><w:r><w:t>表外正文</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const core = await WordCoreSession.open(await zip.generateAsync({ type: 'uint8array' })); cores.push(core); return core;
}
beforeEach(() => {
  vi.stubGlobal('requestAnimationFrame', (fn: FrameRequestCallback) => setTimeout(() => fn(performance.now()), 0));
  vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id));
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
});
afterEach(() => { for (const core of cores.splice(0)) core.dispose(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });

describe('M3 图片与表格', () => {
  // 功能：格式和结构同时变化时验证真实导入段落的未知属性及注释边界；时间：2026-11-11；作者：lq。
  it('格内局部格式与表结构同时修改仍保留书签、语言和未知段落属性', async () => {
    const core = await imported(); const c = cell(core);
    core.editor.commands.setTextSelection({ from: c + 2, to: c + 3 });
    expect(toggleWordInlineFormat(core.editor, core.parsed, 'bold')).toBe(true);
    expect(applyWordTableProperties(core.editor, { widthPx: 380 })).toBe(true);
    expect(runWordTableAction(core.editor, 'row-above')).toBe(true);
    const output = await xml(core);
    expect(output).toMatch(/bookmarkStart[^>]*keep00[\s\S]*?<w:b\s*\/>[\s\S]*?格[\s\S]*?00[\s\S]*?bookmarkEnd/);
    for (const token of ['w:suppressLineNumbers', 'zh-CN', '保留表说明']) expect(output).toContain(token);
    const saved = await reopen(core); expect(saved.editor.state.doc.textContent).toContain('格00');
    expect(await xml(saved)).toContain('keep00');
  });
  // 功能：覆盖联合文字格式修改、空段首次输入和编号属性移除的 XML 边界；时间：2026-11-11；作者：lq。
  it('联合文字格式补写保留追加文字、空段注释且按规范排列属性', () => {
    const source = { xml: '<w:p w:rsidR="123"><w:pPr><w:pStyle w:val="Old"/><w:numPr><w:numId w:val="1"/></w:numPr><w:suppressLineNumbers/></w:pPr><w:bookmarkStart w:id="1"/><w:r><w:t>原文</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>', styleId: 'Old' };
    const result = patchWordCellParagraph(source, { align: 'center' }, [{ text: '原文追加', bold: true }], 'New', { numId: '3', ilvl: 1 })!;
    expect(result).toContain('原文追加'); expect(result).toContain('w:rsidR="123"');
    expect(result.indexOf('<w:pStyle')).toBeLessThan(result.indexOf('<w:numPr'));
    expect(result.indexOf('<w:numPr')).toBeLessThan(result.indexOf('<w:suppressLineNumbers'));
    expect(result.indexOf('原文追加')).toBeLessThan(result.indexOf('<w:bookmarkEnd'));
    expect(patchWordCellParagraph(source, undefined, [{ text: '删', italic: true }], null)).not.toContain('w:numPr');
    const empty = { xml: '<w:p><w:bookmarkStart w:id="2"/><w:bookmarkEnd w:id="2"/></w:p>' };
    expect(patchWordCellParagraph(empty, undefined, [{ text: '首次', bold: true }], null)).toMatch(/bookmarkStart[\s\S]*首次[\s\S]*bookmarkEnd/);
  });
  it('插表保留后续空段、首格光标，撤销重做和保存重开保持网格', async () => {
    const core = await create(); const before = core.editor.getJSON();
    expect(insertWordTable(core.editor, 3, 2)).toBe(true);
    expect(core.editor.state.doc.lastChild!.type.name).toBe('docParagraph');
    expect(core.editor.state.selection.$from.parent.type.name).toBe('docParagraph');
    expect(wordTableState(core.editor)).toMatchObject({ rows: 3, columns: 2 });
    core.editor.commands.insertContent('表格中文'); core.editor.commands.undo(); core.editor.commands.undo();
    expect(core.editor.getJSON()).toEqual(before); core.editor.commands.redo();
    const saved = await reopen(core); cell(saved); expect(wordTableState(saved.editor)).toMatchObject({ rows: 3, columns: 2 });
    expect(await xml(core)).toContain('w:tblBorders');
  });
  it('所有行列增删分别一步撤销，保存重开可继续操作，最后一列允许删除整表', async () => {
    const core = await imported(); cell(core);
    for (const [action, rows, columns] of [['row-above', 3, 2], ['row-below', 3, 2], ['column-before', 2, 3], ['column-after', 2, 3], ['row-delete', 1, 2], ['column-delete', 2, 1]] as const) {
      const before = core.editor.getJSON(); expect(runWordTableAction(core.editor, action)).toBe(true);
      expect(wordTableState(core.editor)).toMatchObject({ rows, columns });
      const saved = await reopen(core); cell(saved); expect(wordTableState(saved.editor)).toMatchObject({ rows, columns });
      core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before); cell(core);
    }
    const single = await create(); insertWordTable(single.editor, 2, 1); expect(runWordTableAction(single.editor, 'column-delete')).toBe(true);
    expect(single.editor.state.doc.content.content.every(node => node.type.name !== 'docTable')).toBe(true);
    expect((await reopen(single)).editor.state.doc.content.content.every(node => node.type.name !== 'docTable')).toBe(true);
  });
  it('真实矩形合并、跨行合并保存 vMerge，拆分后文字不丢失', async () => {
    const core = await imported(); const a = cell(core, 0, 0); const b = cell(core, 1, 0);
    core.editor.view.dispatch(core.editor.state.tr.setSelection(CellSelection.create(core.editor.state.doc, a, b)));
    const before = core.editor.getJSON(); expect(runWordTableAction(core.editor, 'merge')).toBe(true);
    expect(core.editor.state.doc.firstChild!.firstChild!.firstChild!.attrs.rowspan).toBe(2);
    expect(await xml(core)).toContain('w:vMerge');
    const saved = await reopen(core); cell(saved); expect(wordTableState(saved.editor)?.cell.attrs.rowspan).toBe(2);
    expect(runWordTableAction(saved.editor, 'split')).toBe(true); expect(await xml(saved)).not.toContain('w:vMerge w:val="restart"');
    expect((await reopen(saved)).editor.state.doc.textContent).toContain('格00');
    core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before);
  });
  it('属性一次事务提交，保存保留非目标语言、书签、未知段格行属性', async () => {
    const core = await imported(); cell(core); const before = core.editor.getJSON();
    expect(applyWordTableProperties(core.editor, { widthPx: 360, align: 'center', columnWidthPx: 150, rowHeightTwips: 540, rowHeightRule: 'exact', fill: 'FFFF00', vAlign: 'bottom', border: { mode: 'all', color: '123456', widthPt: 1.5, style: 'double' } })).toBe(true);
    const saved = await reopen(core); cell(saved);
    expect(wordTableState(saved.editor)?.table.attrs).toMatchObject({ widthPx: 360, tblAlign: 'center' });
    expect(wordTableState(saved.editor)?.cell.attrs).toMatchObject({ fill: 'FFFF00', vAlign: 'bottom', borders: expect.objectContaining({ top: expect.objectContaining({ color: '123456', style: 'double' }) }) });
    expect(wordTableState(saved.editor)?.row.attrs).toMatchObject({ heightTwips: 540, heightRule: 'exact' });
    const output = await xml(core); for (const token of ['保留表说明', 'w:noWrap', 'w:hideMark', 'w:cantSplit', 'w:suppressLineNumbers', 'keep00', 'zh-CN']) expect(output).toContain(token);
    core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before); core.editor.commands.redo(); expect(await xml(core)).toContain('FFFF00');
  });
  it('整表/行/列选择不改 dirty 或历史，整表属性按全部格生效', async () => {
    const core = await imported(); cell(core); const revision = core.state().revision;
    for (const action of ['select-rows', 'select-columns', 'select-table'] as WordTableAction[]) { expect(runWordTableAction(core.editor, action)).toBe(true); expect(core.state().revision).toBe(revision); }
    expect(wordTableState(core.editor)).toMatchObject({ rows: 2, columns: 2 }); expect(core.editor.state.selection.toJSON().type).toBe('node');
    applyWordTableProperties(core.editor, { fill: '112233' });
    const saved = await reopen(core); saved.editor.state.doc.firstChild!.forEach(row => row.forEach(cell => expect(cell.attrs.fill).toBe('112233')));
    expect(core.editor.commands.undo()).toBe(true); expect(core.state().dirty).toBe(true);
  });
  it('重复表头、平均列宽、适应页宽分别单步历史并持久化', async () => {
    const core = await imported(); cell(core); expect(runWordTableAction(core.editor, 'header')).toBe(true);
    expect(wordTableState(core.editor)?.header.active).toBe(true);
    expect((await xml(core))).toContain('w:tblHeader');
    const saved = await reopen(core); cell(saved); expect(wordTableState(saved.editor)?.header.active).toBe(true);
    const before = core.editor.getJSON(); runWordTableAction(core.editor, 'distribute-columns'); core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before);
    expect(runWordTableAction(core.editor, 'fit-page')).toBe(true); expect(await xml(core)).toContain('w:type="pct"');
    cell(core, 1); expect(runWordTableAction(core.editor, 'header')).toBe(false);
  });
  // 功能：多个分节的表格通过编辑器装饰获得各自净宽，布局尺寸不能混用第一页；时间：2026-11-11；作者：lq。
  it('不同纸张分节的宽度装饰随同一页面事务更新，表格适应当前节净宽', async () => {
    const base = await imported(); const zip = await JSZip.loadAsync((await base.prepareSave()).bytes);
    const source = await zip.file('word/document.xml')!.async('string'); const table = source.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)![0];
    zip.file('word/document.xml', source.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${table}<w:p><w:pPr><w:sectPr><w:pgSz w:w="10000" w:h="16000"/><w:pgMar w:left="1000" w:right="1000" w:top="1000" w:bottom="1000"/></w:sectPr></w:pPr></w:p>${table}<w:p/><w:sectPr><w:pgSz w:w="16000" w:h="18000"/><w:pgMar w:left="2000" w:right="2000" w:top="1000" w:bottom="1000"/></w:sectPr></w:body>`));
    const core = await WordCoreSession.open(await zip.generateAsync({ type: 'uint8array' })); cores.push(core);
    const tables: number[] = []; core.editor.state.doc.forEach((node, at) => { if (node.type.name === 'docTable') tables.push(at); });
    core.editor.commands.setTextSelection(tables[0] + 4); expect(wordContentWidth(core.editor)).toBeCloseTo(8000 / 15);
    core.editor.commands.setTextSelection(tables[1] + 4); expect(wordContentWidth(core.editor)).toBe(800);
    expect(runWordTableAction(core.editor, 'fit-page')).toBe(true);
    expect(wordTableState(core.editor)!.table.attrs).toMatchObject({ widthPx: null, widthPct: 100, tblAutoFit: 'window' });
    const section = core.sections()[1]; const before = core.pageLayout();
    core.commitPageLayout({ ...before, sections: { ...before.sections, [section.lastBlockIndex]: { marginLeft: 3500 } } });
    expect(wordContentWidth(core.editor)).toBe(700); core.editor.commands.undo(); expect(wordContentWidth(core.editor)).toBe(800);
  });
  it('导入纵向合并的隐藏续格在属性变动及前方插行插列后仍保留原行身份与 XML', async () => {
    const fixture = await imported(); const zip = await JSZip.loadAsync((await fixture.prepareSave()).bytes);
    const original = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', original.replace('<w:tcW w:w="3000" w:type="dxa"/>', '<w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/>')
      .replace(/(<w:tr><w:trPr><w:cantSplit\/><\/w:trPr><w:tc><w:tcPr>)(<w:tcW[^>]*\/>)(<w:noWrap\/><w:hideMark\/><\/w:tcPr><w:p>[\s\S]*?格10)/, '$1$2<w:vMerge/>$3'));
    const core = await WordCoreSession.open(await zip.generateAsync({ type: 'uint8array' })); cores.push(core); cell(core);
    expect(wordTableState(core.editor)?.cell.attrs.rowspan).toBe(2);
    applyWordTableProperties(core.editor, { fill: 'FAFA00', widthPx: 360 });
    for (const action of [undefined, 'row-above', 'column-before'] as const) {
      if (action) { cell(core); expect(runWordTableAction(core.editor, action)).toBe(true); }
      const output = await xml(core);
      for (const token of ['格10', 'keep10', 'w:vMerge', 'w:noWrap', 'w:hideMark']) expect(output).toContain(token);
      const saved = await reopen(core); expect(await xml(saved)).toContain('格10');
    }
    // 功能：前方插行插列后拆分，隐藏文字变成独立续行单元格且仍可撤销；时间：2026-11-11；作者：lq。
    cell(core, 1, 1); const before = core.editor.getJSON();
    expect(runWordTableAction(core.editor, 'split')).toBe(true);
    expect(core.editor.state.doc.textContent).toContain('格10');
    const splitXml = await xml(core); for (const token of ['格10', 'keep10', 'zh-CN', 'w:hideMark']) expect(splitXml).toContain(token);
    expect((await reopen(core)).editor.state.doc.textContent).toContain('格10');
    core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before);
  });
  it('正文图片插入、环绕、尺寸及删除可撤销，真实图片部件与关系保持', async () => {
    const core = await create(); const before = core.editor.getJSON();
    expect(insertWordPicture(core.editor, { dataUrl: `data:image/png;base64,${PNG}`, base64: PNG, mime: 'image/png', widthPx: 160, heightPx: 80, label: '图片测试' })).toBe(true);
    expect(core.editor.state.doc.lastChild!.type.name).toBe('docParagraph'); core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before); core.editor.commands.redo();
    let pos = 0; core.editor.state.doc.forEach((node, at) => { if (node.attrs.blockType === 'image') pos = at; }); core.editor.commands.setNodeSelection(pos);
    expect(applyWordImage(core.editor, { wrap: 'square-right', widthPx: 200, heightPx: 100 })).toBe(true);
    const saved = await reopen(core); const img = saved.editor.state.doc.content.content.find(node => node.attrs.blockType === 'image')!;
    expect(img.attrs).toMatchObject({ imageWrap: 'square-right', imageWidthPx: 200, imageHeightPx: 100 });
    const zip = await JSZip.loadAsync((await core.prepareSave()).bytes); expect(Object.keys(zip.files).some(path => /^word\/media\//.test(path))).toBe(true);
    expect(await xml(core)).toContain('wp:anchor'); expect(await xml(core)).toContain('1905000');
    core.editor.commands.setNodeSelection(pos); const old = core.editor.getJSON(); expect(applyWordImage(core.editor, { delete: true })).toBe(true); core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(old);
    applyWordImage(core.editor, { wrap: 'inline' }); expect(await xml(core)).toContain('wp:inline');
  });
  // 功能：剪贴板多图片一次入历史，尾段可继续打字，非法图片不会部分提交；时间：2026-11-11；作者：lq。
  it('多张图片在正文和单元格一次插入、撤销及保存，尾段继续输入不替换图片', async () => {
    const picture = { dataUrl: `data:image/png;base64,${PNG}`, base64: PNG, mime: 'image/png', widthPx: 100, heightPx: 50, label: '剪贴板' };
    for (const inTable of [false, true]) {
      const core = await create(); if (inTable) insertWordTable(core.editor, 2, 2);
      const before = core.editor.getJSON(); expect(insertWordPictures(core.editor, [picture, { ...picture, widthPx: 90 }])).toBe(true);
      let count = 0; core.editor.state.doc.descendants(node => { if (node.type.name === 'docInlineImage' || node.attrs.blockType === 'image') count++; }); expect(count).toBe(2);
      core.editor.commands.insertContent('图片后输入'); core.editor.commands.undo(); core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before);
      core.editor.commands.redo(); const saved = await reopen(core); let images = 0; saved.editor.state.doc.descendants(node => { if (node.type.name === 'docInlineImage' || node.attrs.blockType === 'image') images++; }); expect(images).toBe(2);
      const doc = core.editor.getJSON(); expect(insertWordPictures(core.editor, [picture, { ...picture, widthPx: NaN }])).toBe(false); expect(core.editor.getJSON()).toEqual(doc);
    }
  });
  it('格内图片使用内联 run，调整尺寸/环绕保存重开正确且不破坏表格', async () => {
    const core = await create(); insertWordTable(core.editor, 2, 2); cell(core);
    expect(insertWordPicture(core.editor, { dataUrl: `data:image/png;base64,${PNG}`, base64: PNG, mime: 'image/png', widthPx: 80, heightPx: 40, label: '格内图片' })).toBe(true);
    let pos = 0; core.editor.state.doc.descendants((node, at) => { if (node.type.name === 'docInlineImage') pos = at; }); core.editor.commands.setNodeSelection(pos);
    expect(applyWordImage(core.editor, { widthPx: 60, heightPx: 30, wrap: 'topBottom' })).toBe(true);
    const saved = await reopen(core); let image: Record<string, unknown> | undefined; saved.editor.state.doc.descendants(node => { if (node.type.name === 'docInlineImage') image = node.attrs; });
    expect(image).toMatchObject({ widthPx: 60, heightPx: 30, wrap: 'topBottom' }); cell(saved); expect(wordTableState(saved.editor)).toMatchObject({ rows: 2, columns: 2 });
  });
  it('非法参数、只读与组合输入拒绝修改，属性校验失败不部分提交', async () => {
    const core = await imported(); cell(core); const before = core.editor.getJSON(); const revision = core.state().revision;
    expect(applyWordTableProperties(core.editor, { fill: 'FF0000', widthPx: NaN })).toBe(false); expect(insertWordTable(core.editor, 0, 2)).toBe(false);
    expect(core.editor.getJSON()).toEqual(before); expect(core.state().revision).toBe(revision);
    Object.defineProperty(core.editor.view, 'composing', { configurable: true, get: () => true }); expect(runWordTableAction(core.editor, 'row-above')).toBe(false);
    const locked = await create(true); expect(insertWordTable(locked.editor, 2, 2)).toBe(false); expect(locked.state().dirty).toBe(false);
  });
});
