/** 功能：验证真实 OOXML 表格属性在输入、格式修改和连续保存后仍完整保留；时间：2026-11-07；作者：lq。 */
import { describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { createBlankDoc, getSourceBytes, registerFormatHandler, useWorkspaceStore, type DocModel, type DocParagraph, type DocTable } from '@mini-office/document-model';
import { setPlatformAdapter, type PlatformAdapter } from '@mini-office/platform-adapter';
import { parseDocx } from '../packages/editor-word/src/parse';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import { renderParagraphHtml } from '../packages/editor-word/src/html';
import { serializeDomRuns } from '../packages/editor-word/src/editing';
import { wordFormatHandler } from '../packages/editor-word/src/handler';

/** 功能：构造续格边框与首格不同、含未知属性和书签的合并表格；时间：2026-11-07；作者：lq。 */
const paragraph = '<w:p w:rsidR="12345678"><w:pPr><w:spacing w:after="0"/><w:wordWrap w:val="0"/></w:pPr><w:bookmarkStart w:id="1" w:name="keep"/><w:r><w:rPr><w:lang w:val="zh-CN"/></w:rPr><w:t>建设内容</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>';
/** 功能：独立保存合并续格的边框、底纹与占位段，作为逐字不变断言；时间：2026-11-07；作者：lq。 */
const continuation = '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:gridSpan w:val="2"/><w:vMerge/><w:tcBorders><w:left w:val="single" w:sz="12" w:color="FF0000"/><w:bottom w:val="double" w:sz="8"/></w:tcBorders><w:shd w:fill="DDDDDD"/><w:noWrap/></w:tcPr><w:p w:rsidR="CONTINUE"><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr></w:r></w:p></w:tc>';
/** 功能：保留表格、行、格的未建模属性以检测重建造成的格式损失；时间：2026-11-07；作者：lq。 */
const table = `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders><w:insideH w:val="nil"/></w:tblBorders><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr w:rsidR="ROW"><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/><w:tcBorders><w:top w:val="single" w:sz="4"/></w:tcBorders><w:noWrap/></w:tcPr>${paragraph}<w:p><w:r><w:t>保持原文</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>右侧</w:t></w:r></w:p></w:tc></w:tr><w:tr>${continuation}<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>末格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

/** 功能：复用公开样本的有效包结构，私人文件不进入仓库；时间：2026-11-07；作者：lq。 */
async function open(body = table) {
  const zip = await JSZip.loadAsync(readFileSync('tests/fixtures/docx/contract.docx'));
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const doc = await parseDocx(bytes);
  const parsed = doc.blocks.find(block => block.type === 'table') as DocTable;
  return { doc, bytes, table: parsed, paragraph: parsed.rows[0].cells[0].blocks[0] as DocParagraph };
}

/** 功能：直接读取保存 XML，而不是只比较模型能识别的内容；时间：2026-11-07；作者：lq。 */
async function xml(bytes: Uint8Array) { return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('text'); }

/** 功能：从原始 ZIP 检查全部非目标部件，防止旁路保存污染样式或关系；时间：2026-11-07；作者：lq。 */
async function unchangedParts(before: Uint8Array, after: Uint8Array) {
  const original = await JSZip.loadAsync(before);
  const saved = await JSZip.loadAsync(after);
  expect(Object.keys(saved.files).sort()).toEqual(Object.keys(original.files).sort());
  for (const [name, entry] of Object.entries(original.files)) {
    if (entry.dir || name === 'word/document.xml') continue;
    expect(await saved.file(name)!.async('uint8array'), name).toEqual(await entry.async('uint8array'));
  }
}

describe('表格原 XML 安全保存', () => {
  it('单字编辑只改变目标文字，完整保留合并续格和其他 ZIP 部件', async () => {
    const sample = await open();
    sample.paragraph.runs[0].text = '建设内容测';
    const saved = await serializeDocx(sample.doc, { sourceBytes: sample.bytes, name: 'table.docx' });
    expect(await xml(saved)).toBe((await xml(sample.bytes)).replace('<w:t>建设内容</w:t>', '<w:t xml:space="preserve">建设内容测</w:t>'));
    expect(await xml(saved)).toContain(continuation);
    await unchangedParts(sample.bytes, saved);
  });

  it('实际 HTML 回读后输入汉字、粘贴多行和空段新增仍保留表格属性', async () => {
    const sample = await open();
    const root = document.createElement('div');
    root.innerHTML = renderParagraphHtml(sample.paragraph);
    root.querySelector('span')!.textContent = '中文输入及粘贴';
    sample.paragraph.runs = serializeDomRuns(root, sample.paragraph.runs[0].style, new Map());
    const saved = await serializeDocx(sample.doc, { sourceBytes: sample.bytes, name: 'table.docx' });
    expect(await xml(saved)).toBe((await xml(sample.bytes)).replace('<w:t>建设内容</w:t>', '<w:t xml:space="preserve">中文输入及粘贴</w:t>'));
    // 功能：回车由编辑器生成新段落，验证首段之前插入时 tcPr 仍位于内容前；时间：2026-11-07；作者：lq。
    sample.table.rows[0].cells[0].blocks.unshift({ type: 'paragraph', id: 'new-paragraph', style: {}, runs: [{ text: '粘贴内容', style: {} }] });
    const twice = await serializeDocx(sample.doc, { sourceBytes: saved, name: 'table.docx' });
    expect(await xml(twice)).toContain(continuation);
    const dom = new DOMParser().parseFromString(await xml(twice), 'application/xml');
    expect(dom.getElementsByTagName('w:tc')[0].firstElementChild!.tagName).toBe('w:tcPr');
    expect((await parseDocx(twice)).blocks[0].type).toBe('table');
    // 功能：模拟浏览器多行粘贴和 Shift+Enter 形成的 BR、Tab 节点，再次保存不退回整表生成；时间：2026-11-07；作者：lq。
    root.innerHTML = '<span>第一行<br>第二行</span><span data-tab="1"></span><span>末尾</span>';
    sample.paragraph.runs = serializeDomRuns(root, sample.paragraph.runs[0].style, new Map());
    const pasted = await serializeDocx(sample.doc, { sourceBytes: twice, name: 'table.docx' });
    expect(await xml(pasted)).toContain(continuation);
    expect(await xml(pasted)).toContain('<w:br/>');
    expect(await xml(pasted)).toContain('<w:tab/>');
  });

  it('同格式 run 被浏览器合并后仍保留源 run 格式和未知属性', async () => {
    const body = table.replace('<w:t>建设内容</w:t>', '<w:t>建设</w:t></w:r><w:r><w:rPr><w:lang w:val="zh-CN"/></w:rPr><w:t>内容</w:t>');
    const sample = await open(body);
    sample.paragraph.runs = [{ text: '建设内容测', style: sample.paragraph.runs[0].style }];
    const saved = await serializeDocx(sample.doc, { sourceBytes: sample.bytes, name: 'table.docx' });
    expect(await xml(saved)).toBe((await xml(sample.bytes)).replace('<w:t>内容</w:t>', '<w:t xml:space="preserve">内容测</w:t>'));
  });

  it('字号、字体、加粗和段落对齐按属性组修改，保留原书签、语言和行格属性', async () => {
    const sample = await open();
    sample.paragraph.runs[0].style = { ...sample.paragraph.runs[0].style, bold: true, fontSizePt: 16, fontFamily: '宋体' };
    sample.paragraph.style.alignment = 'center';
    const saved = await serializeDocx(sample.doc, { sourceBytes: sample.bytes, name: 'table.docx' });
    const value = await xml(saved);
    for (const fragment of [continuation, '<w:lang w:val="zh-CN"/>', '<w:wordWrap w:val="0"/>', '<w:bookmarkStart w:id="1" w:name="keep"/>', '<w:cantSplit/>', '<w:tblLook w:val="04A0"/>']) expect(value).toContain(fragment);
    const reopened = (await parseDocx(saved)).blocks[0] as DocTable;
    expect((reopened.rows[0].cells[0].blocks[0] as DocParagraph).runs[0].style).toMatchObject({ bold: true, fontSizePt: 16, fontFamily: '宋体' });
    expect((reopened.rows[0].cells[0].blocks[0] as DocParagraph).style.alignment).toBe('center');
  });

  it('修改单元格底纹和表格对齐不会重建续格', async () => {
    const sample = await open();
    sample.table.rows[0].cells[0].backgroundColor = '#123456';
    sample.table.alignment = 'center';
    const saved = await serializeDocx(sample.doc, { sourceBytes: sample.bytes, name: 'table.docx' });
    expect(await xml(saved)).toContain(continuation);
    expect(await xml(saved)).toContain('<w:tblW w:w="6000" w:type="dxa"/>');
    const parsed = (await parseDocx(saved)).blocks[0] as DocTable;
    expect(parsed.alignment).toBe('center');
    expect(parsed.rows[0].cells[0].backgroundColor).toBe('#123456');
  });

  it('修改行距保留同组自动段距与字符缩进，空段字号保留标记语言', async () => {
    const sample = await open(table.replace('<w:spacing w:after="0"/>', '<w:spacing w:beforeLines="100" w:after="0" w:afterAutospacing="1"/><w:ind w:leftChars="100" w:firstLineChars="200"/>'));
    sample.paragraph.style.lineHeight = 1.5;
    const saved = await serializeDocx(sample.doc, { sourceBytes: sample.bytes, name: 'table.docx' });
    expect(await xml(saved)).toContain('w:beforeLines="100"');
    expect(await xml(saved)).toContain('w:afterAutospacing="1"');
    expect(await xml(saved)).toContain('<w:ind w:leftChars="100" w:firstLineChars="200"/>');
    const empty = await open(table.replace(paragraph, '<w:p><w:pPr><w:rPr><w:sz w:val="24"/><w:lang w:val="zh-CN"/></w:rPr></w:pPr></w:p>'));
    empty.paragraph.runs[0].style.fontSizePt = 18;
    const changed = await serializeDocx(empty.doc, { sourceBytes: empty.bytes, name: 'empty.docx' });
    expect(await xml(changed)).toContain('<w:lang w:val="zh-CN"/>');
    const parsed = (await parseDocx(changed)).blocks[0] as DocTable;
    expect((parsed.rows[0].cells[0].blocks[0] as DocParagraph).markStyle?.fontSizePt).toBe(18);
  });

  it('新建文档多次保存仍可编辑，来源定位失败的导入文档明确拒绝保存', async () => {
    const doc = createBlankDoc();
    let bytes = await serializeDocx(doc, { name: 'new.docx' });
    (doc.blocks[0] as DocParagraph).runs[0].text = '再次保存';
    bytes = await serializeDocx(doc, { sourceBytes: bytes, name: 'new.docx' });
    expect(await xml(bytes)).toContain('再次保存');
    const imported = await open();
    imported.doc.source = undefined;
    imported.doc.sourceMappingFailed = true;
    await expect(serializeDocx(imported.doc, { sourceBytes: imported.bytes, name: 'unsafe.docx' })).rejects.toMatchObject({ code: 'SAVE_FAILED' });
  });

  it('嵌套表格文字编辑只替换目标文字', async () => {
    const nested = '<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:noWrap/></w:tcPr><w:p><w:r><w:t>内表</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const sample = await open(table.replace('<w:p><w:r><w:t>右侧</w:t></w:r></w:p>', nested + '<w:p/>'));
    const inner = sample.table.rows[0].cells[1].blocks[0] as DocTable;
    (inner.rows[0].cells[0].blocks[0] as DocParagraph).runs[0].text = '内表修改';
    const saved = await serializeDocx(sample.doc, { sourceBytes: sample.bytes, name: 'table.docx' });
    expect(await xml(saved)).toBe((await xml(sample.bytes)).replace('<w:t>内表</w:t>', '<w:t xml:space="preserve">内表修改</w:t>'));
  });

  it('连续保存、重开再编辑及撤销不漂移', async () => {
    const sample = await open();
    let bytes = sample.bytes;
    for (const text of ['第一次', '第二次', '建设内容']) {
      sample.paragraph.runs[0].text = text;
      bytes = await serializeDocx(sample.doc, { sourceBytes: bytes, name: 'table.docx' });
      expect(await xml(bytes)).toBe(text === '建设内容' ? await xml(sample.bytes) : (await xml(sample.bytes)).replace('<w:t>建设内容</w:t>', `<w:t xml:space="preserve">${text}</w:t>`));
    }
    const reopened = await parseDocx(bytes);
    ((reopened.blocks[0] as DocTable).rows[0].cells[0].blocks[0] as DocParagraph).runs[0].text += '重开';
    const saved = await serializeDocx(reopened, { sourceBytes: bytes, name: 'table.docx' });
    expect(await xml(saved)).toContain(continuation);
  });

  it('调整页面后仍保留原表，不能因整体外壳重建而丢失续格', async () => {
    const sample = await open();
    sample.doc.page.margins.left += 12;
    sample.paragraph.runs[0].text += '测';
    const saved = await serializeDocx(sample.doc, { sourceBytes: sample.bytes, name: 'table.docx' });
    expect(await xml(saved)).toContain(continuation);
    expect(await xml(saved)).toContain('<w:tblLook w:val="04A0"/>');
  });

  it('含原始属性的分节段落在重建外壳后仍有效，节属性不重复', async () => {
    const sectionParagraph = '<w:p w:rsidR="SECTION"><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>第一节</w:t></w:r></w:p>';
    const sample = await open(sectionParagraph + table);
    sample.doc.page.margins.left += 12;
    const saved = await serializeDocx(sample.doc, { sourceBytes: sample.bytes, name: 'sections.docx' });
    const dom = new DOMParser().parseFromString(await xml(saved), 'application/xml');
    expect(dom.getElementsByTagName('parsererror')).toHaveLength(0);
    const first = dom.getElementsByTagName('w:p')[0];
    expect(first.getAttribute('w:rsidR')).toBe('SECTION');
    expect(first.getElementsByTagName('w:sectPr')).toHaveLength(1);
    expect(await xml(saved)).toContain(continuation);
    expect((await parseDocx(saved)).blocks.some(block => block.type === 'table')).toBe(true);
  });

  it('源表的未支持结构修改显式失败，保留原字节', async () => {
    const sample = await open();
    sample.table.rows.splice(1, 1);
    await expect(serializeDocx(sample.doc, { sourceBytes: sample.bytes, name: 'table.docx' })).rejects.toMatchObject({ code: 'SAVE_FAILED' });
    expect(await xml(sample.bytes)).toContain(continuation);
  });

  it('跨硬分页的复合源段落不能因其他设置修改而被静默重建', async () => {
    const composite = '<w:p w:rsidR="COMPOSITE"><w:r><w:t>分页前</w:t><w:br w:type="page"/><w:t>分页后</w:t></w:r></w:p>';
    const sample = await open(composite + table);
    sample.paragraph.runs[0].text += '测';
    const saved = await serializeDocx(sample.doc, { sourceBytes: sample.bytes, name: 'composite.docx' });
    expect(await xml(saved)).toContain(composite);
    sample.doc.page.margins.left += 12;
    await expect(serializeDocx(sample.doc, { sourceBytes: saved, name: 'composite.docx' })).rejects.toMatchObject({ code: 'SAVE_FAILED' });
  });

  it('复杂域的含义被编辑器改变时拒绝有损整表回退', async () => {
    const sample = await open(table.replace(paragraph, '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>'));
    sample.paragraph.runs = [{ text: '普通文字', style: {} }];
    await expect(serializeDocx(sample.doc, { sourceBytes: sample.bytes, name: 'table.docx' })).rejects.toMatchObject({ code: 'SAVE_FAILED' });
  });

  it('保存失败不调用磁盘写入，保留未保存标记与原始源缓冲', async () => {
    const sample = await open();
    const platform: PlatformAdapter = {
      kind: 'web', openFile: vi.fn(), openPath: vi.fn(), saveBytes: vi.fn(),
      listFonts: vi.fn().mockResolvedValue([]), saveSnapshot: vi.fn().mockResolvedValue(undefined),
      listSnapshots: vi.fn().mockResolvedValue([]), loadSnapshot: vi.fn(), deleteSnapshot: vi.fn(), revealInFolder: vi.fn(),
    };
    registerFormatHandler(wordFormatHandler);
    setPlatformAdapter(platform);
    try {
      const id = await useWorkspaceStore.getState().openBytes({ name: 'safe-table.docx', bytes: sample.bytes });
      expect(id).toBeTruthy();
      const current = structuredClone(useWorkspaceStore.getState().sessions[id!].data) as DocModel;
      (current.blocks[0] as DocTable).rows.splice(1, 1);
      useWorkspaceStore.getState().updateData(id!, current, { dirty: true });
      expect(await useWorkspaceStore.getState().saveSession(id!)).toBe(false);
      expect(platform.saveBytes).not.toHaveBeenCalled();
      expect(useWorkspaceStore.getState().sessions[id!]).toMatchObject({ state: 'SAVE_FAILED', dirty: true });
      expect(getSourceBytes(id!)).toEqual(sample.bytes);
    } finally {
      setPlatformAdapter(null);
      useWorkspaceStore.setState({ sessions: {}, tabOrder: [], activeId: null, notifications: [] });
    }
  });
});
