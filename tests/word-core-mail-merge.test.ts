/** 功能：验证原生合并字段、来源 XML 保留、NEXT/条件、附件关系及输出重开；时间：2026-11-11；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { buildBlankDocx, WordCoreSession, generateTocFieldXml } from '../packages/word-core/src/index';
import { insertWordMergeField, wordMergeFields } from '../packages/word-core/src/editing/mail-merge';
import { inspectWordMergePackage, mergeWordPackage } from '../packages/word-core/src/serialization/mail-merge';
import { mergeWordXml } from '../packages/word-core/src/serialization/merge-fields';
import { patchWordMergeSettings } from '../packages/word-core/src/serialization/mail-merge-settings';
import { WordHeaderFooterEditor } from '../packages/word-core/src/editing/header-footer';
import { xmlSpans, wordXmlNode, wordXmlAttr } from '../packages/word-core/src/serialization/xml-spans';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const cores: WordCoreSession[] = [];
const records = [{ 姓名: '张三 & <一>', 金额: '123', 状态: '已缴' }, { 姓名: '李四', 金额: '0', 状态: '未缴' }, { 姓名: '王五', 金额: '5', 状态: '已缴' }];
/** 功能：真实标准简单域与复杂域夹具，允许域指令拆分到多个 run；时间：2026-11-11；作者：lq。 */
const simple = (instr: string, text = '缓存') => `<w:fldSimple w:instr='${instr}'><w:r><w:rPr><w:b/><w:lang w:val="zh-CN"/><w:extProp w:val="keep"/></w:rPr><w:t>${text}</w:t></w:r></w:fldSimple>`;
const complex = (instr: string, text = '缓存') => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>${text}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
/** 功能：只构造独立公开 DOCX 测试包，未知部件以二进制断言保留；时间：2026-11-11；作者：lq。 */
async function template(body: string, extras?: (zip: JSZip) => void) {
  const zip = await JSZip.loadAsync(await buildBlankDocx());
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  zip.file('customXml/native-keep.bin', new Uint8Array([0, 21, 255])); extras?.(zip);
  return zip.generateAsync({ type: 'uint8array' });
}
async function open(bytes: Uint8Array) { const core = await WordCoreSession.open(bytes); cores.push(core); return core; }
beforeEach(() => { vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null); vi.stubGlobal('requestAnimationFrame', (fn: FrameRequestCallback) => setTimeout(() => fn(performance.now()), 0)); vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id)); });
afterEach(() => { cores.splice(0).forEach(core => core.dispose()); vi.restoreAllMocks(); vi.unstubAllGlobals(); });
describe('M3 原生邮件合并', () => {
  it('页眉新增字段与主摘要同步，撤销删除不显示已移除字段', async () => {
    const core = await open(await buildBlankDocx()), host = document.createElement('div'); document.body.append(host);
    const region = new WordHeaderFooterEditor(core, host, { section: 0, kind: 'header', variant: 'default' });
    try {
      expect(insertWordMergeField(region.editor, '区域姓名')).toBe(true);
      expect(wordMergeFields(core.editor.state.doc, core.parsed).fields).toContain('区域姓名');
      region.editor.commands.undo(); expect(wordMergeFields(core.editor.state.doc, core.parsed).fields).not.toContain('区域姓名');
      region.editor.commands.redo(); expect((await inspectWordMergePackage((await core.prepareSave()).bytes)).fields).toContain('区域姓名');
    } finally { region.dispose(); host.remove(); }
  });
  it('跨段目录与合并字段共存时摘要不抛错，损坏来源不阻断编辑且实际合并仍校验', async () => {
    const toc = generateTocFieldXml([{ text: '一章', level: 1 }, { text: '二章', level: 2 }]).join('');
    const core = await open(await template(toc + `<w:p>${simple('MERGEFIELD 姓名')}</w:p>`));
    expect(wordMergeFields(core.editor.state.doc, core.parsed)).toEqual({ fields: ['姓名'], nextBlocks: 0 });
    const malformed = await template('<w:p>' + complex('MERGEFIELD 姓名').replace('<w:r><w:fldChar w:fldCharType="end"/></w:r>', '') + '</w:p>');
    const broken = await open(malformed);
    expect(() => wordMergeFields(broken.editor.state.doc, broken.parsed)).not.toThrow();
    await expect(mergeWordPackage(malformed, records)).rejects.toThrow('结束');
    expect(broken.state().dirty).toBe(false);
  });
  it('字段插入使用原生域、当前格式和独立历史，保存重开后仍可合并且不修改模板版本', async () => {
    const core = await open(await buildBlankDocx()); core.editor.commands.insertContent('通知：'); const before = core.editor.getJSON();
    expect(insertWordMergeField(core.editor, '姓名')).toBe(true); const inserted = core.editor.getJSON();
    core.editor.commands.insertContent(' 您好'); core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(inserted);
    core.editor.commands.undo(); expect(core.editor.getJSON()).toEqual(before); core.editor.commands.redo();
    const bytes = (await core.prepareSave()).bytes; expect((await inspectWordMergePackage(bytes)).fields).toEqual(['姓名']);
    const revision = core.state().revision; const merged = await mergeWordPackage(bytes, records.slice(0, 2));
    expect(core.state().revision).toBe(revision); expect(core.editor.state.doc.textContent).toBe('通知：«姓名»');
    const result = await open(merged); expect(result.editor.state.doc.textContent).toContain('张三 & <一>'); expect(result.editor.state.doc.textContent).toContain('李四');
    expect(wordMergeFields(result.editor.state.doc).fields).toEqual([]);
    const saved = await result.prepareSave(); expect((await open(saved.bytes)).editor.state.doc.textContent).toBe(result.editor.state.doc.textContent);
  });
  it('复杂域/简单域/嵌套 IF 与开关正确替换，保留未知段落和 run XML 及动态 PAGE', async () => {
    const nested = complex('IF «状态» = "已缴" "谢谢" "催缴"');
    const xml = `<w:document xmlns:w="${W}"><w:body><w:p custom="keep"><w:pPr><w:contextualSpacing/></w:pPr><w:bookmarkStart w:id="1" w:name="保留"/>${simple('MERGEFIELD "姓名" \\b "尊敬的 " \\f "!"')}${complex('MERGEFIELD 金额')}${nested}${complex('PAGE', '3')}<w:bookmarkEnd w:id="1"/></w:p></w:body></w:document>`;
    const merged = mergeWordXml(xml, records, { index: 0, skipped: false });
    expect(merged).toContain('尊敬的 张三 &amp; &lt;一&gt;!'); expect(merged).toContain('123'); expect(merged).toContain('谢谢'); expect(merged).toContain('<w:p custom="keep">');
    expect(merged).toContain('<w:extProp w:val="keep"/>'); expect(merged).toContain('<w:bookmarkStart w:id="1" w:name="保留"/>'); expect(merged).toContain(complex('PAGE', '3')); expect(merged).not.toContain('MERGEFIELD');
    // 功能：内层真实 MERGEFIELD 指令进入外层 IF，不把旧结果缓存当条件；时间：2026-11-11；作者：lq。
    const realNested = `<w:p xmlns:w="${W}"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>IF </w:instrText></w:r>${complex('MERGEFIELD 状态', '过期缓存')}<w:r><w:instrText> = "已缴" "谢谢" "催缴"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>旧结果</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const actual = mergeWordXml(realNested, records, { index: 1, skipped: false }); expect(actual).toContain('催缴'); expect(actual).not.toContain('旧结果'); expect(actual).not.toContain('过期缓存');
  });
  it('NEXT 在表格/正文共享游标，末页空槽不串记录，控制专用段落不增加间距', async () => {
    const bytes = await template(`<w:p>${simple('MERGEFIELD 姓名')}</w:p><w:p>${complex('NEXT', '')}</w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p>${simple('MERGEFIELD 姓名')}</w:p></w:tc></w:tr></w:tbl>`);
    expect((await inspectWordMergePackage(bytes)).nextBlocks).toBe(2);
    const zip = await JSZip.loadAsync(await mergeWordPackage(bytes, records)); const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml.match(/<w:tbl>/g)).toHaveLength(2); expect(xml).not.toContain('MERGEFIELD'); expect(xml).not.toContain('NEXT');
    for (const name of ['张三 &amp; &lt;一&gt;', '李四', '王五']) expect(xml.split(name)).toHaveLength(2);
    expect(await zip.file('customXml/native-keep.bin')!.async('uint8array')).toEqual(new Uint8Array([0, 21, 255]));
    const result = await open(await zip.generateAsync({ type: 'uint8array' })); expect(result.editor.state.doc.textContent).toContain('王五');
  });
  it('重复记录的书签引用、批注/脚注及显式编号使用各自 ID，页眉页脚按记录生成且媒体关系保留', async () => {
    const bytes = await template(`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:bookmarkStart w:id="1" w:name="same"/>${simple('MERGEFIELD 姓名')}<w:hyperlink w:anchor="same"><w:r><w:t>链接</w:t></w:r></w:hyperlink><w:r><w:footnoteReference w:id="1"/><w:commentReference w:id="0"/></w:r><w:bookmarkEnd w:id="1"/></w:p><w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdMergeHeader"/></w:sectPr></w:pPr><w:r><w:t>第二节</w:t></w:r></w:p>`, zip => {
      zip.file('word/header99.xml', `<w:hdr xmlns:w="${W}" xmlns:r="${R}"><w:p>${simple('MERGEFIELD 姓名')}${complex('PAGE', '1')}</w:p></w:hdr>`);
      zip.file('word/_rels/header99.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rImg" Type="image" Target="media/keep.png"/></Relationships>');
      zip.file('word/media/keep.png', new Uint8Array([5, 6, 7]));
      const rel = zip.file('word/_rels/document.xml.rels')!;
      void rel;
      zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdMergeHeader" Type="${R}/header" Target="header99.xml"/><Relationship Id="fn" Type="${R}/footnotes" Target="footnotes.xml"/><Relationship Id="cm" Type="${R}/comments" Target="comments.xml"/><Relationship Id="nu" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`);
      zip.file('word/footnotes.xml', `<w:footnotes xmlns:w="${W}"><w:footnote w:id="-1" w:type="separator"><w:p/></w:footnote><w:footnote w:id="1"><w:p>${simple('MERGEFIELD 姓名')}</w:p></w:footnote></w:footnotes>`);
      zip.file('word/comments.xml', `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="原作者"><w:p>${simple('MERGEFIELD 姓名')}</w:p></w:comment></w:comments>`);
      zip.file('word/numbering.xml', `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
    });
    const zip = await JSZip.loadAsync(await mergeWordPackage(bytes, records.slice(0, 2)));
    const xml = await zip.file('word/document.xml')!.async('string'); const bookmarks = [...xml.matchAll(/<w:bookmarkStart[^>]+w:id="(\d+)"[^>]+w:name="([^"]+)"/g)];
    expect(new Set(bookmarks.map(match => match[1])).size).toBe(2); expect(new Set(bookmarks.map(match => match[2])).size).toBe(2);
    for (const match of bookmarks) expect(xml).toContain(`w:anchor="${match[2]}"`);
    const footnotes = await zip.file('word/footnotes.xml')!.async('string'), comments = await zip.file('word/comments.xml')!.async('string');
    const references = [...xml.matchAll(/<w:footnoteReference w:id="(\d+)"/g)].map(match => match[1]); expect(new Set(references).size).toBe(2);
    for (const id of references) expect(footnotes).toContain(`<w:footnote w:id="${id}">`);
    expect(comments.match(/<w:comment /g)).toHaveLength(2); expect(comments).not.toContain('MERGEFIELD');
    const headers = Object.keys(zip.files).filter(path => /^word\/mo-merge-.*header\.xml$/.test(path)); expect(headers).toHaveLength(2);
    expect(await zip.file(headers[0])!.async('string')).toContain('张三 &amp; &lt;一&gt;'); expect(await zip.file(headers[1])!.async('string')).toContain('李四');
    expect(await zip.file(headers[0].replace(/([^/]+)$/, '_rels/$1.rels'))!.async('string')).toContain('media/keep.png');
    const nums = [...xml.matchAll(/<w:numId w:val="(\d+)"/g)].map(match => match[1]); expect(new Set(nums).size).toBe(2);
    const numbering = await zip.file('word/numbering.xml')!.async('string'); for (const id of nums) expect(numbering).toContain(`<w:num w:numId="${id}">`);
  });
  it('SKIPIF 使用实际记录并清除成品邮件关系，取消及损坏域不污染原字节', async () => {
    const bytes = await template(`<w:p>${simple('MERGEFIELD 姓名')}${simple('SKIPIF «状态» = "未缴"')}</w:p>`, zip => {
      zip.file('word/settings.xml', `<w:settings xmlns:w="${W}" xmlns:r="${R}"><w:mailMerge><w:dataSource r:id="source"/></w:mailMerge><w:docVars><w:docVar w:name="mailMergePanelState" w:val="{&quot;sort&quot;:&quot;姓名&quot;}"/><w:docVar w:name="keep" w:val="yes"/></w:docVars><w:compat/></w:settings>`);
      zip.file('word/_rels/settings.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="source" Type="${R}/mailMergeSource" Target="file:///data.xlsx" TargetMode="External"/><Relationship Id="other" Type="keep" Target="other.xml"/></Relationships>`);
    });
    expect((await inspectWordMergePackage(bytes)).panelState).toBe('{"sort":"姓名"}');
    const original = bytes.slice(); const zip = await JSZip.loadAsync(await mergeWordPackage(bytes, records));
    const xml = await zip.file('word/document.xml')!.async('string'); expect(xml).not.toContain('李四'); expect(xml).toContain('王五');
    const settings = await zip.file('word/settings.xml')!.async('string'); expect(settings).not.toContain('mailMerge'); expect(settings).toContain('<w:docVar w:name="keep" w:val="yes"/>');
    expect(await zip.file('word/_rels/settings.xml.rels')!.async('string')).not.toContain('mailMergeSource');
    let cancel = false; await expect(mergeWordPackage(bytes, records, { isCancelled: () => cancel, onProgress: () => { cancel = true; } })).rejects.toThrow('已取消'); expect(bytes).toEqual(original);
    const broken = await template('<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>MERGEFIELD 姓名</w:instrText></w:r></w:p>');
    await expect(mergeWordPackage(broken, records)).rejects.toThrow('结束标记');
  });
  // 功能：验证 NEXT 两侧的真实脚注和现代线程批注使用各自记录，父回复和持久标识成对重写；时间：2026-11-11；作者：lq。
  it('NEXT 后的注释按引用处记录生成，现代批注线程与父回复映射保持完整', async () => {
    const bytes = await template(`<w:p>${simple('MERGEFIELD 姓名')}<w:commentRangeStart w:id="0"/><w:r><w:footnoteReference w:id="1"/><w:commentReference w:id="0"/></w:r><w:commentRangeEnd w:id="0"/>${complex('NEXT', '')}${simple('MERGEFIELD 姓名')}<w:r><w:footnoteReference w:id="2"/><w:commentReference w:id="2"/></w:r></w:p>`, zip => {
      zip.file('word/footnotes.xml', `<w:footnotes xmlns:w="${W}"><w:footnote w:id="0" w:type="continuationSeparator"><w:p/></w:footnote>${[1,2].map(id => `<w:footnote w:id="${id}"><w:p>${simple('MERGEFIELD 姓名')}</w:p></w:footnote>`).join('')}</w:footnotes>`);
      zip.file('word/comments.xml', `<w:comments xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${[0,1,2].map(id => `<w:comment w:id="${id}" w:author="作者"><w:p w14:paraId="0000000${id + 1}">${simple('MERGEFIELD 姓名')}</w:p></w:comment>`).join('')}</w:comments>`);
      zip.file('word/commentsExtended.xml', '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="00000001"/><w15:commentEx w15:paraId="00000002" w15:paraIdParent="00000001"/><w15:commentEx w15:paraId="00000003"/></w15:commentsEx>');
      zip.file('word/commentsIds.xml', '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w16cid:commentId w16cid:paraId="00000001" w16cid:durableId="00000009"/></w16cid:commentsIds>');
      zip.file('word/commentsExtensible.xml', '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"><w16cex:commentExtensible w16cex:durableId="00000009" w16cex:dateUtc="2026-11-11T00:00:00Z"/></w16cex:commentsExtensible>');
    });
    const zip = await JSZip.loadAsync(await mergeWordPackage(bytes, records.slice(0,2)));
    const xml = await zip.file('word/document.xml')!.async('string');
    const ids = xmlSpans(xml).filter(node => wordXmlNode(node, 'footnoteReference')).map(node => wordXmlAttr(node, 'id'));
    const notes = await zip.file('word/footnotes.xml')!.async('string');
    const text = (id: string | undefined) => { const node = xmlSpans(notes).find(node => wordXmlNode(node, 'footnote') && wordXmlAttr(node, 'id') === id)!; return notes.slice(node.start, node.end); };
    expect(text(ids[0])).toContain('张三 &amp; &lt;一&gt;'); expect(text(ids[1])).toContain('李四');
    expect(notes).toContain('w:type="continuationSeparator"');
    const comments = await zip.file('word/comments.xml')!.async('string'), extended = await zip.file('word/commentsExtended.xml')!.async('string');
    const paragraphs = xmlSpans(comments).filter(node => wordXmlNode(node, 'p')).map(node => node.attrs['w14:paraId']);
    expect(new Set(paragraphs).size).toBe(paragraphs.length);
    const metadata = xmlSpans(extended).filter(node => node.local === 'commentEx');
    expect(metadata).toHaveLength(6);
    for (const node of metadata) { expect(paragraphs).toContain(node.attrs['w15:paraId']); if (node.attrs['w15:paraIdParent']) expect(paragraphs).toContain(node.attrs['w15:paraIdParent']); }
    const cids = xmlSpans(await zip.file('word/commentsIds.xml')!.async('string')).filter(node => node.local === 'commentId');
    const extensible = await zip.file('word/commentsExtensible.xml')!.async('string');
    for (const node of cids) { expect(paragraphs).toContain(node.attrs['w16cid:paraId']); expect(extensible).toContain(`durableId="${node.attrs['w16cid:durableId']}"`); }
  });
  // 功能：验证模板首节空白页眉、附件位置及样式继承编号不能跨记录串用；时间：2026-11-11；作者：lq。
  it('原包 ID 和生成路径避让，首节空白重置继承，子目录附件重定位，样式编号每份独立', async () => {
    const bytes = await template(`<w:p><w:pPr><w:pStyle w:val="Derived"/></w:pPr>${simple('MERGEFIELD 姓名')}<w:bookmarkStart w:id="1000001" w:name="old"/><w:bookmarkEnd w:id="1000001"/></w:p><w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr></w:p>`, zip => {
      zip.file('word/styles.xml', `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="List"><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1000001"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Derived"><w:basedOn w:val="List"/></w:style></w:styles>`);
      zip.file('word/numbering.xml', `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%3."/></w:lvl></w:abstractNum><w:num w:numId="1000001"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
      zip.file('word/nested/header.xml', `<w:hdr xmlns:w="${W}" xmlns:r="${R}"><w:p>${simple('MERGEFIELD 姓名')}</w:p></w:hdr>`);
      zip.file('word/nested/_rels/header.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="image" Type="${R}/image" Target="../media/keep.png"/></Relationships>`);
      zip.file('word/media/keep.png', new Uint8Array([1, 2]));
      zip.file('word/mo-merge-1-header.xml', '<keep/>');
      zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="lastHeader" Type="${R}/header" Target="nested/header.xml"/></Relationships>`);
    });
    const source = await JSZip.loadAsync(bytes); let document = await source.file('word/document.xml')!.async('string');
    const at = document.lastIndexOf('<w:sectPr>') + '<w:sectPr>'.length; document = document.slice(0, at) + '<w:headerReference w:type="default" r:id="lastHeader"/>' + document.slice(at);
    source.file('word/document.xml', document);
    const zip = await JSZip.loadAsync(await mergeWordPackage(await source.generateAsync({ type: 'uint8array' }), records.slice(0,2)));
    expect(await zip.file('word/mo-merge-1-header.xml')!.async('string')).toBe('<keep/>');
    const xml = await zip.file('word/document.xml')!.async('string'), sections = xmlSpans(xml).filter(node => wordXmlNode(node, 'sectPr'));
    expect(sections).toHaveLength(4);
    const refs = sections.map(section => section.children.find(node => wordXmlNode(node, 'headerReference'))?.attrs['r:id']);
    expect(refs[0]).toBeTruthy(); expect(refs[0]).toBe(refs[2]); expect(refs[1]).not.toBe(refs[3]);
    const nums = xmlSpans(xml).filter(node => wordXmlNode(node, 'numId')).map(node => wordXmlAttr(node, 'val'));
    expect(nums).toHaveLength(2); expect(new Set(nums).size).toBe(2); expect(nums).not.toContain('1000001'); expect(xml).toContain('w:ilvl w:val="2"');
    const headers = Object.keys(zip.files).filter(path => /^word\/mo-merge-\d+-header.xml$/.test(path) && path !== 'word/mo-merge-1-header.xml');
    for (const header of headers) if ((await zip.file(header)!.async('string')).includes('张三') || (await zip.file(header)!.async('string')).includes('李四')) expect(await zip.file(header.replace(/([^/]+)$/, '_rels/$1.rels'))!.async('string')).toContain('Target="media/keep.png"');
  });
  // 功能：保存面板配置时来源关系有保留/替换/清除三态；时间：2026-11-11；作者：lq。
  it('面板配置保留其他文档变量，缺省 settings 可补齐，CSV 清除旧连接，快照隔离并发保存', async () => {
    const bytes = await template('<w:p/>', zip => {
      zip.file('word/settings.xml', `<w:settings xmlns:w="${W}" xmlns:r="${R}"><w:mailMerge><w:dataSource r:id="old"/><w:recipientData r:id="recipient"/></w:mailMerge><w:docVars><w:docVar w:name="keep" w:val="yes"/></w:docVars><w:compat/></w:settings>`);
      zip.file('word/_rels/settings.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="old" Type="${R}/mailMergeSource" Target="old.xlsx" TargetMode="External"/><Relationship Id="other" Type="keep" Target="other.xml"/></Relationships>`);
    });
    const preserved = await patchWordMergeSettings(bytes, { panelState: '{"sort":"姓名 & <一>"}' });
    expect((await inspectWordMergePackage(preserved)).panelState).toBe('{"sort":"姓名 & <一>"}');
    let zip = await JSZip.loadAsync(preserved); expect(await zip.file('word/settings.xml')!.async('string')).toContain('r:id="old"');
    zip = await JSZip.loadAsync(await patchWordMergeSettings(preserved, { panelState: '{}', source: null }));
    const xml = await zip.file('word/settings.xml')!.async('string'); expect(xml).not.toContain('<w:mailMerge'); expect(xml).toContain('w:name="keep"'); expect(xml).toContain('<w:compat/>');
    expect(await zip.file('word/_rels/settings.xml.rels')!.async('string')).not.toContain('mailMergeSource');
    const blank = await JSZip.loadAsync(await buildBlankDocx()); blank.remove('word/settings.xml');
    zip = await JSZip.loadAsync(await patchWordMergeSettings(await blank.generateAsync({ type: 'uint8array' }), { panelState: '{"headerRow":3}' }));
    expect(await zip.file('word/settings.xml')!.async('string')).toContain('mailMergePanelState');
    expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain('/settings');
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('/word/settings.xml');
    const core = await open(await buildBlankDocx()); core.editor.commands.insertContent('第一版'); core.updateMergeSettings({ panelState: 'first' });
    const first = core.prepareSave(); core.editor.commands.insertContent('第二版'); core.updateMergeSettings({ panelState: 'second' }); const second = core.prepareSave();
    const [a,b] = await Promise.all([first,second]); expect(a.revision).toBeLessThan(b.revision);
    expect((await inspectWordMergePackage(a.bytes)).panelState).toBe('first'); expect((await inspectWordMergePackage(b.bytes)).panelState).toBe('second');
    expect((await open(a.bytes)).editor.state.doc.textContent).toBe('第一版'); expect((await open(b.bytes)).editor.state.doc.textContent).toBe('第一版第二版');
    expect(core.state().dirty).toBe(true);
  });
  // 功能：带旧图片/软换行的域缓存不能在替换后继续输出旧内容；时间：2026-11-11；作者：lq。
  it('字段多行/制表值替换旧缓存，空简单域也能输出实际值', () => {
    const xml = `<w:p xmlns:w="${W}"><w:fldSimple w:instr='MERGEFIELD 姓名'><w:r><w:t>旧</w:t><w:br/><w:t>缓存</w:t><w:tab/></w:r></w:fldSimple><w:fldSimple w:instr='MERGEFIELD 金额'/></w:p>`;
    const merged = mergeWordXml(xml, [{ 姓名: '新\n行\t值', 金额: '9' }], { index: 0, skipped: false });
    expect(merged).not.toContain('旧'); expect(merged).not.toContain('缓存'); expect(merged.match(/<w:br\/>/g)).toHaveLength(1); expect(merged.match(/<w:tab\/>/g)).toHaveLength(1); expect(merged).toContain('>9</w:t>'); expect(() => xmlSpans(merged)).not.toThrow();
  });
});
