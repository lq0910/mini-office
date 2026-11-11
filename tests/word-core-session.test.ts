/** 功能：验证新 Word 内核独立解析、真实事务、原 XML 保存及版本确认；时间：2026-11-08；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { WordCoreSession, parseDocx, serializeWordDocument, configureWordPlatform, wordPlatformCapabilities } from '../packages/word-core/src/index';
import { measureWordLayout, WordLayoutController } from '../packages/word-core/src/layout/controller';

const active: WordCoreSession[] = [];
const continuation = '<w:tc><w:tcPr><w:vMerge/><w:tcBorders><w:bottom w:val="double" w:sz="8" w:color="FF0000"/></w:tcBorders><w:shd w:fill="DDDDDD"/><w:noWrap/></w:tcPr><w:p/></w:tc>';
/** 功能：用公共包结构构造含未知属性及纵向合并续格的最小表格；时间：2026-11-08；作者：lq。 */
async function sample() {
  const zip = await JSZip.loadAsync(readFileSync('tests/fixtures/docx/contract.docx'));
  const paragraph = '<w:p w:rsidR="KEEP"><w:bookmarkStart w:id="1" w:name="keep"/><w:r><w:rPr><w:lang w:val="zh-CN"/></w:rPr><w:t>建设内容</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>';
  const table = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:vMerge w:val="restart"/><w:tcBorders><w:top w:val="single" w:sz="4"/></w:tcBorders></w:tcPr>${paragraph}</w:tc></w:tr><w:tr>${continuation}</w:tr></w:tbl>`;
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${table}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}
/** 功能：读取保存的真实 XML，不用简化模型掩盖未知格式丢失；时间：2026-11-08；作者：lq。 */
async function documentXml(bytes: Uint8Array) { return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('text'); }
/** 功能：记录并在测试后释放实际核心会话；时间：2026-11-08；作者：lq。 */
async function open(bytes: Uint8Array, readonly = false) {
  const session = await WordCoreSession.open(bytes, { readonly });
  active.push(session);
  return session;
}
/** 功能：使用真实事务替换目标文字，保持表格来源映射；时间：2026-11-08；作者：lq。 */
function replace(session: WordCoreSession, text: string) {
  let position = -1;
  let length = 0;
  session.editor.state.doc.descendants((node, pos) => {
    if (position < 0 && node.isText) { position = pos; length = node.text!.length; }
  });
  if (position < 0) throw new Error('未定位到目标文字');
  session.editor.commands.insertContentAt({ from: position, to: position + length }, text);
}

beforeEach(() => {
  // 功能：jsdom 没有真实帧时钟，仅补浏览器调度，不模拟保存或编辑命令；时间：2026-11-08；作者：lq。
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => setTimeout(() => callback(performance.now()), 0));
  vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id));
  // 功能：此组仅验证事务与 XML；显式禁用 jsdom 缺失的 Canvas，保留上游无 Canvas 的正常回退；时间：2026-11-08；作者：lq。
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
  configureWordPlatform({});
});
afterEach(() => { for (const session of active.splice(0)) session.dispose(); vi.restoreAllMocks(); vi.unstubAllGlobals(); vi.useRealTimers(); });

describe('定制 Word 核心的会话和来源保留', () => {
  // 功能：滚动/选区复用统计，真实编辑及撤销刷新结果，避免以过期缓存换取速度；时间：2026-11-11；作者：lq。
  it('统计按不可变正文缓存，文字编辑与撤销后结果保持正确', async () => {
    const session = await open(await sample());
    const original = session.statistics();
    session.editor.commands.setTextSelection(3);
    expect(session.statistics()).toBe(original);
    replace(session, '测试');
    expect(session.statistics()).not.toBe(original);
    expect(session.statistics().characters).toBe(2);
    session.editor.commands.undo();
    expect(session.statistics()).toEqual(original);
  });
  it('实际 schema 挂载后无操作保存返回完全相同的原包，不引入旧 DocModel', async () => {
    const bytes = await sample();
    const session = await open(bytes);
    expect(session.editor.schema.nodes.docTable).toBeTruthy();
    // 功能：逐字节比较来源包，避免测试框架为大型二进制差异构建巨型结构；时间：2026-11-11；作者：lq。
    expect(session.parsed.internal.originalBytes.length === bytes.length && session.parsed.internal.originalBytes.every((value, index) => value === bytes[index])).toBe(true);
    const saved = (await session.prepareSave()).bytes;
    expect(saved.length === bytes.length && saved.every((value, index) => value === bytes[index])).toBe(true);
    expect(session.state().dirty).toBe(false);
  });

  it('真实表格文字事务保留物理续格、未知属性及非目标 ZIP 部件', async () => {
    const bytes = await sample();
    const session = await open(bytes);
    replace(session, '智慧物业信息系统部署实施');
    expect(session.state().dirty).toBe(true);
    const saved = await session.prepareSave();
    const xml = await documentXml(saved.bytes);
    expect(xml).toContain('智慧物业信息系统部署实施');
    expect(xml).toContain(continuation);
    for (const fragment of ['w:rsidR="KEEP"', '<w:bookmarkStart w:id="1" w:name="keep"/>', '<w:cantSplit/>', '<w:tblLayout w:type="fixed"/>']) expect(xml).toContain(fragment);
    const before = await JSZip.loadAsync(bytes);
    const after = await JSZip.loadAsync(saved.bytes);
    for (const [name, entry] of Object.entries(before.files)) {
      if (entry.dir || name === 'word/document.xml') continue;
      // 功能：上游按预期更新修改时间和修订号；仅放行这两个元数据字段，其余部件逐字比较；时间：2026-11-08；作者：lq。
      if (name === 'docProps/core.xml') {
        const normalize = (xml: string) => xml.replace(/<dcterms:modified\b[^>]*>[\s\S]*?<\/dcterms:modified>/g, '<modified/>').replace(/<cp:revision>[\s\S]*?<\/cp:revision>/g, '<revision/>');
        expect(normalize(await after.file(name)!.async('text'))).toBe(normalize(await entry.async('text')));
      } else expect(await after.file(name)!.async('uint8array'), name).toEqual(await entry.async('uint8array'));
    }
    expect((await parseDocx(saved.bytes)).blocks[0].table!.rows.length).toBe(2);
  });

  it('生成快照不宣布保存成功，期间继续编辑仍保留 dirty，连续保存和重开不漂移', async () => {
    const bytes = await sample();
    const session = await open(bytes);
    replace(session, '第一次编辑');
    const first = await session.prepareSave();
    expect(session.state().dirty).toBe(true);
    replace(session, '第二次编辑');
    session.acknowledgeSave(first);
    expect(session.state().dirty).toBe(true);
    const second = await session.prepareSave();
    session.acknowledgeSave(second);
    expect(session.state().dirty).toBe(false);
    // 功能：逐字节比较来源包，避免测试框架为大型二进制差异构建巨型结构；时间：2026-11-11；作者：lq。
    expect(session.parsed.internal.originalBytes.length === bytes.length && session.parsed.internal.originalBytes.every((value, index) => value === bytes[index])).toBe(true);
    const reopened = await open(second.bytes);
    replace(reopened, '重开后编辑');
    const third = await reopened.prepareSave();
    expect(await documentXml(third.bytes)).toContain(continuation);
    expect(await documentXml(third.bytes)).toContain('重开后编辑');
  });

  it('新核心使用自己的撤销重做，并拒绝跨会话及重复保存确认', async () => {
    const bytes = await sample();
    const session = await open(bytes);
    replace(session, '可撤销编辑');
    expect(session.editor.commands.undo()).toBe(true);
    expect(session.editor.getText()).toContain('建设内容');
    expect(session.editor.commands.redo()).toBe(true);
    expect(session.editor.getText()).toContain('可撤销编辑');
    const saved = await session.prepareSave();
    const other = await open(bytes);
    expect(() => other.acknowledgeSave(saved)).toThrow('不属于');
    session.acknowledgeSave(saved);
    expect(() => session.acknowledgeSave(saved)).toThrow('已经确认');
  });

  it('只读设置不能修改非正文属性，缺少的原生能力明确报告不可用', async () => {
    const session = await open(await sample(), true);
    const initial = session.editor.getText();
    replace(session, '不能绕过只读');
    expect(session.editor.getText()).toBe(initial);
    expect(() => session.updateSaveOptions({ titlePg: true })).toThrow('只读');
    await expect(session.prepareSave()).rejects.toThrow('只读');
    expect(wordPlatformCapabilities()).toEqual({ fontMetrics: false, copyImageToClipboard: false, parseWorker: false });
  });

  it('输入法组合中不生成候选快照，关闭后不允许再保存', async () => {
    const session = await open(await sample());
    const input = (session.editor.view as unknown as { input: { composing: boolean } }).input;
    input.composing = true;
    await expect(session.prepareSave()).rejects.toThrow('输入尚未确认');
    input.composing = false;
    session.dispose();
    await expect(session.prepareSave()).rejects.toThrow('已关闭');
  });
});

describe('Word 排版协调边界', () => {
  // 功能：排版失败恢复自然流和列测量状态，不能遗留隐藏装饰的 CSS 类；时间：2026-11-09；作者：lq。
  it('异常后恢复测量状态，并保留调用方原有的测量类', () => {
    const root = document.createElement('div');
    const input = { root, origin: 0, zoomFactor: 1, sections: [{ contentHeight: -1, forceBreak: false, cols: 2 }] };
    expect(() => measureWordLayout(input)).toThrow('无效');
    expect(root.classList.contains('measuring-natural')).toBe(false);
    expect(root.classList.contains('measuring-columns')).toBe(false);
    root.classList.add('measuring-natural', 'measuring-columns');
    expect(() => measureWordLayout(input)).toThrow('无效');
    expect(root.classList.contains('measuring-natural')).toBe(true);
    expect(root.classList.contains('measuring-columns')).toBe(true);
  });
  it('缺少页面几何时明确失败，组合输入期间不测量，释放后不会迟到重排', () => {
    const root = document.createElement('div');
    expect(() => measureWordLayout({ root, origin: 0, zoomFactor: 1, sections: [] })).toThrow('几何');
    vi.useFakeTimers();
    const apply = vi.fn();
    const error = vi.fn();
    const controller = new WordLayoutController(() => ({ root, origin: 0, zoomFactor: 1, sections: [], composing: true }), apply, error);
    controller.request();
    vi.runAllTimers();
    expect(apply).not.toHaveBeenCalled();
    expect(error).not.toHaveBeenCalled();
    controller.request();
    controller.dispose();
    vi.runAllTimers();
    expect(apply).not.toHaveBeenCalled();
  });
  /** 功能：编辑突发合并到最后一次输入，后续补偿保持 100ms；手动刷新及销毁都清理计时器；时间：2026-11-11；作者：lq。 */
  it('首次即时排版，后续编辑防抖且显式 flush 不等待', () => {
    vi.useFakeTimers();
    const root = document.createElement('div'), apply = vi.fn(), error = vi.fn();
    const controller = new WordLayoutController(() => ({ root, origin: 0, zoomFactor: 1, sections: [{ contentHeight: 100, forceBreak: false }] }), apply, error);
    try {
      controller.request(); vi.advanceTimersByTime(0); expect(apply).toHaveBeenCalledTimes(1);
      controller.request(); vi.advanceTimersByTime(299); expect(apply).toHaveBeenCalledTimes(1);
      controller.request(); vi.advanceTimersByTime(299); expect(apply).toHaveBeenCalledTimes(1);
      vi.advanceTimersByTime(1); expect(apply).toHaveBeenCalledTimes(2);
      controller.followUp('width'); vi.advanceTimersByTime(99); expect(apply).toHaveBeenCalledTimes(2);
      vi.advanceTimersByTime(1); expect(apply).toHaveBeenCalledTimes(3);
      controller.request(); controller.flush(); expect(apply).toHaveBeenCalledTimes(4);
      vi.runAllTimers(); expect(apply).toHaveBeenCalledTimes(4);
      controller.request(); controller.dispose(); vi.runAllTimers(); expect(apply).toHaveBeenCalledTimes(4);
      expect(error).not.toHaveBeenCalled();
    } finally { controller.dispose(); vi.useRealTimers(); }
  });
});
