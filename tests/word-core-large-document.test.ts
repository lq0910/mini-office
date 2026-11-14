/** 功能：验证长文档分批加载的完整保存、历史、只读、多会话隔离和失败保护；时间：2026-11-13；作者：lq。 */
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { WordCoreSession } from '../packages/word-core/src/session/word-session';
import { PHASE1_BLOCKS, setContentPhased, cancelPhasedContent, waitForFullContent, isPhasedContentPending, assertPhasedContentComplete, type PhasedContentHost } from '../packages/word-core/src/upstream/apps/docs/src/renderer/phased-content';
import { batchWordTailDom } from '../packages/word-core/src/layout/dom-batch';

const sessions: WordCoreSession[] = [];
beforeEach(() => {
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => setTimeout(() => callback(performance.now()), 0));
  vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id));
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
});
afterEach(() => { sessions.splice(0).forEach(session => session.dispose()); vi.restoreAllMocks(); vi.unstubAllGlobals(); });

/** 功能：生成含明确尾部标识的公共包，直接核对真实保存 XML；时间：2026-11-13；作者：lq。 */
async function sample() {
  const zip = await JSZip.loadAsync(readFileSync('tests/fixtures/docx/contract.docx'));
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${Array.from({ length: 400 }, (_, index) => `<w:p><w:r><w:t>正文${index}尾部验收</w:t></w:r></w:p>`).join('')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}
/** 功能：登记真实会话以覆盖取消时的异步清理；时间：2026-11-13；作者：lq。 */
async function open(bytes: Uint8Array, readonly = false) { const session = await WordCoreSession.open(bytes, { readonly }); sessions.push(session); return session; }

it('首屏只挂 64 块，立即保存仍等待全文且无操作原包保持逐字节一致', async () => {
  const bytes = await sample(), session = await open(bytes);
  expect(session.editor.state.doc.childCount).toBe(PHASE1_BLOCKS);
  expect(session.state()).toMatchObject({ loading: true, dirty: false, canUndo: false });
  const saved = await session.prepareSave();
  expect(session.editor.state.doc.childCount).toBe(400);
  expect(session.editor.getText()).toContain('正文399尾部验收');
  expect(Buffer.compare(saved.bytes, bytes)).toBe(0);
  expect(session.state()).toMatchObject({ loading: false, dirty: false, canUndo: false });
});

it('加载期间的首屏输入可保存和撤销，后台尾部追加永远不进入撤销历史', async () => {
  const session = await open(await sample());
  session.editor.commands.insertContentAt(2, '用户输入');
  const saved = await session.prepareSave();
  const xml = await (await JSZip.loadAsync(saved.bytes)).file('word/document.xml')!.async('text');
  expect(xml).toContain('用户输入'); expect(xml).toContain('正文399尾部验收');
  expect(session.state().revision).toBe(1);
  expect(session.editor.commands.undo()).toBe(true);
  expect(session.editor.getText()).not.toContain('用户输入');
  expect(session.editor.getText()).toContain('正文399尾部验收');
  expect(session.editor.state.doc.childCount).toBe(400);
  expect(session.editor.commands.redo()).toBe(true);
  expect(session.editor.getText()).toContain('用户输入');
});

it('只读长文档仍装载完整来源，关闭加载中会话时等待中的保存明确失败', async () => {
  const bytes = await sample(), readonly = await open(bytes, true);
  await readonly.whenContentReady();
  expect(readonly.editor.state.doc.childCount).toBe(400);
  readonly.editor.commands.insertContentAt(2, '禁止写入');
  expect(readonly.editor.getText()).not.toContain('禁止写入');
  const closing = await open(bytes);
  const save = closing.prepareSave();
  closing.dispose();
  await expect(save).rejects.toThrow('已关闭');
});

/** 功能：分帧装载不能抢占正在输入的组合文本；完成后继续同一来源尾部；时间：2026-11-14；作者：lq。 */
it('组合输入期间暂停追加，确认后完整加载并保持无编辑历史', async () => {
  const session = await open(await sample());
  const input = (session.editor.view as unknown as { input: { composing: boolean } }).input;
  input.composing = true;
  try {
    await new Promise(resolve => setTimeout(resolve, 12));
    expect(session.editor.state.doc.childCount).toBe(PHASE1_BLOCKS);
    expect(session.contentPending()).toBe(true);
  } finally { input.composing = false; }
  await session.whenContentReady();
  expect(session.editor.state.doc.childCount).toBe(400);
  expect(session.state()).toMatchObject({ dirty: false, canUndo: false });
});

/** 功能：异常和非尾部重排不能遗留被拦截的方法或改变既有 DOM 顺序；时间：2026-11-14；作者：lq。 */
it('尾部批处理遇到既有节点重排与异常时仍保持正确顺序并恢复 DOM 方法', () => {
  const root = document.createElement('div'), a = document.createElement('p'), b = document.createElement('p'), c = document.createElement('p');
  a.textContent = 'A'; b.textContent = 'B'; c.textContent = 'C'; root.append(a);
  const original = root.insertBefore;
  expect(() => batchWordTailDom(root, 10, () => {
    root.insertBefore(b, null);
    root.insertBefore(c, a);
    throw new Error('模拟视图异常');
  })).toThrow('模拟视图异常');
  expect(root.textContent).toBe('CAB');
  expect(root.insertBefore).toBe(original);
  root.insertBefore(b, a);
  expect(root.textContent).toBe('CBA');
});

it('会话独立取消，追加失败且已有用户输入时不覆盖正文也不允许保存半份文档', async () => {
  const tasks: Array<() => void> = [], a = {}, b = {};
  const doc = { type: 'doc', content: Array.from({ length: 400 }, () => ({ type: 'docParagraph' })) };
  const appended: string[] = [];
  const host = (name: string): PhasedContentHost => ({ setContent: () => {}, appendNodes: () => { appended.push(name); }, isDestroyed: () => false, resetHistory: () => {}, setLoading: () => {}, getDirty: () => false, setDirty: () => {} });
  setContentPhased(host('a'), doc, task => tasks.push(task), a);
  setContentPhased(host('b'), doc, task => tasks.push(task), b);
  cancelPhasedContent(a);
  while (tasks.length) tasks.shift()!();
  await waitForFullContent(b);
  expect(appended.length).toBeGreaterThan(0); expect(appended.every(name => name === 'b')).toBe(true);
  expect(isPhasedContentPending(a)).toBe(false); expect(isPhasedContentPending(b)).toBe(false);
  const failed = {}, remount = vi.fn();
  setContentPhased({ ...host('failed'), setContent: remount, getDirty: () => true, appendNodes: () => { throw new Error('追加失败'); } }, doc, task => tasks.push(task), failed);
  while (tasks.length) tasks.shift()!();
  await expect(waitForFullContent(failed)).rejects.toThrow('追加失败');
  expect(() => assertPhasedContentComplete(failed)).toThrow('追加失败');
  expect(remount).toHaveBeenCalledTimes(1);
});
