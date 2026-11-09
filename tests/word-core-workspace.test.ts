/** 功能：用真实 Word 内核验证产品工作区的打开、新建、写盘版本、快照和标签历史；时间：2026-11-09；作者：lq。 */
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { registerFormatHandler, useWorkspaceStore } from '@mini-office/document-model';
import { setPlatformAdapter, type PlatformAdapter } from '@mini-office/platform-adapter';
import { ErrorCodes, OfficeError } from '@mini-office/shared';
import { coreWordFormatHandler } from '../packages/editor-word/src/core-handler';
import { WordCoreSurface } from '../packages/editor-word/src/core-surface';

/** 功能：只替换外部文件接口，解析、事务、序列化和保存确认均使用真实核心；时间：2026-11-09；作者：lq。 */
function platform() {
  return { kind: 'tauri', openPath: vi.fn(), openFile: vi.fn(),
    saveBytes: vi.fn().mockResolvedValue({ displayName: '保存.docx', path: '/tmp/core.docx', location: 'in-place' }),
    saveSnapshot: vi.fn().mockResolvedValue(undefined), listSnapshots: vi.fn().mockResolvedValue([]), loadSnapshot: vi.fn(),
    deleteSnapshot: vi.fn().mockResolvedValue(undefined), listFonts: vi.fn().mockResolvedValue([]), revealInFolder: vi.fn(),
  } satisfies PlatformAdapter;
}
let adapter: ReturnType<typeof platform>;
beforeEach(async () => {
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => setTimeout(() => callback(performance.now()), 0));
  vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id));
  vi.stubGlobal('CSS', { escape: (value: string) => Array.from(value, char => `\\${char.codePointAt(0)!.toString(16)} `).join('') });
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
  await useWorkspaceStore.getState().closeAll(true);
  useWorkspaceStore.setState({ sessions: {}, tabOrder: [], activeId: null, notifications: [], fonts: [], settings: { ...useWorkspaceStore.getState().settings, autoSave: false } });
  adapter = platform(); setPlatformAdapter(adapter); registerFormatHandler(coreWordFormatHandler);
});
afterEach(async () => { await useWorkspaceStore.getState().closeAll(true); setPlatformAdapter(null); document.body.replaceChildren(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });
/** 功能：打开公共样本后等待首次快照完成，避免把外部 I/O 计时当解析行为；时间：2026-11-09；作者：lq。 */
async function open() {
  const id = await useWorkspaceStore.getState().openBytes({ name: '合同.docx', bytes: readFileSync('tests/fixtures/docx/contract.docx'), path: '/tmp/core.docx' });
  expect(id).toBeTruthy();
  await vi.waitFor(() => expect(adapter.saveSnapshot).toHaveBeenCalled());
  return id!;
}

it('新建和打开持有真实完整核心，工作区保存不会调用旧模型序列化', async () => {
  const id = await useWorkspaceStore.getState().createDocument('docx');
  const session = useWorkspaceStore.getState().sessions[id!];
  expect(session.data).toEqual({ engine: 'word-core' }); expect(session.wordCore).toBeTruthy();
  await vi.waitFor(() => expect(adapter.saveSnapshot).toHaveBeenCalled());
  const serialize = vi.spyOn(coreWordFormatHandler, 'serialize');
  session.wordCore!.editor.commands.insertContent('同一新核心保存正文');
  expect(useWorkspaceStore.getState().sessions[id!].dirty).toBe(true);
  expect(await useWorkspaceStore.getState().saveSession(id!)).toBe(true);
  expect(serialize).not.toHaveBeenCalled();
  const bytes = adapter.saveBytes.mock.calls[0][0];
  const JSZip = (await import('jszip')).default;
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('text');
  expect(xml).toContain('同一新核心保存正文');
  expect(useWorkspaceStore.getState().sessions[id!].dirty).toBe(false);
});

it('保存期间同一对象继续编辑保持 dirty，拒绝重复保存和写盘时强制关闭', async () => {
  const id = await open();
  const session = useWorkspaceStore.getState().sessions[id];
  session.wordCore!.editor.commands.insertContent('开始保存');
  let release!: () => void;
  adapter.saveBytes.mockImplementation(async () => { await new Promise<void>(resolve => { release = resolve; }); return { displayName: '保存.docx', path: '/tmp/core.docx', location: 'in-place' }; });
  const saving = useWorkspaceStore.getState().saveSession(id);
  await vi.waitFor(() => expect(adapter.saveBytes).toHaveBeenCalled());
  expect(await useWorkspaceStore.getState().saveSession(id)).toBe(false);
  expect(await useWorkspaceStore.getState().closeTab(id, true)).toBe(false);
  session.wordCore!.editor.commands.insertContent('保存期间的新内容');
  expect(useWorkspaceStore.getState().sessions[id].data).toBe(session.data);
  release(); expect(await saving).toBe(true);
  expect(useWorkspaceStore.getState().sessions[id]).toMatchObject({ dirty: true, state: 'DIRTY' });
  adapter.saveBytes.mockResolvedValue({ displayName: '保存.docx', path: '/tmp/core.docx', location: 'in-place' });
  expect(await useWorkspaceStore.getState().saveSession(id)).toBe(true);
  expect(useWorkspaceStore.getState().sessions[id].dirty).toBe(false);
});

it('失败或取消不确认版本，恢复快照保存为副本后才清除 dirty', async () => {
  const id = await open();
  const core = useWorkspaceStore.getState().sessions[id].wordCore!;
  core.editor.commands.insertContent('不能丢失');
  const before = core.state().savedRevision;
  adapter.saveBytes.mockRejectedValue(new Error('写盘失败'));
  expect(await useWorkspaceStore.getState().saveSession(id)).toBe(false);
  expect(core.state()).toMatchObject({ dirty: true, savedRevision: before });
  adapter.saveBytes.mockRejectedValue(new OfficeError(ErrorCodes.CANCELLED, '取消保存'));
  expect(await useWorkspaceStore.getState().saveSession(id, { asNew: true })).toBe(false);
  expect(core.state()).toMatchObject({ dirty: true, savedRevision: before });
  adapter.saveBytes.mockResolvedValue({ displayName: '保存.docx', path: '/tmp/core.docx', location: 'in-place' });
  const source = (await core.prepareSave()).bytes;
  const restored = await useWorkspaceStore.getState().openBytes({ name: '恢复.docx', bytes: source, restored: { dirty: true } });
  expect(useWorkspaceStore.getState().sessions[restored!]).toMatchObject({ dirty: true, saveAsCopyRequired: true });
  expect(await useWorkspaceStore.getState().saveSession(restored!)).toBe(true);
  expect(adapter.saveBytes.mock.calls.at(-1)?.[1].forcePicker).toBe(true);
  expect(useWorkspaceStore.getState().sessions[restored!].dirty).toBe(false);
});

it('正文卸载及标签切换保留原会话与撤销历史，关闭后才释放', async () => {
  const id = await open(); const core = useWorkspaceStore.getState().sessions[id].wordCore!;
  const root = document.createElement('div'); document.body.append(root);
  const first = await WordCoreSurface.attach(root, core);
  const original = core.editor.getText();
  core.editor.commands.insertContent('标签切换保留内容');
  first.dispose();
  expect(core.editor.isDestroyed).toBe(false);
  // 功能：非活动标签仍从自己的核心保存，不依赖当前活动编辑器的 DOM；时间：2026-11-09；作者：lq。
  const another = await useWorkspaceStore.getState().createDocument('docx');
  expect(useWorkspaceStore.getState().activeId).toBe(another);
  expect(await useWorkspaceStore.getState().saveSession(id)).toBe(true);
  useWorkspaceStore.getState().activateTab(id);
  const second = await WordCoreSurface.attach(root, core);
  expect(second.session).toBe(core); expect(core.editor.getText()).toContain('标签切换保留内容');
  core.editor.commands.undo(); expect(core.editor.getText()).toBe(original);
  second.dispose();
  expect(await useWorkspaceStore.getState().closeTab(id, true)).toBe(true);
  expect(core.editor.isDestroyed).toBe(true);
});

it('原生保护文档在事务层只读，不能通过工作区写盘', async () => {
  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(readFileSync('tests/fixtures/docx/contract.docx'));
  zip.file('word/settings.xml', '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>');
  const id = await useWorkspaceStore.getState().openBytes({ name: '保护.docx', bytes: await zip.generateAsync({ type: 'uint8array' }) });
  const session = useWorkspaceStore.getState().sessions[id!];
  expect(session.readonly).toBe(true);
  const text = session.wordCore!.editor.getText(); session.wordCore!.editor.commands.insertContent('绕过保护');
  expect(session.wordCore!.editor.getText()).toBe(text);
  expect(await useWorkspaceStore.getState().saveSession(id!)).toBe(false);
});

// 功能：重现慢快照写入期间继续输入并再次请求快照，后续请求不能被旧任务吞掉；时间：2026-11-09；作者：lq。
it('快照任务进行中继续编辑，完成后补存最新版本且不提前清除 dirty', async () => {
  const id = await open();
  const core = useWorkspaceStore.getState().sessions[id].wordCore!;
  useWorkspaceStore.getState().updateSettings({ autoSave: true, autoSaveDelayMs: 500 });
  let release!: () => void;
  const saveSnapshot = vi.fn<PlatformAdapter['saveSnapshot']>().mockResolvedValue(undefined);
  saveSnapshot.mockImplementationOnce(async () => { await new Promise<void>(resolve => { release = resolve; }); });
  adapter.saveSnapshot = saveSnapshot;
  core.editor.commands.insertContent('快照版本一');
  await vi.waitFor(() => expect(saveSnapshot).toHaveBeenCalledTimes(1), { timeout: 3000 });
  core.editor.commands.insertContent('快照版本二');
  // 功能：等待第二次自动保存请求确实进入队列，再释放慢 I/O；时间：2026-11-09；作者：lq。
  await new Promise(resolve => setTimeout(resolve, 650));
  release();
  await vi.waitFor(() => expect(saveSnapshot).toHaveBeenCalledTimes(2), { timeout: 3000 });
  const JSZip = (await import('jszip')).default;
  const xml = await (await JSZip.loadAsync(saveSnapshot.mock.calls[1][0].bytes)).file('word/document.xml')!.async('text');
  expect(xml).toContain('快照版本二');
  expect(core.state().dirty).toBe(true);
});

// 功能：宏和模板原包保留时，同步验证另存名称和原始内容类型，避免扩展名与容器不一致；时间：2026-11-09；作者：lq。
it.each([
  ['宏文档.docm', 'application/vnd.ms-word.document.macroEnabled.main+xml'],
  ['模板.dotx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'],
])('新核心保存 %s 保留包类型与扩展名', async (name, contentType) => {
  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(readFileSync('tests/fixtures/docx/contract.docx'));
  const types = await zip.file('[Content_Types].xml')!.async('text');
  zip.file('[Content_Types].xml', types.replace('application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', contentType));
  const id = await useWorkspaceStore.getState().openBytes({ name, bytes: await zip.generateAsync({ type: 'uint8array' }) });
  expect(id).toBeTruthy();
  expect(await useWorkspaceStore.getState().saveSession(id!, { asNew: true })).toBe(true);
  expect(adapter.saveBytes.mock.calls[0][1].suggestedName).toBe(name);
  const saved = await JSZip.loadAsync(adapter.saveBytes.mock.calls[0][0]);
  expect(await saved.file('[Content_Types].xml')!.async('text')).toContain(contentType);
});
