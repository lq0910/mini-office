/** 功能：验证保存期间继续编辑、并发保存和访达重复打开时保留文档；时间：2026-10-18；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createBlankDoc, registerFormatHandler, useWorkspaceStore, type DocumentSession } from '@mini-office/document-model';
import { OfficeError, ErrorCodes } from '@mini-office/shared';
import { setPlatformAdapter, type PlatformAdapter, type SaveOutcome } from '@mini-office/platform-adapter';

/** 功能：提供可控制异步写盘的适配器，不写入用户文件；时间：2026-10-18；作者：lq。 */
function adapter() {
  return {
    kind: 'tauri', openPath: vi.fn(), openFile: vi.fn(), saveBytes: vi.fn(),
    saveSnapshot: vi.fn().mockResolvedValue(undefined), listSnapshots: vi.fn().mockResolvedValue([]),
    loadSnapshot: vi.fn(), deleteSnapshot: vi.fn(), listFonts: vi.fn(), revealInFolder: vi.fn(),
  } satisfies PlatformAdapter;
}

beforeEach(() => {
  registerFormatHandler({
    kind: 'docx', extensions: ['.docx'], createBlank: createBlankDoc,
    open: vi.fn(), serialize: async (data) => new TextEncoder().encode(JSON.stringify(data)),
  });
  useWorkspaceStore.setState({ sessions: {}, tabOrder: [], activeId: null, notifications: [] });
});
afterEach(() => setPlatformAdapter(null));

/** 功能：注入有已知磁盘路径的未保存会话；时间：2026-10-18；作者：lq。 */
function seed(id: string) {
  const session = {
    id, kind: 'docx', displayName: '合同.docx', path: `/tmp/${id}.docx`, data: createBlankDoc(),
    dirty: true, state: 'DIRTY', readonly: false, capability: { level: 3 },
  } as DocumentSession;
  useWorkspaceStore.setState({ sessions: { [id]: session }, tabOrder: [id], activeId: id });
  return session;
}

describe('安全保存生命周期', () => {
  // 功能：面板与内部刷新不能清除已有未保存正文的退出保护；时间：2026-11-03；作者：lq。
  it('非用户刷新保留已有 dirty 状态', () => {
    seed('keep-dirty');
    useWorkspaceStore.getState().updateDocMeta('keep-dirty', { mailMergeState: '{}' });
    expect(useWorkspaceStore.getState().sessions['keep-dirty'].dirty).toBe(true);
    useWorkspaceStore.getState().updateData('keep-dirty', (data) => data, { dirty: false });
    expect(useWorkspaceStore.getState().sessions['keep-dirty'].state).toBe('DIRTY');
  });
  // 功能：快捷保存及另存为共享现有保存链路，不绕过路径选择策略；时间：2026-10-30；作者：lq。
  it.each([
    { id: 'quick-save-existing', newDocument: false, asNew: false, forcePicker: false },
    { id: 'quick-save-new', newDocument: true, asNew: false, forcePicker: true },
    { id: 'quick-save-as', newDocument: false, asNew: true, forcePicker: true },
  ])('快捷保存遵循原位保存/新建/另存为策略：$id', async ({ id, newDocument, asNew, forcePicker }) => {
    const platform = adapter();
    platform.saveBytes.mockResolvedValue({ displayName: '合同.docx', path: `/tmp/${id}-saved.docx`, location: 'in-place' });
    setPlatformAdapter(platform);
    const session = seed(id);
    if (newDocument) {
      useWorkspaceStore.setState({ sessions: { [id]: { ...session, path: undefined } } });
    }
    expect(await useWorkspaceStore.getState().saveActive(asNew ? { asNew: true } : undefined)).toBe(true);
    expect(platform.saveBytes).toHaveBeenCalledOnce();
    expect(platform.saveBytes.mock.calls[0][1]).toMatchObject({
      suggestedName: '合同.docx', path: newDocument ? undefined : session.path, forcePicker,
    });
    expect(useWorkspaceStore.getState().sessions[id].state).toBe('SAVED');
    expect(useWorkspaceStore.getState().sessions[id].dirty).toBe(false);
  });

  it('内部刷新不置脏，只有用户编辑才标记未保存', () => {
    seed('internal-refresh');
    useWorkspaceStore.setState((state) => ({ sessions: { ...state.sessions, 'internal-refresh': { ...state.sessions['internal-refresh'], dirty: false, state: 'EDIT_READY' } } }));
    const blank = createBlankDoc();
    useWorkspaceStore.getState().updateData('internal-refresh', createBlankDoc());
    expect(useWorkspaceStore.getState().sessions['internal-refresh'].dirty).toBe(false);
    expect(useWorkspaceStore.getState().sessions['internal-refresh'].state).toBe('EDIT_READY');
    useWorkspaceStore.getState().updateData('internal-refresh', blank, { dirty: true });
    expect(useWorkspaceStore.getState().sessions['internal-refresh'].dirty).toBe(true);
  });

  it('写盘等待期间新增内容，写盘结束仍保持未保存', async () => {
    const platform = adapter();
    let finish!: (value: SaveOutcome) => void;
    platform.saveBytes.mockImplementation(() => new Promise<SaveOutcome>((resolve) => { finish = resolve; }));
    setPlatformAdapter(platform);
    seed('concurrent-edit');
    const saving = useWorkspaceStore.getState().saveSession('concurrent-edit');
    await vi.waitFor(() => expect(platform.saveBytes).toHaveBeenCalledOnce());
    useWorkspaceStore.getState().updateData('concurrent-edit', createBlankDoc(), { dirty: true });
    expect(useWorkspaceStore.getState().sessions['concurrent-edit'].state).toBe('SAVING');
    expect(await useWorkspaceStore.getState().saveSession('concurrent-edit')).toBe(false);
    finish({ displayName: '合同.docx', path: '/tmp/concurrent-edit.docx', location: 'in-place' });
    expect(await saving).toBe(true);
    expect(useWorkspaceStore.getState().sessions['concurrent-edit'].dirty).toBe(true);
    expect(useWorkspaceStore.getState().sessions['concurrent-edit'].state).toBe('DIRTY');
  });
  it('重复保存不并发写盘，无新编辑时正确清除未保存标记', async () => {
    const platform = adapter();
    let finish!: (value: SaveOutcome) => void;
    platform.saveBytes.mockImplementation(() => new Promise<SaveOutcome>((resolve) => { finish = resolve; }));
    setPlatformAdapter(platform);
    seed('duplicate-save');
    const saving = useWorkspaceStore.getState().saveSession('duplicate-save');
    expect(await useWorkspaceStore.getState().saveSession('duplicate-save')).toBe(false);
    await vi.waitFor(() => expect(platform.saveBytes).toHaveBeenCalledOnce());
    finish({ displayName: '合同.docx', path: '/tmp/duplicate-save.docx', location: 'in-place' });
    expect(await saving).toBe(true);
    expect(useWorkspaceStore.getState().sessions['duplicate-save'].dirty).toBe(false);
  });
  it.each([true, false])('保存取消/失败后保留未保存状态，并允许重试：%s', async (cancelled) => {
    const platform = adapter();
    platform.saveBytes.mockRejectedValueOnce(cancelled
      ? new OfficeError(ErrorCodes.CANCELLED, '保存已取消')
      : new Error('磁盘写入失败'));
    setPlatformAdapter(platform);
    seed('failed-save');
    expect(await useWorkspaceStore.getState().saveSession('failed-save')).toBe(false);
    expect(useWorkspaceStore.getState().sessions['failed-save'].dirty).toBe(true);
    platform.saveBytes.mockResolvedValueOnce({ displayName: '合同.docx', path: '/tmp/failed-save.docx', location: 'in-place' });
    expect(await useWorkspaceStore.getState().saveSession('failed-save')).toBe(true);
    expect(useWorkspaceStore.getState().sessions['failed-save'].dirty).toBe(false);
  });
  it('访达打开已有文件只激活标签，保留未保存模型', async () => {
    const platform = adapter();
    setPlatformAdapter(platform);
    const session = seed('existing-file');
    expect(await useWorkspaceStore.getState().openPath(session.path!)).toBe(session.id);
    expect(platform.openPath).not.toHaveBeenCalled();
    expect(useWorkspaceStore.getState().sessions[session.id].data).toBe(session.data);
    expect(useWorkspaceStore.getState().sessions[session.id].dirty).toBe(true);
  });
});
