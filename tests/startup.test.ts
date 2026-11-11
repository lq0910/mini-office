/** 功能：验证慢字体扫描不阻塞首页、并发初始化去重及首帧交接时序；时间：2026-10-24；作者：lq。 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { PlatformAdapter } from '@mini-office/platform-adapter';
import type { FontInfo } from '@mini-office/shared';

/** 功能：构造受控异步结果，复现冷启动扫描和 IndexedDB 等待；时间：2026-10-24；作者：lq。 */
function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

/** 功能：隔离原生文件、字体和快照访问，不读写用户文档；时间：2026-10-24；作者：lq。 */
function platform(): PlatformAdapter {
  return {
    kind: 'web', openPath: vi.fn(), openFile: vi.fn(), saveBytes: vi.fn(),
    saveSnapshot: vi.fn().mockResolvedValue(undefined), listSnapshots: vi.fn().mockResolvedValue([]),
    loadSnapshot: vi.fn(), deleteSnapshot: vi.fn(), listFonts: vi.fn().mockResolvedValue([]),
    revealInFolder: vi.fn(),
  };
}

beforeEach(() => { vi.resetModules(); });
afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
  document.body.innerHTML = '';
});

describe('工作区启动', () => {
  it('并发初始化只读取一次快照，首页就绪无需等待字体扫描', async () => {
    const { setPlatformAdapter } = await import('@mini-office/platform-adapter');
    const { useWorkspaceStore } = await import('@mini-office/document-model');
    const adapter = platform();
    const snapshots = deferred<[]>();
    const fonts = deferred<FontInfo[]>();
    vi.mocked(adapter.listSnapshots).mockReturnValue(snapshots.promise);
    vi.mocked(adapter.listFonts).mockReturnValue(fonts.promise);
    setPlatformAdapter(adapter);
    const first = useWorkspaceStore.getState().initialize();
    const second = useWorkspaceStore.getState().initialize();
    expect(adapter.listSnapshots).toHaveBeenCalledOnce();
    snapshots.resolve([]);
    await Promise.all([first, second]);
    expect(useWorkspaceStore.getState().hydrated).toBe(true);
    expect(adapter.listFonts).not.toHaveBeenCalled();
    const loadA = useWorkspaceStore.getState().loadFonts();
    const loadB = useWorkspaceStore.getState().loadFonts();
    await Promise.resolve();
    expect(adapter.listFonts).toHaveBeenCalledOnce();
    expect(useWorkspaceStore.getState().hydrated).toBe(true);
    fonts.resolve([{ family: '宋体', source: 'system', hasChinese: true }]);
    await Promise.all([loadA, loadB]);
    expect(useWorkspaceStore.getState().fonts[0].family).toBe('宋体');
  });

  it('打开文档使用扫描后的真实字体清单，扫描失败可重试', async () => {
    const { setPlatformAdapter } = await import('@mini-office/platform-adapter');
    const { useWorkspaceStore, registerFormatHandler, createBlankDoc } = await import('@mini-office/document-model');
    const adapter = platform();
    vi.mocked(adapter.listFonts).mockRejectedValueOnce(new Error('临时不可用'));
    setPlatformAdapter(adapter);
    expect(await useWorkspaceStore.getState().loadFonts()).toEqual([]);
    const fonts = deferred<FontInfo[]>();
    vi.mocked(adapter.listFonts).mockReturnValue(fonts.promise);
    const open = vi.fn().mockResolvedValue({
      data: createBlankDoc(), capability: { level: 3, warnings: [], features: [], missingFonts: [], scannedAt: 0 },
    });
    registerFormatHandler({ kind: 'docx', extensions: ['.docx'], createBlank: createBlankDoc, open, serialize: vi.fn() });
    const opening = useWorkspaceStore.getState().openBytes({ name: '合同.docx', bytes: new Uint8Array([1]) });
    await vi.waitFor(() => expect(adapter.listFonts).toHaveBeenCalledTimes(2));
    expect(open).not.toHaveBeenCalled();
    fonts.resolve([{ family: '宋体', source: 'system', hasChinese: true }]);
    expect(await opening).toBeTruthy();
    expect(open).toHaveBeenCalledWith(expect.objectContaining({ knownFonts: ['宋体'] }));
  });

  // 功能：默认 Word 使用真实会话快照，禁止以固定标识调用旧 serialize 旁路；时间：2026-11-11；作者：lq。
  it.each(['docx', 'xlsx', 'pptx'] as const)('延迟注册的 %s 格式支持新建、序列化和重新打开', async (kind) => {
    const { registerDeferredFormats } = await import('../apps/desktop/src/format-handlers');
    const { getFormatHandler } = await import('@mini-office/document-model');
    registerDeferredFormats();
    const handler = getFormatHandler(kind)!;
    const name = `启动验证.${kind}`;
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
    const created = await handler.createNew?.();
    const bytes = created?.wordCore ? (await created.wordCore.prepareSave()).bytes : await handler.serialize(handler.createBlank(), { name });
    expect(bytes.byteLength).toBeGreaterThan(0);
    const reopened = await handler.open({ name, bytes, knownFonts: [] });
    expect(reopened.data).toBeTruthy();
    if (kind === 'docx') { expect(reopened.data).toEqual({ engine: 'word-core' }); expect(reopened.wordCore).toBeTruthy(); }
    created?.wordCore?.dispose(); reopened.wordCore?.dispose();
  });
});

describe('启动面板交接', () => {
  it('主界面经过绘制机会后移除面板，取消挂载时不会提前移除', async () => {
    vi.useFakeTimers();
    const frames = new Map<number, FrameRequestCallback>();
    let sequence = 0;
    vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
      frames.set(++sequence, callback);
      return sequence;
    });
    vi.stubGlobal('cancelAnimationFrame', (id: number) => { frames.delete(id); });
    const { finishStartup, dismissStartup } = await import('../apps/desktop/src/startup');
    document.body.innerHTML = '<div id="startup-splash"></div><div class="mo-app"></div>';
    const cancel = finishStartup();
    cancel();
    expect(frames.size).toBe(0);
    expect(document.getElementById('startup-splash')).not.toBeNull();
    finishStartup();
    const flushFrame = () => {
      const pending = [...frames.values()];
      frames.clear();
      pending.forEach((callback) => callback(0));
    };
    flushFrame();
    expect(document.getElementById('startup-splash')).not.toBeNull();
    flushFrame();
    expect(document.getElementById('startup-splash')).toBeNull();
    document.body.innerHTML = '<div id="startup-splash"></div>';
    dismissStartup();
    expect(document.getElementById('startup-splash')).toBeNull();
  });
});
