/** 功能：验证产品正文宿主的真实核心、取消和写盘生命周期；时间：2026-11-08；作者：lq。 */
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { WordCoreSurface } from '../packages/editor-word/src/core-surface';
import { WordCoreSurfaceView } from '../packages/editor-word/src/WordCoreSurface';
const mounted: WordCoreSurface[] = [];
beforeEach(() => {
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => setTimeout(() => callback(performance.now()), 0));
  vi.stubGlobal('cancelAnimationFrame', (id: number) => clearTimeout(id));
  // 功能：jsdom 缺少 CSS.escape，测试样本标识符使用等价的 CSS 十六进制转义；时间：2026-11-08；作者：lq。
  vi.stubGlobal('CSS', { escape: (value: string) => Array.from(value, char => `\\${char.codePointAt(0)!.toString(16)} `).join('') });
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
});
afterEach(() => { mounted.splice(0).forEach(surface => surface.dispose()); document.body.replaceChildren(); vi.restoreAllMocks(); vi.unstubAllGlobals(); });
/** 功能：在真实产品容器里打开公共样本，保留旁边的工具栏节点；时间：2026-11-08；作者：lq。 */
async function open() {
  const root = document.createElement('div');
  const toolbar = document.createElement('nav');
  root.append(toolbar);
  document.body.append(root);
  const surface = await WordCoreSurface.mount(root, readFileSync('tests/fixtures/docx/contract.docx'));
  mounted.push(surface);
  return { root, toolbar, surface };
}
it('挂载实际核心并清理自身节点，保留现有 UI', async () => {
  expect(WordCoreSurfaceView).toBeTypeOf('function');
  const { root, toolbar, surface } = await open();
  expect(root.querySelector('.ProseMirror')).toBeTruthy();
  // 功能：jsdom 没有实际几何，宿主不伪造页数；释放后的同步重排安全返回；时间：2026-11-09；作者：lq。
  expect(surface.layout.current()).toBeUndefined();
  surface.dispose(); surface.dispose();
  expect(root.children.length).toBe(1);
  expect(root.firstChild).toBe(toolbar);
  expect(surface.layout.flush()).toBeUndefined();
  await expect(surface.save(async () => {})).rejects.toThrow('已关闭');
});
it('失败不确认，写盘期间编辑保持 dirty，下一次成功保存才清除', async () => {
  const { surface } = await open();
  surface.session.editor.commands.insertContent('智慧物业信息系统部署实施');
  await expect(surface.save(async () => { throw new Error('磁盘失败'); })).rejects.toThrow('磁盘失败');
  expect(surface.session.state().dirty).toBe(true);
  let release!: () => void;
  let started!: () => void;
  const ready = new Promise<void>(resolve => { started = resolve; });
  const pending = surface.save(async () => { started(); await new Promise<void>(resolve => { release = resolve; }); });
  await ready;
  await expect(surface.save(async () => {})).rejects.toThrow('正在保存');
  surface.session.editor.commands.insertContent('继续编辑');
  release(); await pending;
  expect(surface.session.state().dirty).toBe(true);
  await surface.save(async bytes => { expect(bytes.byteLength).toBeGreaterThan(0); });
  expect(surface.session.state().dirty).toBe(false);
});
it('中止打开不遗留迟到编辑器，也不移除宿主节点', async () => {
  const root = document.createElement('div');
  const toolbar = document.createElement('nav'); root.append(toolbar);
  const abort = new AbortController();
  const pending = WordCoreSurface.mount(root, readFileSync('tests/fixtures/docx/contract.docx'), {}, abort.signal);
  abort.abort();
  await expect(pending).rejects.toThrow();
  expect(root.children.length).toBe(1);
  expect(root.firstChild).toBe(toolbar);
});
