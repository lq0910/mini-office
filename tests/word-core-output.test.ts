/** 功能：验证原生输出页码范围及拖动取消/卸载边界；时间：2026-11-11；作者：lq。 */
import { describe, expect, it, vi } from 'vitest';
import { validateCorePageRange } from '../packages/editor-word/src/WordCoreOutput';
import { attachPreviewMargins } from '../packages/platform-adapter/src/print';
/** 功能：jsdom 用携带真实指针字段的事件验证监听器生命周期，不模拟提交实现；时间：2026-11-11；作者：lq。 */
function pointer(target: EventTarget, type: string, x: number, id = 1) { const event = new MouseEvent(type, { bubbles: true, button: 0, clientX: x, clientY: 100 }); Object.defineProperty(event, 'pointerId', { value: id }); target.dispatchEvent(event); }
describe('原生 Word 输出', () => {
  it('页码范围严格拒绝越界、倒序、小数和混合非法项，允许重复但由导出服务去重', () => {
    for (const range of ['', '1', '1-3,5', '1–3', '2,2,4-5']) expect(() => validateCorePageRange(range, 5)).not.toThrow();
    for (const range of ['0', '6', '2-1', '1-6', '1,错误', '1.5', ',1', '1,', '1e2', '-1', '9'.repeat(1001)]) expect(() => validateCorePageRange(range, 5)).toThrow();
  });
  it('页边距只有有效指针松开时提交，取消、其他指针和卸载不会提交', () => {
    const frame = document.createElement('div'), page = document.createElement('div'); frame.append(page); document.body.append(frame);
    const change = vi.fn(); const dispose = attachPreviewMargins({ frame, page, scale: 1, index: 0 }, { margins: { left: 96, right: 96, top: 96, bottom: 96 }, pageWidthPx: 794, pageHeightPx: 1123, onChange: change });
    const handle = frame.querySelector('[aria-label="拖动调整左页边距"]')!;
    pointer(handle, 'pointerdown', 96); pointer(window, 'pointermove', 130); pointer(window, 'pointercancel', 130); expect(change).not.toHaveBeenCalled();
    pointer(handle, 'pointerdown', 96); pointer(window, 'pointermove', 130, 2); pointer(window, 'pointerup', 130, 2); expect(change).not.toHaveBeenCalled(); pointer(window, 'pointermove', 130); pointer(window, 'pointerup', 130); expect(change).toHaveBeenLastCalledWith('left', 130);
    change.mockClear(); pointer(handle, 'pointerdown', 130); pointer(window, 'pointermove', 150); dispose(); pointer(window, 'pointerup', 150); expect(change).not.toHaveBeenCalled(); expect(frame.querySelector('[data-preview-margins]')).toBeNull(); frame.remove();
  });
});
