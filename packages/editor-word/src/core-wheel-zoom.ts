/** 功能：Ctrl+鼠标滚轮缩放当前 Word 文档，普通滚轮继续翻页，缩放不写入文档历史；时间：2026-11-11；作者：lq。 */
import { useEffect, type RefObject } from 'react';
import { useWorkspaceStore } from '@mini-office/document-model';

/** 功能：非 passive 原生监听阻止浏览器整窗缩放；像素/行/页滚轮统一为 10% 步长并保持鼠标下正文位置；时间：2026-11-11；作者：lq。 */
export function useWordCoreWheelZoom(scroll: RefObject<HTMLDivElement>, sessionId: string): void {
  useEffect(() => {
    const viewport = scroll.current; if (!viewport) return;
    let delta = 0, time = 0, frame: number | undefined;
    const wheel = (event: WheelEvent) => {
      if (!event.ctrlKey || event.altKey || event.shiftKey || event.defaultPrevented || !event.cancelable || document.querySelector('[role="dialog"][aria-modal="true"]')) return;
      event.preventDefault(); event.stopPropagation();
      const value = event.deltaY * (event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? viewport.clientHeight : 1);
      if (!value) return;
      if (performance.now() - time > 180 || Math.sign(value) !== Math.sign(delta)) delta = 0;
      time = performance.now(); delta += value;
      if (Math.abs(delta) < 32) return;
      const state = useWorkspaceStore.getState(), session = state.sessions[sessionId]; if (!session) return;
      const before = session.view.zoom ?? 1, next = Math.round(Math.max(0.4, Math.min(2.5, before + (delta < 0 ? 0.1 : -0.1))) * 100) / 100;
      delta = 0; if (before === next) return;
      const content = viewport.querySelector<HTMLElement>('.doc-zoom'), box = content?.getBoundingClientRect();
      const x = box ? (event.clientX - box.left) / before : 0, y = box ? (event.clientY - box.top) / before : 0;
      state.updateView(sessionId, { zoom: next });
      if (frame !== undefined) cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => {
        frame = undefined; if (!content?.isConnected || !box) return;
        const after = content.getBoundingClientRect();
        viewport.scrollLeft += after.left + x * next - event.clientX;
        viewport.scrollTop += after.top + y * next - event.clientY;
      });
    };
    viewport.addEventListener('wheel', wheel, { capture: true, passive: false });
    return () => { viewport.removeEventListener('wheel', wheel, true); if (frame !== undefined) cancelAnimationFrame(frame); };
  }, [scroll, sessionId]);
}
