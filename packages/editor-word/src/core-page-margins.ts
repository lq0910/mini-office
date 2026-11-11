/** 功能：页面边距拖动使用原生分节设置，移动仅预览、松手一次历史提交；时间：2026-11-11；作者：lq。 */
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import type { WordPageFrame } from './core-layout';
type Edge = 'left' | 'right' | 'top' | 'bottom';
/** 功能：独立显示层不进入可编辑 DOM，取消/旧版本/卸载均不修改正文；时间：2026-11-11；作者：lq。 */
export class WordPageMarginGuides {
  private readonly layer = document.createElement('div');
  private cancel?: () => void;
  private signature = '';
  private frames: readonly WordPageFrame[] = [];
  private readonly guides: { element: HTMLElement; frame: WordPageFrame }[] = [];
  constructor(private readonly core: WordCoreSession, private readonly wrap: HTMLElement, private readonly onError: (error: unknown) => void) {
    this.layer.className = 'mo-core-margin-layer mo-no-print'; wrap.append(this.layer);
  }
  /** 功能：按实际分页位置及分节纸张绘制四侧手柄，镜像和竖排保留设置面板入口；时间：2026-11-11；作者：lq。 */
  refresh(frames: readonly WordPageFrame[]): void {
    this.frames = frames;
    if (this.cancel) return;
    const sections = this.core.sections(), signature = JSON.stringify([frames, sections.map(section => section.settings)]);
    if (signature === this.signature) return;
    this.signature = signature;
    if (!this.core.editor.isEditable || this.core.parsed.mirrorMargins) { this.layer.replaceChildren(); this.guides.length = 0; return; }
    // 功能：复用每页手柄，异步重排/滚动更新位置不会销毁正在定位的元素；时间：2026-11-11；作者：lq。
    while (this.guides.length > frames.length) this.guides.pop()!.element.remove();
    frames.forEach((frame, index) => {
      const settings = sections[frame.section].settings;
      let record = this.guides[index];
      if (!record) {
        const element = document.createElement('div'); element.className = 'mo-core-margin-guide';
        record = { element, frame }; this.guides.push(record); this.layer.append(element);
        for (const edge of ['left','right','top','bottom'] as const) {
          const button = document.createElement('button'); button.type = 'button'; button.className = 'mo-core-margin-handle'; button.dataset.edge = edge;
          const name = { left: '左', right: '右', top: '上', bottom: '下' }[edge];
          button.setAttribute('aria-label', `第 ${index+1} 页${name}页边距`); button.title = `拖动调整${name}页边距`;
          const target = record;
          button.addEventListener('pointerdown', event => this.start(event, edge, target.frame, target.element)); element.append(button);
        }
      }
      record.frame = frame; const guide = record.element;
      guide.hidden = !!settings.textDirection && settings.textDirection !== 'lrTb';
      guide.style.cssText = `left:${frame.left + settings.marginLeft / 15}px;top:${frame.top + settings.marginTop / 15}px;width:${(settings.pageWidth-settings.marginLeft-settings.marginRight)/15}px;height:${(settings.pageHeight-settings.marginTop-settings.marginBottom)/15}px`;
    });
  }
  /** 功能：以真实缩放换算 twips，预览过程不排版、不增加 dirty，Esc/取消恢复；时间：2026-11-11；作者：lq。 */
  private start(event: PointerEvent, edge: Edge, frame: WordPageFrame, guide: HTMLElement): void {
    if (event.button !== 0 || this.core.commandEditor().view.composing) return;
    event.preventDefault(); event.stopPropagation(); this.cancel?.();
    const document = this.core.editor.state.doc, sections = this.core.sections(), section = sections[frame.section], settings = section.settings;
    const key = ({ left: 'marginLeft', right: 'marginRight', top: 'marginTop', bottom: 'marginBottom' } as const)[edge];
    const base = settings[key], horizontal = edge === 'left' || edge === 'right';
    const start = horizontal ? event.clientX : event.clientY, factor = this.wrap.getBoundingClientRect().width / this.wrap.offsetWidth || 1;
    const opposite = settings[({ left: 'marginRight', right: 'marginLeft', top: 'marginBottom', bottom: 'marginTop' } as const)[edge]];
    const maximum = (horizontal ? settings.pageWidth : settings.pageHeight)-opposite-720;
    const minimum = edge === 'left' && !settings.gutterAtTop || edge === 'top' && settings.gutterAtTop ? settings.gutter ?? 0 : 0;
    let value = base; const original = guide.style.cssText; guide.classList.add('dragging');
    const move = (pointer: PointerEvent) => {
      const delta = ((horizontal ? pointer.clientX : pointer.clientY)-start)/factor*15*(edge === 'right' || edge === 'bottom' ? -1 : 1);
      value = Math.round(Math.min(maximum, Math.max(minimum, base+delta)));
      const left = edge === 'left' ? value : settings.marginLeft, top = edge === 'top' ? value : settings.marginTop;
      const right = edge === 'right' ? value : settings.marginRight, bottom = edge === 'bottom' ? value : settings.marginBottom;
      guide.style.cssText = `left:${frame.left+left/15}px;top:${frame.top+top/15}px;width:${(settings.pageWidth-left-right)/15}px;height:${(settings.pageHeight-top-bottom)/15}px`;
    };
    const cleanup = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', finish); window.removeEventListener('pointercancel', cancel); window.removeEventListener('keydown', keydown); guide.style.cssText = original; guide.classList.remove('dragging'); this.cancel = undefined; this.signature = ''; this.refresh(this.frames); };
    const cancel = () => cleanup();
    const keydown = (keyEvent: KeyboardEvent) => { if (keyEvent.key === 'Escape') { keyEvent.preventDefault(); cancel(); } };
    const finish = () => {
      cleanup(); if (value === base || this.core.editor.state.doc !== document) return;
      try { const catalog = this.core.pageLayout(); this.core.commitPageLayout({ ...catalog, sections: { ...catalog.sections, [section.lastBlockIndex]: { ...catalog.sections[section.lastBlockIndex], [key]: value } } }); }
      catch (error) { this.onError(error); }
    };
    this.cancel = cancel; window.addEventListener('pointermove', move); window.addEventListener('pointerup', finish); window.addEventListener('pointercancel', cancel); window.addEventListener('keydown', keydown);
  }
  /** 功能：宿主卸载移除拖动监听和显示层，已提交历史仍归主会话；时间：2026-11-11；作者：lq。 */
  dispose(): void { this.cancel?.(); this.layer.remove(); }
}
