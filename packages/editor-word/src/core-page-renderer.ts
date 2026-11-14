/** 功能：承接 GenOffice MultipageGrid 的活动页编辑窗口与邻页快照思路，滚动仅绘制窗口内页面；保留唯一 ProseMirror 正文，不使用 content-visibility；时间：2026-11-14；作者：lq。 */
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import type { WordSurfaceLayout, WordPageFrame } from './core-layout';
import { canvasZoomOf } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/pagination-gaps';

/** 功能：记录排版完成后的块几何，页面快照沿用原字体、表格和跨页装饰；时间：2026-11-14；作者：lq。 */
interface PaintBlock { element: HTMLElement; top: number; left: number; width: number; height: number; paintTop: number; paintBottom: number; }

/** 功能：显示层与编辑层分离，最多保留六页快照；原生拖选、输入法及输出临时使用完整编辑层；时间：2026-11-14；作者：lq。 */
export class WordPageRenderer {
  private readonly paper: HTMLElement;
  private readonly wrap: HTMLElement;
  private readonly viewport: HTMLElement | null;
  private stage?: HTMLDivElement;
  private live?: HTMLDivElement;
  private layer?: HTMLDivElement;
  private frames: readonly WordPageFrame[] = [];
  private blocks: PaintBlock[] = [];
  private readonly pages = new Map<number, HTMLDivElement>();
  private active = 0;
  private timer?: ReturnType<typeof setTimeout>;
  private paintFrame?: number;
  private holds = 0;
  private gesture = false;
  private disposed = false;
  private zoom = 1;
  private origin = 0;
  private viewportHeight = 0;
  private revision = -1;
  private readonly originalTop: string;
  private readonly unsubscribe: () => void;
  private drag?: { anchor: number; x: number; y: number };
  private pendingClick?: { target: Element; event: PointerEvent };
  // 功能：原生点击保留分页窗口，仅在确认拖选后展开全文；时间：2026-11-14；作者：lq。
  private nativeDrag?: { x: number; y: number };

  private readonly scroll = () => {
    if (this.paintFrame !== undefined) return;
    this.paintFrame = requestAnimationFrame(() => { this.paintFrame = undefined; this.paintVisible(); });
  };
  private readonly pointerDown = (event: PointerEvent) => {
    if (event.button !== 0) return;
    this.gesture = true; this.nativeDrag = { x: event.clientX, y: event.clientY };
  };
  private readonly pointerUp = (event: PointerEvent) => {
    if (!this.gesture) return;
    const click = this.pendingClick; this.pendingClick = undefined;
    // 功能：快照接管的按下事件需成对释放原编辑器的鼠标处理，单击仍交给原目录/对象事件；时间：2026-11-14；作者：lq。
    if (this.drag || click) {
      const target = document.elementFromPoint(event.clientX, event.clientY) ?? this.paper;
      target.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, clientX: event.clientX, clientY: event.clientY }));
    }
    if (event.type !== 'pointercancel' && click?.target.isConnected) click.target.dispatchEvent(new MouseEvent('click', { bubbles: true, detail: 1, clientX: event.clientX, clientY: event.clientY, ctrlKey: click.event.ctrlKey, metaKey: click.event.metaKey, shiftKey: click.event.shiftKey }));
    this.gesture = false; this.drag = undefined; this.nativeDrag = undefined; this.resume();
  };
  private readonly pointerMove = (event: PointerEvent) => {
    // 功能：点击页底不再先隐藏快照、展开连续纸张；跨页拖选达到阈值后才保留原生全文选区；时间：2026-11-14；作者：lq。
    if (this.nativeDrag && (event.buttons & 1) && Math.hypot(event.clientX - this.nativeDrag.x, event.clientY - this.nativeDrag.y) >= 3) { this.nativeDrag = undefined; this.suspend(); }
    if (!this.drag || !(event.buttons & 1) || this.session.editor.isDestroyed) return;
    if (Math.hypot(event.clientX - this.drag.x, event.clientY - this.drag.y) < 3) return;
    this.suspend();
    this.pendingClick = undefined;
    const hit = this.session.editor.view.posAtCoords({ left: event.clientX, top: event.clientY });
    if (hit) this.session.editor.commands.setTextSelection({ from: this.drag.anchor, to: hit.pos });
  };
  private readonly compositionStart = () => this.suspend();
  private readonly compositionEnd = () => this.refresh();

  constructor(private readonly session: WordCoreSession, private readonly layout: WordSurfaceLayout) {
    this.paper = session.editor.view.dom;
    this.wrap = this.paper.closest<HTMLElement>('.page-wrap')!;
    this.viewport = this.wrap.closest<HTMLElement>('.mo-scroll');
    this.originalTop = this.wrap.style.top;
    this.viewport?.addEventListener('scroll', this.scroll, { passive: true });
    this.wrap.addEventListener('pointerdown', this.pointerDown, true);
    this.wrap.addEventListener('compositionstart', this.compositionStart, true);
    this.wrap.addEventListener('compositionend', this.compositionEnd, true);
    document.addEventListener('pointerup', this.pointerUp);
    document.addEventListener('pointercancel', this.pointerUp);
    document.addEventListener('pointermove', this.pointerMove);
    this.unsubscribe = session.subscribe(state => {
      if (state.revision !== this.revision) { this.suspend(); this.refresh(); }
      else if (!this.gesture) this.resume();
    });
  }

  /** 功能：测量及交互时原位打开编辑层；页面坐标始终保持不变，不改正文样式或高度；时间：2026-11-14；作者：lq。 */
  suspend(): void {
    if (!this.live) return;
    this.live.classList.add('mo-word-live-full');
    this.live.style.top = '0px'; this.live.style.height = 'auto'; this.wrap.style.top = this.originalTop;
    if (this.layer) this.layer.hidden = true;
  }

  /** 功能：打印、PDF 与异步页面克隆持有完整显示，结束后恢复分页显示；时间：2026-11-14；作者：lq。 */
  hold(): () => void {
    this.holds++; this.suspend(); let released = false;
    return () => { if (released) return; released = true; this.holds--; this.refresh(); };
  }

  /** 功能：排版稳定后刷新几何；编辑快速路径没有新分页事件时也能恢复快照，滚动本身不测量全文；时间：2026-11-14；作者：lq。 */
  refresh(): void {
    if (this.disposed || !this.viewport || this.holds) return;
    if (this.timer !== undefined) clearTimeout(this.timer);
    this.timer = setTimeout(() => {
      this.timer = undefined;
      if (this.disposed || this.holds || this.session.editor.isDestroyed) return;
      if (this.session.editor.view.composing || this.session.commandEditor().view.composing) return;
      if (!this.layout.isSettled()) { this.refresh(); return; }
      const frames = this.layout.current()?.frames;
      if (!frames || frames.length < 10) { this.reset(); this.revision = this.session.state().revision; return; }
      this.suspend();
      this.frames = frames;
      this.zoom = canvasZoomOf(this.wrap.closest<HTMLElement>('.doc-zoom'), 1);
      const box = this.wrap.getBoundingClientRect(), viewportBox = this.viewport!.getBoundingClientRect();
      if (!box.width) return;
      this.origin = box.top - viewportBox.top + this.viewport!.scrollTop;
      this.viewportHeight = this.viewport!.clientHeight;
      const take = (element: HTMLElement): PaintBlock => {
        const r = element.getBoundingClientRect();
        let top = r.top, bottom = r.bottom;
        // 功能：零高度浮动锚点及页眉页脚层按实际可见子对象确定绘制范围，避免漏图；时间：2026-11-14；作者：lq。
        if (!r.height || element.classList.contains('mo-core-page-regions')) for (const child of element.querySelectorAll<HTMLElement>('.doc-protected,.doc-textbox,.doc-img-wrap,img,[data-hf-page]')) {
          const c = child.getBoundingClientRect(); if (c.height) { top = Math.min(top, c.top); bottom = Math.max(bottom, c.bottom); }
        }
        return { element, top: (r.top - box.top) / this.zoom, left: (r.left - box.left) / this.zoom, width: r.width / this.zoom, height: r.height / this.zoom, paintTop: (top - box.top) / this.zoom, paintBottom: (bottom - box.top) / this.zoom };
      };
      this.blocks = Array.from(this.paper.children, element => take(element as HTMLElement));
      // 功能：纸面与阴影由分页窗口统一承接，不再把整篇底纸层复制进单页；时间：2026-11-14；作者：lq。
      for (const element of this.wrap.children) if (element !== this.paper && !element.contains(this.paper) && !element.matches('style,script,.mo-no-print,.page-sheets')) this.blocks.push(take(element as HTMLElement));
      if (!this.stage) {
        const stage = document.createElement('div'); stage.className = 'mo-word-page-stage';
        this.wrap.before(stage);
        const live = document.createElement('div'); live.className = 'mo-word-live-window';
        const layer = document.createElement('div'); layer.className = 'mo-word-page-layer';
        // 功能：原编辑器保持在快照前，所有 first-match 编辑器查找继续命中真实正文；时间：2026-11-14；作者：lq。
        stage.append(live, layer); live.append(this.wrap);
        this.stage = stage; this.live = live; this.layer = layer;
      }
      this.stage.style.width = `${box.width / this.zoom}px`;
      const last = frames[frames.length - 1];
      this.stage.style.height = `${Math.max(box.height / this.zoom, last.top + last.heightPx)}px`;
      for (const page of this.pages.values()) page.remove(); this.pages.clear();
      this.revision = this.session.state().revision;
      this.resume();
    }, 100);
  }

  /** 功能：光标所在页由原编辑器绘制，其他页采用惰性快照；非空选区及输入法期间保留完整原生交互；时间：2026-11-14；作者：lq。 */
  private resume(): void {
    if (!this.live || this.disposed || this.holds || this.gesture || this.session.editor.isDestroyed) return;
    const editor = this.session.editor;
    if (!editor.state.selection.empty || editor.view.composing || this.session.commandEditor() !== editor || !this.layout.isSettled()) { this.suspend(); return; }
    try {
      const y = (editor.view.coordsAtPos(editor.state.selection.head).top - this.wrap.getBoundingClientRect().top) / this.zoom;
      this.active = this.pageAt(y);
    } catch { this.active = Math.min(this.active, this.frames.length - 1); }
    const frame = this.frames[this.active]; if (!frame) return;
    this.showActivePage();
  }

  /** 功能：原位切换真实活动页，保持快照和页缝，单击缓存页无需先展开整篇；时间：2026-11-14；作者：lq。 */
  private showActivePage(): void {
    const frame = this.frames[this.active]; if (!frame || !this.live) return;
    this.live.classList.remove('mo-word-live-full');
    this.live.style.top = `${frame.top}px`; this.live.style.height = `${frame.heightPx}px`;
    this.wrap.style.top = `${-frame.top}px`;
    if (this.layer) this.layer.hidden = false;
    this.paintVisible();
  }

  /** 功能：按已缓存的页面坐标定位，不在滚动中查询正文几何；时间：2026-11-14；作者：lq。 */
  private pageAt(y: number): number {
    let low = 0, high = this.frames.length;
    while (low + 1 < high) { const mid = (low + high) >>> 1; if (this.frames[mid].top <= y) low = mid; else high = mid; }
    return low;
  }

  /** 功能：仅为视窗及相邻半屏制作页面；LRU 最多六页，滚动不触碰原编辑 DOM；时间：2026-11-14；作者：lq。 */
  private paintVisible(): void {
    if (!this.viewport || !this.layer || this.layer.hidden || !this.stage) return;
    const top = (this.viewport.scrollTop - this.origin - this.viewportHeight / 2) / this.zoom;
    const bottom = top + this.viewportHeight * 2 / this.zoom;
    const visible = new Set<number>();
    const start = this.pageAt(top);
    for (let index = start; index < this.frames.length && this.frames[index].top < bottom; index++) {
      if (index === this.active) continue;
      visible.add(index);
      let page = this.pages.get(index);
      if (!page) { page = this.paintPage(index); this.pages.set(index, page); this.layer.append(page); }
      page.hidden = false;
      this.pages.delete(index); this.pages.set(index, page);
    }
    for (const [index, page] of this.pages) page.hidden = !visible.has(index);
    for (const [index, page] of this.pages) if (this.pages.size > 6 && !visible.has(index)) { page.remove(); this.pages.delete(index); }
  }

  /** 功能：复用 GenOffice 的绝对定位块快照方式，原始表格及跨页装饰保持原样，快照不接受编辑；时间：2026-11-14；作者：lq。 */
  private paintPage(index: number): HTMLDivElement {
    const frame = this.frames[index], page = document.createElement('div');
    page.className = 'mo-word-cached-page'; page.dataset.cachedPage = String(index + 1);
    page.style.cssText = `left:${frame.left}px;top:${frame.top}px;width:${frame.widthPx}px;height:${frame.heightPx}px;`;
    const wrap = this.wrap.cloneNode(false) as HTMLElement; wrap.classList.add('mo-word-cached-wrap');
    wrap.style.cssText = `position:relative;top:0;margin:0;width:${frame.widthPx}px;height:${frame.heightPx}px;`;
    const body = this.paper.cloneNode(false) as HTMLElement;
    body.classList.add('mo-word-cached-body'); body.removeAttribute('id'); body.setAttribute('contenteditable', 'false'); body.setAttribute('aria-hidden', 'true');
    body.style.position = 'relative'; body.style.width = `${frame.widthPx}px`; body.style.height = `${frame.heightPx}px`; body.style.minHeight = '0'; body.style.padding = '0'; body.style.margin = '0'; body.style.boxShadow = 'none';
    const fragment = document.createDocumentFragment();
    for (const block of this.blocks) {
      if (block.paintBottom < frame.top || block.paintTop > frame.top + frame.heightPx) continue;
      const slot = document.createElement('div'); slot.className = 'mo-word-cached-block';
      slot.style.cssText = `left:${block.left - frame.left}px;top:${block.top - frame.top}px;width:${block.width}px;height:${block.height}px;`;
      const clone = block.element.cloneNode(true) as HTMLElement;
      clone.style.margin = '0'; clone.style.transform = 'none'; clone.style.position = 'relative'; clone.style.top = '0'; clone.style.left = '0'; clone.style.right = 'auto'; clone.style.bottom = 'auto';
      clone.removeAttribute('id'); clone.setAttribute('contenteditable', 'false');
      for (const element of clone.querySelectorAll('[id],[contenteditable],[tabindex]')) { element.removeAttribute('id'); element.removeAttribute('tabindex'); element.setAttribute('contenteditable', 'false'); }
      clone.querySelectorAll('.mo-no-print,.doc-table-handle,.doc-table-resize,.doc-image-resize-handle,.ProseMirror-gapcursor').forEach(element => element.remove());
      slot.append(clone); fragment.append(slot);
    }
    body.append(fragment); wrap.append(body); page.append(wrap);
    page.addEventListener('pointerdown', event => this.pick(index, event));
    page.addEventListener('click', event => { event.preventDefault(); event.stopPropagation(); });
    // 功能：缓存页的右键菜单仍由原节点和原选区处理，不对快照执行编辑命令；时间：2026-11-14；作者：lq。
    page.addEventListener('contextmenu', event => {
      event.preventDefault(); event.stopPropagation(); this.suspend();
      const target = document.elementFromPoint(event.clientX, event.clientY);
      if (!target || !this.wrap.contains(target)) { this.resume(); return; }
      const view = this.session.editor.view, hit = view.posAtCoords({ left: event.clientX, top: event.clientY });
      if (hit && view.state.selection.empty && this.paper.contains(target)) this.session.editor.commands.setTextSelection(hit.pos);
      target.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, button: 2, clientX: event.clientX, clientY: event.clientY }));
      this.resume();
    });
    return page;
  }

  /** 功能：点击快照先恢复原编辑层，再把输入交回真实节点；拖选期间始终使用原始全文坐标；时间：2026-11-14；作者：lq。 */
  private pick(index: number, event: PointerEvent): void {
    if (event.button !== 0) return;
    event.preventDefault(); event.stopPropagation(); this.active = index; this.gesture = true; this.showActivePage();
    const view = this.session.editor.view, target = document.elementFromPoint(event.clientX, event.clientY);
    if (target && this.wrap.contains(target)) this.pendingClick = { target, event };
    const hit = view.posAtCoords({ left: event.clientX, top: event.clientY });
    if (hit && target && this.paper.contains(target)) {
      this.session.editor.commands.setTextSelection(hit.pos); view.focus();
      this.drag = { anchor: hit.pos, x: event.clientX, y: event.clientY };
      target.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0, clientX: event.clientX, clientY: event.clientY, shiftKey: event.shiftKey, ctrlKey: event.ctrlKey, metaKey: event.metaKey }));
    }
  }

  /** 功能：缩短文档或卸载后移除分页显示容器，原编辑器恢复原层级且不残留空白高度；时间：2026-11-14；作者：lq。 */
  private reset(): void {
    this.suspend();
    if (this.stage) { this.stage.before(this.wrap); this.stage.remove(); }
    this.stage = undefined; this.live = undefined; this.layer = undefined;
    this.wrap.style.top = this.originalTop; this.pages.clear();
  }

  /** 功能：卸载恢复原层级与显示，关闭标签不保留快照、监听或计时任务；时间：2026-11-14；作者：lq。 */
  dispose(): void {
    if (this.disposed) return;
    this.suspend(); this.disposed = true;
    if (this.timer !== undefined) clearTimeout(this.timer);
    if (this.paintFrame !== undefined) cancelAnimationFrame(this.paintFrame);
    this.unsubscribe(); this.viewport?.removeEventListener('scroll', this.scroll);
    this.wrap.removeEventListener('pointerdown', this.pointerDown, true); this.wrap.removeEventListener('compositionstart', this.compositionStart, true); this.wrap.removeEventListener('compositionend', this.compositionEnd, true);
    document.removeEventListener('pointerup', this.pointerUp); document.removeEventListener('pointercancel', this.pointerUp); document.removeEventListener('pointermove', this.pointerMove);
    this.reset();
  }
}
