/** 功能：页眉页脚双击入口和稳定编辑宿主，分页重绘只移动投影而不打断中文输入；时间：2026-11-11；作者：lq。 */
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import { WordHeaderFooterEditor, type WordHeaderFooterTarget } from '../../word-core/src/editing/header-footer';
import { insertHfField, setHfFieldNumbers, type HfFieldKind } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/hf-editor';

/** 功能：区域内容即时进入主会话；退出、标签卸载和保存均不依赖额外提交；时间：2026-11-11；作者：lq。 */
export class WordHeaderFooterRegions {
  private active?: { region: WordHeaderFooterEditor; page: string; host: HTMLElement; field: HTMLElement; wrapper: HTMLElement };
  private readonly doubleClick = (event: MouseEvent) => {
    const element = (event.target as Element).closest<HTMLElement>('[data-hf-section]');
    if (!element || !this.wrap.contains(element) || !this.core.editor.isEditable) return;
    event.preventDefault(); event.stopPropagation(); this.open(element);
  };
  private readonly keyDown = (event: KeyboardEvent) => { if (event.key === 'Escape' && !event.isComposing && this.active && !this.active.region.editor.view.composing) { event.preventDefault(); this.close(); this.core.editor.view.focus(); } };
  private readonly focusMain = () => { if (this.active && !this.active.region.editor.view.composing) this.close(); };
  constructor(private readonly core: WordCoreSession, private readonly wrap: HTMLElement, private readonly onError: (error: unknown) => void) {
    wrap.addEventListener('dblclick', this.doubleClick);
    wrap.addEventListener('keydown', this.keyDown);
    core.editor.view.dom.addEventListener('focusin', this.focusMain);
  }
  /** 功能：新页界完成后更新宿主位置和动态域，活动子编辑器始终保持连接；时间：2026-11-11；作者：lq。 */
  refresh(): void {
    const active = this.active;
    if (!active) return;
    const target = active.region.target;
    const element = this.wrap.querySelector<HTMLElement>(`.mo-core-page-regions [data-hf-page="${active.page}"][data-hf-kind="${target.kind}"][data-hf-section="${target.section}"][data-hf-variant="${target.variant}"]`);
    if (!element) { if (!active.region.editor.view.composing) this.close(); return; }
    active.wrapper.style.top = element.style.top; active.wrapper.style.left = element.style.left; active.wrapper.style.width = element.style.width;
    element.style.visibility = 'hidden';
    setHfFieldNumbers(active.region.editor, element.dataset.hfNumber ?? '1', element.dataset.hfTotal ?? '1', element.dataset.hfSectionPages ?? '1');
  }
  /** 功能：在纸张原区域显示富编辑器和必要的域工具，其他工具继续复用现有顶部工具栏；时间：2026-11-11；作者：lq。 */
  private open(element: HTMLElement): void {
    if (this.active?.region.editor.view.composing) return;
    this.close();
    const wrapper = document.createElement('div'); wrapper.className = 'mo-core-hf-editing';
    const field = document.createElement('div'); field.className = 'mo-core-hf-controls'; field.setAttribute('role', 'toolbar'); field.setAttribute('aria-label', '页眉页脚区域工具');
    const label = document.createElement('span'); label.textContent = `${element.dataset.hfKind === 'header' ? '页眉' : '页脚'} · 第 ${Number(element.dataset.hfSection) + 1} 节`; field.append(label);
    const host = document.createElement('div'); wrapper.append(field, host); this.wrap.append(wrapper);
    const target: WordHeaderFooterTarget = { section: Number(element.dataset.hfSection), kind: element.dataset.hfKind as 'header' | 'footer', variant: element.dataset.hfVariant as WordHeaderFooterTarget['variant'] };
    // 功能：标尺读取区域分节身份，不借用正文旧选区的纸张设置；时间：2026-11-11；作者：lq。
    wrapper.dataset.wordRegionSection = String(target.section);
    const region = new WordHeaderFooterEditor(this.core, host, target, this.onError);
    for (const [kind, name] of [['PAGE', '页码'], ['NUMPAGES', '总页数'], ['SECTIONPAGES', '本节页数']] as const) {
      const button = document.createElement('button'); button.type = 'button'; button.textContent = name; button.setAttribute('aria-label', `区域插入${name}`);
      button.addEventListener('mousedown', event => event.preventDefault());
      button.addEventListener('click', () => { if (!region.editor.view.composing) insertHfField(region.editor, kind as HfFieldKind); }); field.append(button);
    }
    const close = document.createElement('button'); close.type = 'button'; close.textContent = '关闭'; close.setAttribute('aria-label', '关闭页眉页脚编辑');
    close.addEventListener('mousedown', event => event.preventDefault()); close.addEventListener('click', () => { if (!region.editor.view.composing) { this.close(); this.core.editor.view.focus(); } }); field.append(close);
    this.active = { region, page: element.dataset.hfPage!, host, field, wrapper };
    this.refresh(); region.editor.commands.focus('end');
  }
  /** 功能：去掉显示投影时恢复原条带，主会话内容及历史不受影响；时间：2026-11-11；作者：lq。 */
  private close(): void {
    const active = this.active; this.active = undefined;
    if (!active) return;
    active.region.dispose(); active.wrapper.remove();
    this.wrap.querySelectorAll<HTMLElement>('.mo-core-page-regions [data-hf-section]').forEach(element => { element.style.visibility = ''; });
  }
  /** 功能：宿主卸载释放监听和区域注册，不销毁工作区唯一会话；时间：2026-11-11；作者：lq。 */
  dispose(): void {
    this.close(); this.wrap.removeEventListener('dblclick', this.doubleClick); this.wrap.removeEventListener('keydown', this.keyDown);
    this.core.editor.view.dom.removeEventListener('focusin', this.focusMain);
  }
}
