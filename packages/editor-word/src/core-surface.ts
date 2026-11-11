/** 功能：将定制核心挂到 Mini Office 提供的正文容器，管理异步打开与写盘确认；时间：2026-11-08；作者：lq。 */
import { WordCoreSession, type PreparedWordSave, type WordSessionState } from '../../word-core/src/index';
import { WordSurfaceLayout, type WordSurfaceLayoutOptions } from './core-layout';
// 功能：只在样式目录事务变化时更新 CSS，避免每次选区移动重新生成；时间：2026-11-10；作者：lq。
import { wordStyleCatalogOf } from '../../word-core/src/editing/styles';
// 功能：富文本页眉页脚使用稳定投影宿主，保留会话统一历史；时间：2026-11-11；作者：lq。
import { WordHeaderFooterRegions } from './core-header-footer';
import './core-header-footer.css';
// 功能：页边距拖动由当前分节和实际分页驱动；时间：2026-11-11；作者：lq。
import { WordPageMarginGuides } from './core-page-margins';

/** 功能：宿主继续负责文件选择、写盘和现有 UI，正文保持完整核心模型；时间：2026-11-08；作者：lq。 */
export interface WordCoreSurfaceOptions extends WordSurfaceLayoutOptions {
  readonly?: boolean;
  onState?: (state: WordSessionState) => void;
}

/** 功能：按容器持有一个核心会话，失败或取消时清理 DOM，不销毁其他标签的会话；时间：2026-11-08；作者：lq。 */
export class WordCoreSurface {
  readonly session: WordCoreSession;
  /** 功能：正文宿主持有真实分页协调器，供页面状态和验收读取；时间：2026-11-09；作者：lq。 */
  readonly layout: WordSurfaceLayout;
  layoutError?: unknown;
  private disposed = false;
  private saving = false;
  private readonly unsubscribe: () => void;
  private readonly regions?: WordHeaderFooterRegions;
  private readonly margins: WordPageMarginGuides;
  private constructor(
    readonly root: HTMLElement,
    private readonly content: HTMLElement,
    private readonly style: HTMLStyleElement,
    session: WordCoreSession,
    options: WordCoreSurfaceOptions,
    private readonly addedScope: boolean,
    private readonly ownsSession: boolean,
  ) {
    this.session = session;
    this.margins = new WordPageMarginGuides(session, session.editor.view.dom.closest<HTMLElement>('.page-wrap')!, error => options.onLayoutError?.(error));
    if (session.editor.isEditable) this.regions = new WordHeaderFooterRegions(session, session.editor.view.dom.closest<HTMLElement>('.page-wrap')!, error => options.onLayoutError?.(error));
    this.layout = new WordSurfaceLayout(session, { ...options, onLayout: state => {
      // 功能：后续真实重排成功才清除旧错误，避免界面永久保留过期提示；时间：2026-11-09；作者：lq。
      this.layoutError = undefined;
      this.regions?.refresh();
      this.margins.refresh(state.frames);
      options.onLayout?.(state);
    }, onLayoutError: error => {
      this.layoutError = error;
      options.onLayoutError?.(error);
    } });
    let catalog = wordStyleCatalogOf(session.editor.state.doc);
    this.unsubscribe = session.subscribe(state => {
      const next = wordStyleCatalogOf(session.editor.state.doc);
      if (next !== catalog) { catalog = next; style.textContent = session.styleCss(); }
      options.onState?.(state);
    });
    options.onState?.(session.state());
  }

  /** 功能：只挂载完整新核心；取消发生在字体/解析等待期间也释放迟到的编辑器；时间：2026-11-08；作者：lq。 */
  static async mount(root: HTMLElement, bytes: Uint8Array | undefined, options: WordCoreSurfaceOptions = {}, signal?: AbortSignal): Promise<WordCoreSurface> {
    return this.mountSession(root, element => bytes
      ? WordCoreSession.open(bytes, { element, readonly: options.readonly })
      : WordCoreSession.create({ element, readonly: options.readonly }), true, options, signal);
  }

  /** 功能：产品标签重新挂载复用工作区持有的唯一会话，隐藏时保留正文、选区和历史；时间：2026-11-09；作者：lq。 */
  static async attach(root: HTMLElement, session: WordCoreSession, options: WordCoreSurfaceOptions = {}, signal?: AbortSignal): Promise<WordCoreSurface> {
    return this.mountSession(root, async element => { element.append(session.editor.view.dom); return session; }, false, options, signal);
  }

  /** 功能：统一新建和复用的 DOM、字体与异步取消流程，仅释放本宿主拥有的会话；时间：2026-11-09；作者：lq。 */
  private static async mountSession(root: HTMLElement, open: (element: HTMLElement) => Promise<WordCoreSession>, ownsSession: boolean, options: WordCoreSurfaceOptions, signal?: AbortSignal): Promise<WordCoreSurface> {
    signal?.throwIfAborted();
    const addedScope = !root.classList.contains('mo-word-core');
    root.classList.add('mo-word-core');
    const content = document.createElement('div');
    // 功能：上游自然流样式限定在 editor-scroll 中，按同源容器层级挂载；时间：2026-11-09；作者：lq。
    content.className = 'editor-scroll';
    const page = document.createElement('div');
    page.className = 'page-wrap';
    // 功能：页眉页脚、注释及页界覆盖层相对本会话纸张定位；时间：2026-11-09；作者：lq。
    page.style.position = 'relative';
    content.append(page);
    const style = document.createElement('style');
    root.append(style, content);
    let session: WordCoreSession | undefined;
    try {
      session = await open(page);
      signal?.throwIfAborted();
      style.textContent = await session.activate();
      signal?.throwIfAborted();
      // 功能：页面尺寸从真实分节设置取得，不沿用旧编辑器页面或默认美式纸张；时间：2026-11-09；作者：lq。
      const settings = session.sections()[0]?.settings;
      if (!settings) throw new Error('Word 文档缺少有效分节设置');
      const paper = session.editor.view.dom;
      paper.style.boxSizing = 'border-box';
      paper.style.setProperty('--page-w', `${settings.pageWidth / 15}px`);
      paper.style.setProperty('--page-h', `${settings.pageHeight / 15}px`);
      paper.style.setProperty('--page-pad', `${settings.marginTop / 15}px ${settings.marginRight / 15}px ${settings.marginBottom / 15}px ${settings.marginLeft / 15}px`);
      return new WordCoreSurface(root, content, style, session, options, addedScope, ownsSession);
    } catch (error) {
      if (ownsSession) session?.dispose();
      content.remove();
      style.remove();
      if (addedScope) root.classList.remove('mo-word-core');
      throw error;
    }
  }

  /** 功能：写盘失败/取消保留脏状态，写盘期间的更新由核心版本号保留；时间：2026-11-08；作者：lq。 */
  async save(write: (bytes: Uint8Array) => Promise<void>): Promise<PreparedWordSave> {
    if (this.disposed) throw new Error('Word 正文容器已关闭');
    if (this.saving) throw new Error('Word 正在保存，请等待本次写盘完成');
    this.saving = true;
    try {
      const prepared = await this.session.prepareSave();
      // 功能：写盘回调只能拿到副本，避免外部修改核心签发的快照；时间：2026-11-08；作者：lq。
      await write(prepared.bytes.slice());
      if (!this.disposed) this.session.acknowledgeSave(prepared);
      return prepared;
    } finally { this.saving = false; }
  }

  /** 功能：标签卸载只移除本容器创建的节点，允许安全重复清理；时间：2026-11-08；作者：lq。 */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.layout.dispose();
    this.regions?.dispose();
    this.margins.dispose();
    this.unsubscribe();
    if (this.ownsSession) this.session.dispose();
    this.content.remove();
    this.style.remove();
    if (this.addedScope) this.root.classList.remove('mo-word-core');
  }
}
