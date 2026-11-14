/** 功能：在产品正文宿主应用同源分页结果，不回读 DOM 或创建第二份正文模型；时间：2026-11-09；作者：lq。
 * 上游显示接线参考：GenOffice e1f50e90d246f3ecc6427d2a0af90f04f37810ec 的 apps/docs/src/renderer/App.tsx；来源与许可见 packages/word-core/UPSTREAM.md。
 */
// 功能：目录页码按当前标题所在实际切片及分节编号格式显示；时间：2026-11-11；作者：lq。
import { wordHeadings, wordTocSnapshot } from '../../word-core/src/editing/toc';
import { pageOfY } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-slices';
import { formatPageNumber } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-hf';
import { WRAP_WIDTHS_COMMITTED_EVENT } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/cjk-punct-shrink';
import type { FieldDisplay } from '../../word-core/src/index';
import type { PmNode } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/convert';
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import { WordLayoutController, type WordLayoutResult } from '../../word-core/src/layout/controller';
import { sectionGeoms, effectiveTopPx, effectiveBottomPx, paperWidthPx, pageLeftPx, sectionWidthSpecs, columnLayoutSpecs, vAlignShiftSpecs, mirrorMarginSpecs, verticalTextSpecs, sectionGridPitchSpecs, sectionCharSpaceSpecs, sectionTopMarginSpecs, sectionPageStartSpecs, wrapWidthSig, newWidthPassState, resetWidthPassHistory, widthPassGate } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-sections';
import { wordLayoutContext, type WordLayoutContext } from './core-layout-context';
import { setColumnLayout } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/column-layout';
import { syncWrapObstacles } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/wrap-obstacles';
import { hfVariantReservedPx } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-hf';
import type { ColumnBlockPlacement } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-types';
import type { LineAnchor } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination';
import { planByElement, planPropsOf } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-plan';
import { buildPageCuts, applyGapLedger, pagePaperEdges, type PageCut } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-cuts';
import { createLineRectsCache, outerTableRows, singleCutCell, anchorElement, nextLineAnchor } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-lines';
// 功能：同一表格行的多个续行同步推到下一张纸，取代穿过内容的零高度虚线；时间：2026-11-11；作者：lq。
import { setWordCellGaps, alignWordTableCellText } from '../../word-core/src/layout/cell-gaps';
// 功能：无文字尾部以真实块末位置承接页缝，逐轮缓存行框避免重复测量长表格；时间：2026-11-11；作者：lq。
import { TopLevelPositions } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/top-level-pos';
import { resolveWordPageAnchor, wordTableRowRects } from './core-page-anchor';
// 功能：页缝与重复表头按实际合并网格生成，不额外扩出列；时间：2026-11-11；作者：lq。
import { wordTableGapSpans, wordRepeatedTableHeaders } from './core-table-gaps';
import { LayoutBatch, setPageGaps, setRowFills, rowFillsSig, gapLedgerOf, stampGapFlowTops, floatShiftsOf, floatCarriesOf, syncPhantomRowspans, alignTableGapFills, syncCutOverlays, setFloatVShifts, syncFloatShifts, setOversizeClips, clampCellBoxTops, clampCellImageTops, pinCellBoxesToPage, pinPageRelXPictures, pinParagraphRelVPictures, syncPageSheets, pageBorderStyleOf, syncPageBorders, canvasZoomOf, type PageGapSpec } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/pagination-gaps';
import { pageMargins } from '../../word-core/src/upstream/apps/docs/src/renderer/page-margins';
// 功能：字体和两端对齐的装饰会改变真实换行，分页订阅同源信号并按有界后续轮次合并；时间：2026-11-11；作者：lq。
import { bumpLineSampleFontEpoch } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-lines';
import { bumpHfProbeFontEpoch } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/hf-dom';
import { cjkPunctShrinkPluginKey } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/cjk-punct-shrink';
import { justifyShrinkPluginKey } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/justify-shrink';
// 功能：复用固定版本 GenOffice 的等高段落快速检查与首个脏块定位；时间：2026-11-13；作者：lq。
import { blocksKeepSlicing, singleInlineEditIndex, firstChangedTopLevelIndex } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/pagination-fast-path';
import { PHASED_APPEND } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/streaming-tail-guard';

/** 功能：分页摘要只暴露真实测量的页数与切点，不把无几何的测试环境伪装为完成排版；时间：2026-11-09；作者：lq。 */
/** 功能：显示、打印与 PDF 使用同一实际纸张原点及分节尺寸；时间：2026-11-11；作者：lq。 */
export interface WordPageFrame { top: number; left: number; widthPx: number; heightPx: number; section: number; }
export interface WordSurfaceLayoutState { pageCount: number; cuts: readonly PageCut[]; frames: readonly WordPageFrame[]; }
/** 功能：页面与字体尺寸变化触发排版，异常由产品宿主明确展示；时间：2026-11-09；作者：lq。 */
export interface WordSurfaceLayoutOptions {
  onLayout?: (state: WordSurfaceLayoutState) => void;
  onLayoutError?: (error: unknown) => void;
  /** 功能：实际测量前恢复完整显示，屏幕渲染优化不得改变分页输入；时间：2026-11-14；作者：lq。 */
  onBeforeMeasure?: () => void;
}

/** 功能：承接上游自然流测量及页间装饰，首轮落实单栏正文和普通表格分页；时间：2026-11-09；作者：lq。 */
export class WordSurfaceLayout {
  // 功能：宿主清理可能晚于会话销毁，保留监听节点供无 view 清理；时间：2026-11-11；作者：lq。
  private readonly editorDom: HTMLElement;
  private readonly controller: WordLayoutController;
  private readonly observer?: ResizeObserver;
  private readonly unsubscribe: () => void;
  private lastRevision: number;
  private disposed = false;
  private state?: WordSurfaceLayoutState;
  private rowFillSignature = '';
  private widthSignature = '';
  private context!: WordLayoutContext;
  private contextRevision = -1;
  // 功能：与上游相同维护已排版版本、可快速复用的编辑集合及脏块前沿；时间：2026-11-13；作者：lq。
  private lastResult?: WordLayoutResult;
  private lastDocument?: WordCoreSession['editor']['state']['doc'];
  private fastEdits: Set<number> | null = null;
  private dirtyFrom: number | null = null;
  readonly performanceStats = { fullPasses: 0, fastPasses: 0, resumedPasses: 0 };
  private readonly fontRefresh = () => {
    // 功能：字体加载后同时清除行框与页眉高度缓存，防止固定行高下沿用替代字体的旧换行；时间：2026-11-11；作者：lq。
    bumpLineSampleFontEpoch(); bumpHfProbeFontEpoch(); this.contextRevision = -1; this.refresh();
  };
  private readonly shrinkRefresh = ({ transaction }: { transaction: WordCoreSession['editor']['state']['tr'] }) => {
    if (transaction.getMeta(PHASED_APPEND)) return;
    if (transaction.docChanged) {
      const index = singleInlineEditIndex(transaction), dirty = firstChangedTopLevelIndex(transaction);
      if (index === null) this.fastEdits = null;
      else this.fastEdits?.add(index);
      this.dirtyFrom = dirty === null || this.dirtyFrom === null ? null : Math.min(this.dirtyFrom, dirty);
    }
    if (transaction.getMeta(cjkPunctShrinkPluginKey) || transaction.getMeta(justifyShrinkPluginKey)) {
      this.fastEdits = null; this.dirtyFrom = null;
      this.controller.followUp('text-wrap');
    }
  };
  // 功能：相同版面不重建整篇页眉页脚，避免每轮字体/宽度补偿重复创建 116 个区域；时间：2026-11-11；作者：lq。
  private regionsSignature = '';
  private readonly widthPass = newWidthPassState();
  private readonly refresh = () => {
    // 功能：字体、加载完成及外部尺寸触发必须完整重排，不能借用编辑前缀；时间：2026-11-13；作者：lq。
    this.fastEdits = null; this.dirtyFrom = null;
    // 功能：隐藏容器没有可测行盒，等尺寸恢复再排版，不捏造页数或反复轮询；时间：2026-11-09；作者：lq。
    if (!this.disposed && !this.session.editor.isDestroyed && this.editorDom.getBoundingClientRect().width > 0) this.controller.request();
  };

  constructor(private readonly session: WordCoreSession, private readonly options: WordSurfaceLayoutOptions = {}) {
    this.editorDom = session.editor.view.dom;
    this.lastRevision = session.state().revision;
    this.controller = new WordLayoutController(() => {
      // 功能：会话销毁与 React 清理之间已排队的重排也必须安全结束；时间：2026-11-11；作者：lq。
      if (session.editor.isDestroyed) return { root: this.editorDom, origin: 0, zoomFactor: 1, sections: [], composing: true };
      const root = session.editor.view.dom;
      // 功能：组合输入期间连字体探测及页面样式写入也跳过，DOM 继续归输入法管理；时间：2026-11-09；作者：lq。
      if (session.editor.view.composing) return { root, origin: 0, zoomFactor: this.factor(), sections: [], composing: true };
      // 功能：分批装载只测首屏和完整正文，尾部追加不会反复启动全篇测量；时间：2026-11-13；作者：lq。
      if (session.contentPending() && this.lastResult) return { root, origin: 0, zoomFactor: this.factor(), sections: [], composing: true };
      const sections = session.sections();
      const settings = sections[0]?.settings;
      if (!settings) throw new Error('Word 文档缺少页面设置');
      // 功能：宽度补偿等后续轮次复用本版本区域度量；正文/页脚/字体改变后才重建；时间：2026-11-11；作者：lq。
      if (this.contextRevision !== this.lastRevision) { this.context = wordLayoutContext(session, sections); this.contextRevision = this.lastRevision; }
      const factor = this.factor();
      // 功能：装饰层与居中纸张使用同一坐标原点，横纵切换后页眉和页脚不偏到容器左侧；时间：2026-11-10；作者：lq。
      const wrap = root.closest<HTMLElement>('.page-wrap');
      if (wrap) { wrap.style.width = `${paperWidthPx(sections)}px`; wrap.style.marginInline = 'auto'; }
      root.style.setProperty('--page-pad', `${this.context.top}px ${settings.marginRight / 15}px ${this.context.bottom}px ${settings.marginLeft / 15}px`);
      // 功能：方向或纸张改变后同步自然流纸张高度，不能沿用打开时的尺寸；时间：2026-11-10；作者：lq。
      root.style.setProperty('--page-h', `${settings.pageHeight / 15}px`);
      root.style.setProperty('--page-w', `${paperWidthPx(sections)}px`);
      // 功能：尺寸命令读取正文净宽，不能将含页边距的整张纸宽写入表格；时间：2026-11-11；作者：lq。
      root.style.setProperty('--section-content-w', `${(settings.pageWidth - settings.marginLeft - settings.marginRight) / 15}px`);
      const planOf = planByElement(session.editor.state.doc, session.editor.view, planPropsOf(session.parsed, sections, false), floatShiftsOf(session.editor.state));
      if (!session.editor.view.composing && syncWrapObstacles(root, factor, session.editor.view, planOf)) this.controller.followUp('obstacles');
      return {
        root, origin: root.getBoundingClientRect().top + this.context.top * factor,
        zoomFactor: factor, sections: sectionGeoms(sections, this.context.heights, session.evenAndOddHeaders()), sectionInfos: sections,
        headerFooterHeights: this.context.heights, evenOddHeaders: session.evenAndOddHeaders(), mirrorMargins: session.parsed.mirrorMargins,
        blockMetaOf: this.context.blockMetaOf, endnotes: this.context.endnotes,
        deletedSectionBreaks: new Set(session.parsed.blocks.filter(block => block.paraMarkDel && block.docxIndex !== null).map(block => block.docxIndex!)),
        composing: session.editor.view.composing,
        planOf,
        carryOf: floatCarriesOf(session.editor.state),
        previous: this.previousPrefix(),
      };
    }, result => this.apply(result), error => options.onLayoutError?.(error), run => {
      options.onBeforeMeasure?.();
      // 功能：同 GenOffice 在 100% 下度量，消除缩放取整导致的页数/文字锚点漂移；时间：2026-11-11；作者：lq。
      const zoom = this.editorDom.closest<HTMLElement>('.doc-zoom');
      const previous = zoom?.style.zoom;
      if (zoom && previous !== '1') zoom.style.zoom = '1';
      // 功能：测量隐藏页缝时同步还原源 rowspan；显示桥接只修改 DOM，不能被编辑器回读成正文；时间：2026-11-11；作者：lq。
      const observer = (session.editor.view as unknown as { domObserver?: { stop(): void; start(): void } }).domObserver;
      observer?.stop();
      this.editorDom.classList.add('measuring-natural');
      syncPhantomRowspans(this.editorDom);
      try { return run(); } finally {
        this.editorDom.classList.remove('measuring-natural');
        observer?.stop(); syncPhantomRowspans(this.editorDom); observer?.start();
        if (zoom && previous !== undefined && zoom.style.zoom !== previous) zoom.style.zoom = previous;
      }
    }, () => this.tryFastPass());
    this.unsubscribe = session.subscribe(state => {
      if (state.revision === this.lastRevision) return;
      this.lastRevision = state.revision;
      resetWidthPassHistory(this.widthPass);
      // 功能：事务细节随后由 transaction 监听补齐；此处不把普通编辑误标为外部失效；时间：2026-11-13；作者：lq。
      if (!this.disposed && this.editorDom.getBoundingClientRect().width > 0) this.controller.request();
    });
    session.editor.view.dom.addEventListener('compositionend', this.refresh);
    session.editor.view.dom.addEventListener('load', this.refresh, true);
    window.addEventListener('resize', this.refresh);
    document.fonts?.addEventListener('loadingdone', this.fontRefresh);
    session.editor.on('transaction', this.shrinkRefresh);
    if (typeof ResizeObserver !== 'undefined') {
      // 功能：只把可用宽度变化视为外部尺寸变化，避免装饰拉高页面后反复重排；时间：2026-11-09；作者：lq。
      let width = session.editor.view.dom.getBoundingClientRect().width;
      this.observer = new ResizeObserver(() => {
        if (session.editor.isDestroyed || this.disposed) return;
        const nextWidth = this.editorDom.getBoundingClientRect().width;
        if (Math.abs(nextWidth - width) < 0.1) return;
        width = nextWidth;
        this.refresh();
      });
      this.observer.observe(session.editor.view.dom);
    }
    this.refresh();
    if (session.contentPending()) void session.whenContentReady().then(() => {
      if (this.disposed) return;
      this.contextRevision = -1; this.lastResult = undefined; this.refresh();
    }).catch(error => { if (!this.disposed) options.onLayoutError?.(error); });
  }

  /** 功能：保存脏块的真实 DOM 身份，布局输入只向同源切片器交付安全的旧前缀；时间：2026-11-13；作者：lq。 */
  private previousPrefix() {
    const result = this.lastResult, dirty = this.dirtyFrom;
    if (!result || dirty === null || !Number.isFinite(dirty) || this.session.sections().some(section => !!section.settings.textDirection && section.settings.textDirection !== 'lrTb')) return;
    const doc = this.session.editor.state.doc;
    if (dirty >= doc.childCount) return;
    let pos = 0;
    for (let index = 0; index < dirty; index++) pos += doc.child(index).nodeSize;
    return { blocks: result.blocks, slices: result.outputs.preParity ?? result.slices, outputs: result.outputs, geometries: result.geometries,
      sectionSignature: result.effectiveSections?.map(section => `${section.firstBlockIndex}:${section.lastBlockIndex}:${section.startType ?? ''}`).join('|') ?? '',
      dirtyElement: this.session.editor.view.nodeDOM(pos) as Element | null };
  }

  /** 功能：逐一核对编辑段落高度，满足上游条件时完全跳过整篇测量与装饰重建；时间：2026-11-13；作者：lq。 */
  private tryFastPass(): WordLayoutResult | undefined {
    const result = this.lastResult, edited = this.fastEdits;
    if (!result || !edited?.size || this.disposed || this.session.editor.isDestroyed || this.session.editor.view.composing || this.session.contentPending()) return;
    if (this.session.sections().some(section => !!section.settings.textDirection && section.settings.textDirection !== 'lrTb')) return;
    const view = this.session.editor.view, doc = view.state.doc;
    const previous = new Map(result.blocks.map(block => [block.el, block]));
    const items = [];
    let index = 0, pos = 0;
    for (const target of [...edited].sort((a, b) => a - b)) {
      if (target >= doc.childCount) return;
      for (; index < target; index++) pos += doc.child(index).nodeSize;
      const element = view.nodeDOM(pos);
      if (!(element instanceof HTMLElement)) return;
      const block = previous.get(element);
      if (!block || element.classList.contains('doc-para-join-head') || element.classList.contains('doc-para-join-tail')) return;
      items.push({ el: element, prev: block });
    }
    if (!blocksKeepSlicing(items, result.slices, this.factor())) return;
    this.fastEdits = new Set(); this.dirtyFrom = Infinity;
    this.lastDocument = doc;
    this.performanceStats.fastPasses++;
    this.recordHeadingPages(result);
    return result;
  }

  /** 功能：等高编辑后重新绑定标题位置与原切片，保存时目录缓存对应当前版本；时间：2026-11-13；作者：lq。 */
  private recordHeadingPages(result: WordLayoutResult) {
    const sections = result.effectiveSections ?? this.session.sections(), cuts = this.state!.cuts;
    const headings = wordHeadings(this.session.editor.state.doc, this.session.formattingSource().styles);
    const byElement = new Map(result.blocks.filter(block => block.el).map(block => [block.el, block.top]));
    const pages = new Map<number, string>();
    for (const heading of headings) {
      const element = this.session.editor.view.nodeDOM(heading.pos) as HTMLElement | null;
      const top = element ? byElement.get(element) : undefined;
      const cut = top === undefined ? undefined : cuts[pageOfY(cuts, top)];
      if (cut) pages.set(heading.pos, formatPageNumber(cut.pageNo, sections[cut.section]?.pageNumberFmt));
    }
    if (!this.session.contentPending()) this.session.recordHeadingPages(pages);
    return { headings, pages };
  }

  /** 功能：供正文宿主和验收读取最后一次成功排版，不将装饰事务计为正文修改；时间：2026-11-09；作者：lq。 */
  current(): WordSurfaceLayoutState | undefined { return this.state; }
  /** 功能：验收和输出以完整正文版本及待执行补偿判断稳定，不能用暂时的首屏页数冒充完成；时间：2026-11-13；作者：lq。 */
  isSettled(): boolean { return !this.disposed && !this.session.editor.isDestroyed && !this.session.contentPending() && !!this.lastResult && this.lastDocument === this.session.editor.state.doc && !this.controller.hasPendingPass(); }
  /** 功能：同步重排使用实际浏览器几何，组合输入期间不修改输入法管理的 DOM；时间：2026-11-09；作者：lq。 */
  flush(): WordLayoutResult | undefined {
    if (this.disposed || this.session.editor.isDestroyed || this.editorDom.getBoundingClientRect().width <= 0) return;
    return this.controller.flush();
  }
  /** 功能：关闭宿主清理尺寸、字体和组合输入监听，防止迟到重排访问已销毁编辑器；时间：2026-11-09；作者：lq。 */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.controller.dispose();
    this.unsubscribe();
    this.observer?.disconnect();
    this.editorDom.removeEventListener('compositionend', this.refresh);
    this.editorDom.removeEventListener('load', this.refresh, true);
    window.removeEventListener('resize', this.refresh);
    document.fonts?.removeEventListener('loadingdone', this.fontRefresh);
    this.session.editor.off('transaction', this.shrinkRefresh);
  }

  /** 功能：沿用上游切点与 ProseMirror 装饰，页间空白不进入序列化和撤销历史；时间：2026-11-09；作者：lq。 */
  private apply(result: WordLayoutResult): void {
    this.performanceStats.fullPasses++;
    if (result.resumePage !== undefined) this.performanceStats.resumedPasses++;
    const sections = result.effectiveSections ?? this.session.sections();
    // 功能：复用本轮页眉/脚注度量，分节变体由同一来源列表解析；时间：2026-11-11；作者：lq。
    const context = this.context;
    const cuts = buildPageCuts(result.blocks, result.slices, result.geometries, context.heights, {
      titlePgOf: index => sections[index]?.titlePg,
      pageNumberFmtOf: index => sections[index]?.pageNumberFmt,
      floatSplits: result.outputs.floatSplits,
      evenOddHf: this.session.evenAndOddHeaders(),
    });
    const view = this.session.editor.view;
    const root = view.dom;
    const rect = root.getBoundingClientRect();
    const factor = this.factor();
    const paperWidth = paperWidthPx(sections);
    const pagesOfNotes = context.pageNotes(result.blocks, result.slices);
    const gaps: PageGapSpec[] = [];
    const cellGaps: Array<{ pos: number; height: number; page: number }> = [];
    const overlays: LineAnchor[] = [];
    const overlayPages: number[] = [];
    // 功能：切点的文本锚点也必须在自然流中解析，避免旧页间换行锁住下一轮位置；时间：2026-11-09；作者：lq。
    const hadNatural = root.classList.contains('measuring-natural');
    root.classList.add('measuring-natural');
    try {
      // 功能：同一轮自然流几何保持稳定，所有页缝共享行框与顶层事务位置；时间：2026-11-11；作者：lq。
      const rectsOf = createLineRectsCache();
      const positions = new TopLevelPositions(view);
      for (let index = 1; index < cuts.length; index++) {
        const cut = cuts[index];
        const previous = cuts[index - 1];
        const nextSettings = sections[cut.section].settings;
        const previousSettings = sections[previous.section].settings;
        const notesHeight = context.notesElement(pagesOfNotes[index - 1], previous.section)?.height ?? 0;
        const pad = Math.max(notesHeight, previous.padPx);
        const nextReserved = hfVariantReservedPx(context.heights[cut.section], cut.hf.variant);
        const previousReserved = hfVariantReservedPx(context.heights[previous.section], previous.hf.variant);
        const metrics = {
          marginTop: effectiveTopPx(nextSettings, nextReserved.headerPx),
          marginBottom: effectiveBottomPx(previousSettings, previousReserved.footerPx) + pad,
          marginLeft: nextSettings.marginLeft / 15,
          marginRight: nextSettings.marginRight / 15,
          bodyPad: pad,
          pageLeft: pageLeftPx(nextSettings, paperWidth), pageWidth: nextSettings.pageWidth / 15,
        };
        // 功能：压回切片允许的末行 leading 溢出与栏堆叠空区，保持物理纸张页界不逐页漂移；时间：2026-11-09；作者：lq。
        const pullUp = Math.max(0, previous.usedPx - previous.physUsedPx) + Math.max(0, previous.physUsedPx - previous.capacityPx);
        const base = { page: index, boundaryY: cut.start, metrics, ...(pullUp > 0.5 ? { pullUp } : {}) };
        const block = result.blocks[cut.lead.block];
        if (cut.lead.kind === 'tail') {
          gaps.push({ ...base, pos: view.state.doc.content.size, kind: 'inline' });
        } else if (cut.lead.kind === 'block' && block?.el) {
          gaps.push({ ...base, el: block.el, suppressLeadMt: cut.leadDropsSpaceBefore, leadMtPx: cut.leadMtPx });
        } else if (cut.lead.kind === 'row' && block?.el) {
          const row = outerTableRows(block.el)[cut.lead.row ?? -1];
          if (!row) throw new Error(`第 ${index + 1} 页的表格行定位失败`);
          const position = view.state.doc.resolve(view.posAtDOM(row, 0));
          let depth = position.depth;
          while (depth > 0 && position.node(depth).type.name !== 'docTableRow') depth--;
          if (!depth) throw new Error(`第 ${index + 1} 页的表格行不属于当前事务`);
          const cols = Math.max(1, ...outerTableRows(block.el).map(element => Array.from((element as HTMLTableRowElement).cells).reduce((sum, cell) => sum + cell.colSpan, 0)));
          // 功能：重复表头由装饰复制，不向真实文档插入行；时间：2026-11-09；作者：lq。
          const repeated = wordRepeatedTableHeaders(block.el, cut.repeatHeader?.height ?? 0, factor);
          const headers = repeated.rows, headerHeight = repeated.height;
          gaps.push({ ...base, pos: position.before(depth), kind: 'table', cols,
            tableCellSpans: wordTableGapSpans(block.el, cut.lead.row ?? 0),
            repeatHeaderEls: headers, repeatHeaderPx: headerHeight, repeatHeaderKey: headers.map(header => header.innerHTML).join('|') });
        } else if (block?.el) {
          // 功能：用 DOM 文本锚点定位跨页行，滚出视口后仍不依赖屏幕坐标命中；时间：2026-11-09；作者：lq。
          // 功能：跨页单元格按计划行选择主锚点，主/伴随页缝必须属于同一真实表格行；时间：2026-11-11；作者：lq。
          const rowRects = cut.lead.kind === 'cell' ? wordTableRowRects(block.el, cut.lead.row ?? -1, rectsOf) : rectsOf;
          const anchor = resolveWordPageAnchor(block.el, cut.lead.lineOffset ?? 0, block.height, factor, rowRects);
          // 功能：切点落在表格尾部边框/留白时，用块末真实事务边界；补偿已消耗高度，避免终止整份文档分页；时间：2026-11-11；作者：lq。
          if (anchor === 'after-block') {
            // 功能：计划行尾留白后还有真实表格行时，分页应落在下一行前，不能跳到整张表格末尾；时间：2026-11-11；作者：lq。
            const rows = cut.lead.kind === 'cell' ? outerTableRows(block.el) : [];
            const nextRow = rows[(cut.lead.row ?? -1) + 1];
            if (nextRow) {
              const position = view.state.doc.resolve(view.posAtDOM(nextRow, 0));
              let depth = position.depth;
              while (depth > 0 && position.node(depth).type.name !== 'docTableRow') depth--;
              if (!depth) throw new Error(`第 ${index + 1} 页的表格续行定位失败`);
              const cols = Math.max(1, ...rows.map(row => Array.from((row as HTMLTableRowElement).cells).reduce((sum, cell) => sum + cell.colSpan, 0)));
              gaps.push({ ...base, pos: position.before(depth), kind: 'table', cols, tableCellSpans: wordTableGapSpans(block.el, (cut.lead.row ?? -1) + 1) });
              continue;
            }
            const range = positions.of(block.el);
            if (!range) throw new Error(`第 ${index + 1} 页的正文块末定位失败`);
            const consumed = Math.max(0, block.height - (cut.lead.lineOffset ?? 0));
            // 功能：尾部行高只抵扣正文补白，不能吃掉物理页脚边距，否则末行会被下一页纸边截断；时间：2026-11-11；作者：lq。
            const footerMargin = effectiveBottomPx(previousSettings, previousReserved.footerPx);
            gaps.push({ ...base, pos: range.to, kind: 'inline', metrics: { ...metrics, marginBottom: Math.max(footerMargin, metrics.marginBottom - consumed) } });
            continue;
          }
          if (!anchor) throw new Error(`第 ${index + 1} 页缺少真实行锚点`);
          // 功能：末行整体移页时补足前页剩余正文高度，纸张高度与页脚位置保持完整；时间：2026-11-11；作者：lq。
          if (anchor.layoutOffset !== undefined) {
            const remainder = Math.max(0, (cut.lead.lineOffset ?? 0) - anchor.layoutOffset);
            metrics.marginBottom += remainder; metrics.bodyPad += remainder;
          }
          if (anchorElement(anchor)?.closest('.doc-nested-table')) { overlays.push(anchor); overlayPages.push(index); continue; }
          const pos = anchor.node instanceof Element
            ? view.posAtDOM(anchor.node.parentNode!, Array.prototype.indexOf.call(anchor.node.parentNode!.childNodes, anchor.node))
            : view.posAtDOM(anchor.node, anchor.charOffset);
          const box = block.el.getBoundingClientRect();
          // 功能：单格行插入真实页间装饰，多格或嵌套行沿用同源覆盖层，选区仍属于原事务；时间：2026-11-09；作者：lq。
          let cell = cut.lead.kind === 'cell' ? singleCutCell(outerTableRows(block.el)[cut.lead.row ?? -1], anchor) : undefined;
          if (cut.lead.kind === 'cell' && !cell) {
            const rows = outerTableRows(block.el) as HTMLTableRowElement[];
            const rowIndex = cut.lead.row ?? -1;
            const host = anchorElement(anchor)?.closest<HTMLTableCellElement>('td,th');
            const table = host?.closest('table');
            // 功能：只有映射到可编辑表格的真实单元格才建立同步续行，合并格也按实际覆盖行参与；时间：2026-11-11；作者：lq。
            if (host && table && !host.closest('.doc-nested-table')) {
              const cutTop = block.el.getBoundingClientRect().top + (cut.lead.lineOffset ?? 0) * factor;
              for (let ri = 0; ri <= rowIndex; ri++) for (const other of Array.from(rows[ri]?.cells ?? [])) {
                const span = Number(other.dataset.baseRowspan) || other.rowSpan;
                if (other === host || ri + span <= rowIndex) continue;
                const next = nextLineAnchor(other, (cutTop - other.getBoundingClientRect().top) / factor, factor, rectsOf);
                if (!next || anchorElement(next)?.closest('.doc-nested-table')) continue;
                const companionPos = next.node instanceof Element
                  ? view.posAtDOM(next.node.parentNode!, Array.prototype.indexOf.call(next.node.parentNode!.childNodes, next.node))
                  : view.posAtDOM(next.node, next.charOffset);
                cellGaps.push({ pos: companionPos, height: metrics.marginBottom + 28 + metrics.marginTop, page: index });
              }
              cell = host;
            }
          }
          const cellBox = cell ? (anchorElement(anchor)?.closest('td > *, th > *') ?? cell).getBoundingClientRect() : box;
          gaps.push({ ...base, pos, kind: cut.lead.kind === 'cell' ? cell ? 'cell' : 'cut' : 'inline', metrics: { ...metrics, marginLeft: (cellBox.left - rect.left) / factor, marginRight: (rect.right - cellBox.right) / factor } });
        } else if (block?.isEndnotes || block?.isFloatSpill) {
          gaps.push({ ...base, pos: view.state.doc.content.size, kind: 'inline' });
        } else throw new Error(`第 ${index + 1} 页缺少可定位的正文`);
      }
    } finally {
      // 功能：所有锚点解析结束后才恢复显示几何，提交装饰及纸面坐标读取使用最终 DOM；时间：2026-11-11；作者：lq。
      root.classList.remove('measuring-natural');
    }
    const batch = new LayoutBatch(view);
    setWordCellGaps(view, cellGaps, batch);
    // 功能：行高补偿按上游块索引和真实表格行关联，不能把算法记录直接当 DOM；时间：2026-11-09；作者：lq。
    const rowFills = (result.outputs.rowFills ?? []).flatMap(fill => {
      const block = result.blocks.find(block => block.tableRows && Math.abs(block.top - fill.blockTop) < 0.5);
      const el = block?.el ? outerTableRows(block.el)[fill.row] : undefined;
      return el ? [{ el, targetPx: fill.targetPx, extraPx: fill.extraPx }] : [];
    });
    setRowFills(view, rowFills, batch);
    // 功能：完整迁移上游按分节、栏、对称边距和竖排组合显示装饰；时间：2026-11-09；作者：lq。
    const columns = columnLayoutSpecs(result.blocks, result.slices, sections);
    const widthsChanged = widthPassGate(this.widthPass, columns);
    const specLists = [sectionGridPitchSpecs(result.blocks, sections), sectionCharSpaceSpecs(result.blocks, sections, this.session.parsed.balanceDbcsSpacing),
      sectionTopMarginSpecs(result.blocks, sections), sectionPageStartSpecs(result.blocks, sections),
      sectionWidthSpecs(result.blocks, sections, result.geometries, paperWidth),
      this.session.parsed.mirrorMargins ? mirrorMarginSpecs(result.blocks, result.slices, sections, sections[0].settings, cuts.map(cut => cut.pageNo)) : [],
      columns, vAlignShiftSpecs(result.blocks, result.slices, sections, result.geometries), verticalTextSpecs(result.blocks, result.slices, sections, result.geometries)];
    const specs = new Map<HTMLElement, ColumnBlockPlacement>();
    for (const spec of specLists.flat()) {
      const previous = specs.get(spec.el);
      specs.set(spec.el, previous ? { ...previous, ...spec, dx: previous.dx + spec.dx, dy: previous.dy + spec.dy } : spec);
    }
    setColumnLayout(view, [...specs.values()], batch);
    const floatShifts = [...(result.outputs.floatVShifts ?? []), ...(result.outputs.floatFlows ?? [])].flatMap(shift => {
      const block = result.blocks.find(block => block.el && Math.abs(block.top - shift.blockTop) < 0.5);
      return block?.el ? [{ el: block.el, dyPx: 'dyPx' in shift ? Number(shift.dyPx) : 0, flow: !('dyPx' in shift) }] : [];
    });
    setFloatVShifts(view, floatShifts, batch);
    setPageGaps(view, gaps, undefined, batch);
    batch.commit();
    alignTableGapFills(root, factor);
    // 功能：只修正背景覆盖范围，不改段落缩进和可编辑行盒；先集中读取再写入，避免每个页缝强制重排；时间：2026-11-11；作者：lq。
    const paperRect = root.getBoundingClientRect();
    const bands = Array.from(root.querySelectorAll<HTMLElement>('.page-gap:not(.page-gap-table)'));
    const bandBoxes = bands.map(band => band.getBoundingClientRect());
    bands.forEach((band, index) => {
      band.style.setProperty('--mo-gap-left', `${(paperRect.left - bandBoxes[index].left) / factor}px`);
      band.style.setProperty('--mo-gap-width', `${paperRect.width / factor}px`);
    });
    applyGapLedger(cuts, gapLedgerOf(gaps, overlayPages));
    // 功能：装饰提交后回填真实页界流坐标，吸收页首段落边距折叠的显示差异；时间：2026-11-09；作者：lq。
    stampGapFlowTops(root, cuts, root.getBoundingClientRect().top + context.top * factor, factor);
    const origin = root.getBoundingClientRect().top + context.top * factor;
    syncFloatShifts(root, result.floats, origin, factor, context.top - sections[0].settings.marginTop / 15, { cuts, carries: [] });
    clampCellBoxTops(root, root.getBoundingClientRect().top, factor);
    clampCellImageTops(root, factor);
    pinCellBoxesToPage(root, factor, cuts);
    pinPageRelXPictures(root, factor, cuts);
    pinParagraphRelVPictures(root, planByElement(view.state.doc, view, planPropsOf(this.session.parsed, sections, false)), new Set(gaps.flatMap(gap => 'el' in gap && gap.suppressLeadMt ? [gap.el] : [])));
    setOversizeClips(view, (result.outputs.oversizeClips ?? []).flatMap(clip => {
      const block = result.blocks.find(block => block.el && Math.abs(block.top - clip.blockTop) < 0.5);
      return block?.el ? [{ el: block.el, clipPx: clip.clipPx }] : [];
    }));
    const wrap = root.closest<HTMLElement>('.page-wrap')!;
    syncCutOverlays(wrap, overlays, factor);
    // 功能：沿用上游纸张边缘计算补齐末页，短正文的最后一页也保持完整纸张高度；时间：2026-11-09；作者：lq。
    const edges = pagePaperEdges(cuts, context.top);
    const lastEdge = edges[edges.length - 1];
    const lastSection = sections[cuts[cuts.length - 1]?.section ?? 0].settings;
    if (lastEdge) {
      const paperHeight = `${Math.round(lastEdge.top + lastSection.pageHeight / 15)}px`;
      root.style.minHeight = paperHeight;
      // 功能：多栏末页仍在 DOM 中堆叠，按上游规则固定纸张高度，防止留下不可见的尾部空白；时间：2026-11-09；作者：lq。
      const lastSlice = result.slices[result.slices.length - 1];
      if (lastSlice?.regions && (lastSlice.physHeight ?? 0) > 0) root.style.height = paperHeight;
      else root.style.removeProperty('height');
    } else { root.style.removeProperty('min-height'); root.style.removeProperty('height'); }
    const firstFrame = { left: pageLeftPx(sections[0].settings, paperWidth), width: sections[0].settings.pageWidth / 15 };
    const mixedPaper = sections.some(section => section.settings.pageWidth !== sections[0].settings.pageWidth);
    wrap.classList.toggle('paper-mixed', mixedPaper);
    // 功能：所有文档复用 GenOffice 的逐页底纸层，展开原生选区时也不退回带连续外框的整篇纸面；时间：2026-11-14；作者：lq。
    wrap.classList.add('mo-word-separate-sheets');
    syncPageSheets(wrap, factor, firstFrame, cuts);
    syncPageBorders(wrap, pageBorderStyleOf(sections[0].settings), factor, firstFrame, cuts);
    const paperTops = new Map(pagePaperEdges(cuts, context.top).map(edge => [edge.page, edge.top]));
    const frames: WordPageFrame[] = [];
    cuts.forEach((cut, index) => {
      const settings = sections[cut.section].settings;
      frames.push({ section: cut.section, widthPx: settings.pageWidth / 15, heightPx: settings.pageHeight / 15, left: pageLeftPx(settings, paperWidth), top: paperTops.get(index) ?? (index ? frames[index - 1].top + frames[index - 1].heightPx + 28 : 0) });
    });
    // 功能：浏览器的跨页合并格会把页缝也计入垂直居中；仅校正短文字的显示位置，保留源格式；时间：2026-11-11；作者：lq。
    const alignment = new LayoutBatch(view);
    alignWordTableCellText(view, frames.map((frame, index) => {
      const settings = sections[frame.section].settings, reserved = hfVariantReservedPx(context.heights[frame.section], cuts[index].hf.variant);
      return { top: frame.top + effectiveTopPx(settings, reserved.headerPx), bottom: frame.top + frame.heightPx - effectiveBottomPx(settings, reserved.footerPx) };
    }), factor, alignment);
    alignment.commit();
    this.paintPageRegions(wrap, cuts, context, pagesOfNotes, result, frames);
    const widthSignature = wrapWidthSig([...specs.values()]);
    if (widthSignature !== this.widthSignature || widthsChanged) {
      this.widthSignature = widthSignature;
      this.controller.followUp('width');
    }
    const fillSignature = rowFillsSig(result.outputs.rowFills ?? []);
    if (fillSignature !== this.rowFillSignature) {
      this.rowFillSignature = fillSignature;
      this.controller.followUp('rowFills');
    }
    this.state = { pageCount: cuts.length, cuts, frames };
    if (!this.session.contentPending()) this.session.recordPageCount(cuts.length);
    // 功能：按元素身份映射自然流位置，缺少测量时不伪造页数；只刷新 DOM/快照缓存；时间：2026-11-11；作者：lq。
    const { headings, pages } = this.recordHeadingPages(result);
    // 功能：布局只生成顶层属性摘要，避免每轮分页复制整份文字/表格 JSON；时间：2026-11-11；作者：lq。
    const content: PmNode[] = [];
    this.session.editor.state.doc.forEach(node => content.push({ type: node.type.name, attrs: node.attrs }));
    const snapshot = wordTocSnapshot({ type: 'doc', content }, headings, pages);
    let contentIndex = 0;
    let tocPagesChanged = false;
    // 功能：文档 dirty 在本轮固定，不对数千个顶层段落重复执行撤销/重做能力检查；时间：2026-11-11；作者：lq。
    const dirty = this.session.state().dirty;
    this.session.editor.state.doc.forEach((node, pos) => {
      const display = snapshot.content?.[contentIndex++]?.attrs?.fieldDisplay as FieldDisplay | undefined;
      if (display?.kind !== 'tocLine' || !display.right) return;
      const element = view.nodeDOM(pos) as HTMLElement | null;
      const page = element?.querySelector('.doc-toc-page');
      // 功能：打开未修改文档时保留已存页码；空缓存与新目录继续跟随布局；时间：2026-11-11；作者：lq。
      if (page && (dirty || !String(node.attrs.fieldDisplay?.right ?? '').trim()) && page.textContent !== String(display.right)) { page.textContent = String(display.right); tocPagesChanged = true; }
    });
    // 功能：页码从空缓存变为文字后重新度量引导点，避免尾部宽度按旧值把页码挤到下一行；时间：2026-11-11；作者：lq。
    if (tocPagesChanged) {
      document.dispatchEvent(new Event(WRAP_WIDTHS_COMMITTED_EVENT));
      this.controller.followUp('toc-pages');
    }
    this.lastResult = result; this.fastEdits = new Set(); this.dirtyFrom = Infinity;
    this.lastDocument = view.state.doc;
    this.options.onLayout?.(this.state);
  }

  /** 功能：缩放只改变显示与测量单位，不改 DOCX 纸张参数；时间：2026-11-09；作者：lq。 */
  private factor(): number { return canvasZoomOf(this.session.editor.view.dom.closest<HTMLElement>('.doc-zoom'), 1); }

  /** 功能：页眉页脚与注释按实际页界显示，页面设置变化与撤销时同步重绘；时间：2026-11-10；作者：lq。 */
  private paintPageRegions(wrap: HTMLElement, cuts: PageCut[], context: WordLayoutContext, notes: ReturnType<WordLayoutContext['pageNotes']>, result: WordLayoutResult, frames: readonly WordPageFrame[]): void {
    const signature = JSON.stringify([this.lastRevision, frames, cuts.map(cut => [cut.hf, cut.display]), notes.map(items => items.map(item => [item.id, item.height])), result.endnotesTop]);
    if (signature === this.regionsSignature && wrap.querySelector(':scope > .mo-core-page-regions')) return;
    this.regionsSignature = signature;
    const sectionPages = new Map<number, number>();
    for (const cut of cuts) sectionPages.set(cut.section, (sectionPages.get(cut.section) ?? 0) + 1);
    let layer = wrap.querySelector<HTMLElement>(':scope > .mo-core-page-regions');
    if (!layer) { layer = document.createElement('div'); layer.className = 'mo-core-page-regions'; layer.contentEditable = 'false'; layer.style.cssText = 'position:absolute;inset:0;pointer-events:none'; wrap.append(layer); }
    // 功能：在离线片段中构建整批区域后一次挂载；WebKit 逐页插入会重复更新整篇布局和辅助功能树；时间：2026-11-11；作者：lq。
    const fragment = document.createDocumentFragment();
    cuts.forEach((cut, index) => {
      const settings = context.sections[cut.section].settings;
      const { top, left } = frames[index];
      const margins = pageMargins(settings, cut.pageNo, !!this.session.parsed.mirrorMargins);
      const reserved = hfVariantReservedPx(context.heights[cut.section], cut.hf.variant);
      for (const kind of ['header', 'footer'] as const) {
        const strip = context.strip(cut.section, kind, cut.hf.variant, cut.display, cuts.length, sectionPages.get(cut.section) ?? 0);
        // 功能：空白区域也提供双击入口；只读报告仍保留原生显示层；时间：2026-11-11；作者：lq。
        const region = strip ?? (this.session.editor.isEditable ? document.createElement('div') : undefined);
        if (!region) continue;
        region.classList.add(strip ? `mo-core-${kind}` : `mo-core-empty-${kind}`);
        region.dataset.hfSection = String(cut.section); region.dataset.hfKind = kind; region.dataset.hfVariant = cut.hf.variant;
        region.dataset.hfPage = String(index); region.dataset.hfNumber = cut.display;
        region.dataset.hfTotal = String(cuts.length); region.dataset.hfSectionPages = String(sectionPages.get(cut.section) ?? 0);
        // 功能：宿主已经按纸张左上角定位，清除同源居中条带的半宽位移；时间：2026-11-11；作者：lq。
        // 功能：同 GenOffice 以页脚条带自身下沿锚定距纸底的距离，不能减去正文预留高度代替真实条带高度；时间：2026-11-11；作者：lq。
        region.style.position = 'absolute'; region.style.bottom = 'auto'; region.style.transform = kind === 'footer' ? 'translateY(-100%)' : 'none';
        region.style.left = `${left + margins.left / 15}px`;
        region.style.width = `${(settings.pageWidth - margins.left - margins.right) / 15}px`;
        region.style.top = `${top + (kind === 'header' ? (settings.headerDist ?? 720) / 15 : settings.pageHeight / 15 - (settings.footerDist ?? 720) / 15)}px`;
        if (this.session.editor.isEditable) { region.style.pointerEvents = 'auto'; region.style.minHeight = '20px'; region.title = `双击编辑${kind === 'header' ? '页眉' : '页脚'}`; }
        fragment.append(region);
      }
      const footnotes = context.notesElement(notes[index] ?? [], cut.section);
      if (footnotes) {
        footnotes.element.style.top = `${top + settings.pageHeight / 15 - effectiveBottomPx(settings, reserved.footerPx) - footnotes.height}px`;
        footnotes.element.style.left = `${left + margins.left / 15}px`;
        fragment.append(footnotes.element);
      }
    });
    if (context.endnotes.length && result.endnotesTop !== undefined) {
      const notes = context.notesElement(context.endnotes, context.sections.length - 1, 'endnote')!;
      notes.element.classList.add('mo-core-endnotes');
      const cut = [...cuts].reverse().find(cut => cut.start <= result.endnotesTop!) ?? cuts[0];
      notes.element.style.top = `${context.top + result.endnotesTop + cut.flowOffsetPx}px`;
      fragment.append(notes.element);
    }
    layer.replaceChildren(fragment);
  }
}
