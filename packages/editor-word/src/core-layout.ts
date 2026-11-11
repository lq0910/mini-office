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
import { lineStartAnchor, nextLineAnchor, outerTableRows, singleCutCell, anchorElement } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-lines';
import { LayoutBatch, setPageGaps, setRowFills, rowFillsSig, gapLedgerOf, stampGapFlowTops, floatShiftsOf, floatCarriesOf, syncPhantomRowspans, alignTableGapFills, syncCutOverlays, setFloatVShifts, syncFloatShifts, setOversizeClips, clampCellBoxTops, clampCellImageTops, pinCellBoxesToPage, pinPageRelXPictures, pinParagraphRelVPictures, syncPageSheets, pageBorderStyleOf, syncPageBorders, canvasZoomOf, type PageGapSpec } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/pagination-gaps';
import { pageMargins } from '../../word-core/src/upstream/apps/docs/src/renderer/page-margins';

/** 功能：分页摘要只暴露真实测量的页数与切点，不把无几何的测试环境伪装为完成排版；时间：2026-11-09；作者：lq。 */
/** 功能：显示、打印与 PDF 使用同一实际纸张原点及分节尺寸；时间：2026-11-11；作者：lq。 */
export interface WordPageFrame { top: number; left: number; widthPx: number; heightPx: number; section: number; }
export interface WordSurfaceLayoutState { pageCount: number; cuts: readonly PageCut[]; frames: readonly WordPageFrame[]; }
/** 功能：页面与字体尺寸变化触发排版，异常由产品宿主明确展示；时间：2026-11-09；作者：lq。 */
export interface WordSurfaceLayoutOptions {
  onLayout?: (state: WordSurfaceLayoutState) => void;
  onLayoutError?: (error: unknown) => void;
}

/** 功能：承接上游自然流测量及页间装饰，首轮落实单栏正文和普通表格分页；时间：2026-11-09；作者：lq。 */
export class WordSurfaceLayout {
  private readonly controller: WordLayoutController;
  private readonly observer?: ResizeObserver;
  private readonly unsubscribe: () => void;
  private lastRevision: number;
  private disposed = false;
  private state?: WordSurfaceLayoutState;
  private rowFillSignature = '';
  private widthSignature = '';
  private context!: WordLayoutContext;
  private readonly widthPass = newWidthPassState();
  private readonly refresh = () => {
    // 功能：隐藏容器没有可测行盒，等尺寸恢复再排版，不捏造页数或反复轮询；时间：2026-11-09；作者：lq。
    if (!this.disposed && this.session.editor.view.dom.getBoundingClientRect().width > 0) this.controller.request();
  };

  constructor(private readonly session: WordCoreSession, private readonly options: WordSurfaceLayoutOptions = {}) {
    this.lastRevision = session.state().revision;
    this.controller = new WordLayoutController(() => {
      const root = session.editor.view.dom;
      // 功能：组合输入期间连字体探测及页面样式写入也跳过，DOM 继续归输入法管理；时间：2026-11-09；作者：lq。
      if (session.editor.view.composing) return { root, origin: 0, zoomFactor: this.factor(), sections: [], composing: true };
      const sections = session.sections();
      const settings = sections[0]?.settings;
      if (!settings) throw new Error('Word 文档缺少页面设置');
      this.context = wordLayoutContext(session, sections);
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
      };
    }, result => this.apply(result), error => options.onLayoutError?.(error));
    this.unsubscribe = session.subscribe(state => {
      if (state.revision === this.lastRevision) return;
      this.lastRevision = state.revision;
      resetWidthPassHistory(this.widthPass);
      this.refresh();
    });
    session.editor.view.dom.addEventListener('compositionend', this.refresh);
    session.editor.view.dom.addEventListener('load', this.refresh, true);
    window.addEventListener('resize', this.refresh);
    document.fonts?.addEventListener('loadingdone', this.refresh);
    if (typeof ResizeObserver !== 'undefined') {
      // 功能：只把可用宽度变化视为外部尺寸变化，避免装饰拉高页面后反复重排；时间：2026-11-09；作者：lq。
      let width = session.editor.view.dom.getBoundingClientRect().width;
      this.observer = new ResizeObserver(() => {
        const nextWidth = session.editor.view.dom.getBoundingClientRect().width;
        if (Math.abs(nextWidth - width) < 0.1) return;
        width = nextWidth;
        this.refresh();
      });
      this.observer.observe(session.editor.view.dom);
    }
    this.refresh();
  }

  /** 功能：供正文宿主和验收读取最后一次成功排版，不将装饰事务计为正文修改；时间：2026-11-09；作者：lq。 */
  current(): WordSurfaceLayoutState | undefined { return this.state; }
  /** 功能：同步重排使用实际浏览器几何，组合输入期间不修改输入法管理的 DOM；时间：2026-11-09；作者：lq。 */
  flush(): WordLayoutResult | undefined {
    if (this.disposed || this.session.editor.view.dom.getBoundingClientRect().width <= 0) return;
    return this.controller.flush();
  }
  /** 功能：关闭宿主清理尺寸、字体和组合输入监听，防止迟到重排访问已销毁编辑器；时间：2026-11-09；作者：lq。 */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.controller.dispose();
    this.unsubscribe();
    this.observer?.disconnect();
    this.session.editor.view.dom.removeEventListener('compositionend', this.refresh);
    this.session.editor.view.dom.removeEventListener('load', this.refresh, true);
    window.removeEventListener('resize', this.refresh);
    document.fonts?.removeEventListener('loadingdone', this.refresh);
  }

  /** 功能：沿用上游切点与 ProseMirror 装饰，页间空白不进入序列化和撤销历史；时间：2026-11-09；作者：lq。 */
  private apply(result: WordLayoutResult): void {
    const sections = result.effectiveSections ?? this.session.sections();
    const context = wordLayoutContext(this.session, sections);
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
    const overlays: LineAnchor[] = [];
    const overlayPages: number[] = [];
    // 功能：切点的文本锚点也必须在自然流中解析，避免旧页间换行锁住下一轮位置；时间：2026-11-09；作者：lq。
    const hadNatural = root.classList.contains('measuring-natural');
    root.classList.add('measuring-natural');
    try {
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
          const headers: HTMLElement[] = [];
          let headerHeight = 0;
          for (const source of outerTableRows(block.el)) {
            if (!cut.repeatHeader || headerHeight >= cut.repeatHeader.height - 1.5) break;
            const clone = source.cloneNode(true) as HTMLElement;
            clone.classList.add('page-gap-inline', 'page-repeat-header');
            clone.contentEditable = 'false';
            headers.push(clone);
          headerHeight += source.getBoundingClientRect().height / factor;
          }
          gaps.push({ ...base, pos: position.before(depth), kind: 'table', cols,
            repeatHeaderEls: headers, repeatHeaderPx: headerHeight, repeatHeaderKey: headers.map(header => header.innerHTML).join('|') });
        } else if (block?.el) {
          // 功能：用 DOM 文本锚点定位跨页行，滚出视口后仍不依赖屏幕坐标命中；时间：2026-11-09；作者：lq。
          const anchor = lineStartAnchor(block.el, cut.lead.lineOffset ?? 0, factor) ?? nextLineAnchor(block.el, cut.lead.lineOffset ?? 0, factor);
          if (!anchor) throw new Error(`第 ${index + 1} 页缺少真实行锚点`);
          if (anchorElement(anchor)?.closest('.doc-nested-table')) { overlays.push(anchor); overlayPages.push(index); continue; }
          const pos = anchor.node instanceof Element
            ? view.posAtDOM(anchor.node.parentNode!, Array.prototype.indexOf.call(anchor.node.parentNode!.childNodes, anchor.node))
            : view.posAtDOM(anchor.node, anchor.charOffset);
          const box = block.el.getBoundingClientRect();
          // 功能：单格行插入真实页间装饰，多格或嵌套行沿用同源覆盖层，选区仍属于原事务；时间：2026-11-09；作者：lq。
          const cell = cut.lead.kind === 'cell' ? singleCutCell(outerTableRows(block.el)[cut.lead.row ?? -1], anchor) : undefined;
          const cellBox = cell ? (anchorElement(anchor)?.closest('td > *, th > *') ?? cell).getBoundingClientRect() : box;
          gaps.push({ ...base, pos, kind: cut.lead.kind === 'cell' ? cell ? 'cell' : 'cut' : 'inline', metrics: { ...metrics, marginLeft: (cellBox.left - rect.left) / factor, marginRight: (rect.right - cellBox.right) / factor } });
        } else if (block?.isEndnotes || block?.isFloatSpill) {
          gaps.push({ ...base, pos: view.state.doc.content.size, kind: 'inline' });
        } else throw new Error(`第 ${index + 1} 页缺少可定位的正文`);
      }
    } finally {
      if (!hadNatural) root.classList.remove('measuring-natural');
    }
    const batch = new LayoutBatch(view);
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
    syncPhantomRowspans(root);
    alignTableGapFills(root, factor);
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
    syncPageSheets(wrap, factor, mixedPaper ? firstFrame : null, cuts);
    syncPageBorders(wrap, pageBorderStyleOf(sections[0].settings), factor, firstFrame, cuts);
    const paperTops = new Map(pagePaperEdges(cuts, context.top).map(edge => [edge.page, edge.top]));
    const frames: WordPageFrame[] = [];
    cuts.forEach((cut, index) => {
      const settings = sections[cut.section].settings;
      frames.push({ section: cut.section, widthPx: settings.pageWidth / 15, heightPx: settings.pageHeight / 15, left: pageLeftPx(settings, paperWidth), top: paperTops.get(index) ?? (index ? frames[index - 1].top + frames[index - 1].heightPx + 28 : 0) });
    });
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
    this.session.recordPageCount(cuts.length);
    // 功能：按元素身份映射自然流位置，缺少测量时不伪造页数；只刷新 DOM/快照缓存；时间：2026-11-11；作者：lq。
    const headings = wordHeadings(this.session.editor.state.doc, this.session.formattingSource().styles);
    const byElement = new Map(result.blocks.filter(block => block.el).map(block => [block.el, block.top]));
    const pages = new Map<number, string>();
    for (const heading of headings) {
      const element = view.nodeDOM(heading.pos) as HTMLElement | null;
      const top = element ? byElement.get(element) : undefined;
      const cut = top === undefined ? undefined : cuts[pageOfY(cuts, top)];
      if (cut) pages.set(heading.pos, formatPageNumber(cut.pageNo, sections[cut.section]?.pageNumberFmt));
    }
    this.session.recordHeadingPages(pages);
    // 功能：布局只生成顶层属性摘要，避免每轮分页复制整份文字/表格 JSON；时间：2026-11-11；作者：lq。
    const content: PmNode[] = [];
    this.session.editor.state.doc.forEach(node => content.push({ type: node.type.name, attrs: node.attrs }));
    const snapshot = wordTocSnapshot({ type: 'doc', content }, headings, pages);
    let contentIndex = 0;
    let tocPagesChanged = false;
    this.session.editor.state.doc.forEach((node, pos) => {
      const display = snapshot.content?.[contentIndex++]?.attrs?.fieldDisplay as FieldDisplay | undefined;
      if (display?.kind !== 'tocLine' || !display.right) return;
      const element = view.nodeDOM(pos) as HTMLElement | null;
      const page = element?.querySelector('.doc-toc-page');
      // 功能：打开未修改文档时保留已存页码；空缓存与新目录继续跟随布局；时间：2026-11-11；作者：lq。
      if (page && (this.session.state().dirty || !String(node.attrs.fieldDisplay?.right ?? '').trim()) && page.textContent !== String(display.right)) { page.textContent = String(display.right); tocPagesChanged = true; }
    });
    // 功能：页码从空缓存变为文字后重新度量引导点，避免尾部宽度按旧值把页码挤到下一行；时间：2026-11-11；作者：lq。
    if (tocPagesChanged) {
      document.dispatchEvent(new Event(WRAP_WIDTHS_COMMITTED_EVENT));
      this.controller.followUp('toc-pages');
    }
    this.options.onLayout?.(this.state);
  }

  /** 功能：缩放只改变显示与测量单位，不改 DOCX 纸张参数；时间：2026-11-09；作者：lq。 */
  private factor(): number { return canvasZoomOf(this.session.editor.view.dom.closest<HTMLElement>('.doc-zoom'), 1); }

  /** 功能：页眉页脚与注释按实际页界显示，页面设置变化与撤销时同步重绘；时间：2026-11-10；作者：lq。 */
  private paintPageRegions(wrap: HTMLElement, cuts: PageCut[], context: WordLayoutContext, notes: ReturnType<WordLayoutContext['pageNotes']>, result: WordLayoutResult, frames: readonly WordPageFrame[]): void {
    let layer = wrap.querySelector<HTMLElement>(':scope > .mo-core-page-regions');
    if (!layer) { layer = document.createElement('div'); layer.className = 'mo-core-page-regions'; layer.contentEditable = 'false'; layer.style.cssText = 'position:absolute;inset:0;pointer-events:none'; wrap.append(layer); }
    layer.replaceChildren();
    cuts.forEach((cut, index) => {
      const settings = context.sections[cut.section].settings;
      const { top, left } = frames[index];
      const margins = pageMargins(settings, cut.pageNo, !!this.session.parsed.mirrorMargins);
      const reserved = hfVariantReservedPx(context.heights[cut.section], cut.hf.variant);
      for (const kind of ['header', 'footer'] as const) {
        const strip = context.strip(cut.section, kind, cut.hf.variant, cut.display, cuts.length, cuts.filter(page => page.section === cut.section).length);
        // 功能：空白区域也提供双击入口；只读报告仍保留原生显示层；时间：2026-11-11；作者：lq。
        const region = strip ?? (this.session.editor.isEditable ? document.createElement('div') : undefined);
        if (!region) continue;
        region.classList.add(strip ? `mo-core-${kind}` : `mo-core-empty-${kind}`);
        region.dataset.hfSection = String(cut.section); region.dataset.hfKind = kind; region.dataset.hfVariant = cut.hf.variant;
        region.dataset.hfPage = String(index); region.dataset.hfNumber = cut.display;
        region.dataset.hfTotal = String(cuts.length); region.dataset.hfSectionPages = String(cuts.filter(page => page.section === cut.section).length);
        // 功能：宿主已经按纸张左上角定位，清除同源居中条带的半宽位移；时间：2026-11-11；作者：lq。
        region.style.position = 'absolute'; region.style.bottom = 'auto'; region.style.transform = 'none';
        region.style.left = `${left + margins.left / 15}px`;
        region.style.width = `${(settings.pageWidth - margins.left - margins.right) / 15}px`;
        region.style.top = `${top + (kind === 'header' ? (settings.headerDist ?? 720) / 15 : settings.pageHeight / 15 - (settings.footerDist ?? 720) / 15 - reserved.footerPx)}px`;
        if (this.session.editor.isEditable) { region.style.pointerEvents = 'auto'; region.style.minHeight = '20px'; region.title = `双击编辑${kind === 'header' ? '页眉' : '页脚'}`; }
        layer!.append(region);
      }
      const footnotes = context.notesElement(notes[index] ?? [], cut.section);
      if (footnotes) {
        footnotes.element.style.top = `${top + settings.pageHeight / 15 - effectiveBottomPx(settings, reserved.footerPx) - footnotes.height}px`;
        footnotes.element.style.left = `${left + margins.left / 15}px`;
        layer!.append(footnotes.element);
      }
    });
    if (context.endnotes.length && result.endnotesTop !== undefined) {
      const notes = context.notesElement(context.endnotes, context.sections.length - 1, 'endnote')!;
      notes.element.classList.add('mo-core-endnotes');
      const cut = [...cuts].reverse().find(cut => cut.start <= result.endnotesTop!) ?? cuts[0];
      notes.element.style.top = `${context.top + result.endnotesTop + cut.flowOffsetPx}px`;
      layer.append(notes.element);
    }
  }
}
