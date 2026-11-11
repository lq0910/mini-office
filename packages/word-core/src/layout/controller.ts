/** 功能：把原 App 的测量/重排调度抽成非 React 会话能力，继续使用上游分页算法；时间：2026-11-08；作者：lq。 */
import { measureBlocks, appendEndnotesBlock, appendFloatSpillBlock } from '../upstream/apps/docs/src/renderer/pagination-measure';
import { sliceWithLineSplit } from '../upstream/apps/docs/src/renderer/pagination-lines';
import { FollowUpChain } from '../upstream/apps/docs/src/renderer/pagination-loop';
import type { SectionGeom, SliceOutputs, PageNoteItem } from '../upstream/apps/docs/src/renderer/pagination-types';
import { FOOTNOTE_SEPARATOR_H } from '../upstream/apps/docs/src/renderer/line-metrics';
import { assignSections, liveSections, sectionGeoms, sectionVertical } from '../upstream/apps/docs/src/renderer/pagination-sections';
import type { SectionInfo } from '../upstream/packages/docx-engine/src/index';

/** 功能：调用方显式传入当前文档几何与映射，禁止猜测页面尺寸或返回伪造页数；时间：2026-11-08；作者：lq。 */
export interface WordLayoutInput {
  root: HTMLElement;
  origin: number;
  zoomFactor: number;
  sections: SectionGeom[];
  sectionInfos?: SectionInfo[];
  headerFooterHeights?: Parameters<typeof sectionGeoms>[1];
  evenOddHeaders?: boolean;
  deletedSectionBreaks?: Set<number>;
  mirrorMargins?: boolean;
  planOf?: Parameters<typeof measureBlocks>[3];
  carryOf?: Parameters<typeof measureBlocks>[5];
  blockMetaOf?: Parameters<typeof sliceWithLineSplit>[4];
  composing?: boolean;
  /** 功能：尾注区参与同源分页，不在排版结束后盖到正文上；时间：2026-11-09；作者：lq。 */
  endnotes?: PageNoteItem[];
}

/** 功能：执行同源 DOM 测量、行盒补采样与分节切片，结果供我们的页面 UI 应用；时间：2026-11-08；作者：lq。 */
export function measureWordLayout(input: WordLayoutInput) {
  if (!(input.zoomFactor > 0) || (!input.sections.length && !input.sectionInfos?.length)) throw new Error('Word 排版缺少有效的缩放或分节几何');
  if (input.sections.length > 1 && !input.sectionInfos?.length) throw new Error('多节 Word 排版缺少来源范围映射');
  // 功能：承接上游 App 的单流测量状态，暂时去掉显示层列平移，完成后恢复；时间：2026-11-08；作者：lq。
  const hadMeasuring = input.root.classList.contains('measuring-columns');
  // 功能：测量自然行流，避免上一轮页间装饰强制的换行反过来锁死分页；时间：2026-11-09；作者：lq。
  const hadNatural = input.root.classList.contains('measuring-natural');
  input.root.classList.add('measuring-natural');
  const neutralize = input.mirrorMargins || input.sections.some(section => (section.cols ?? 1) > 1) || input.sectionInfos?.some(section => section.settings.vAlign === 'center' || section.settings.vAlign === 'bottom' || !!sectionVertical(section.settings));
  if (neutralize) input.root.classList.add('measuring-columns');
  try {
    const measured = measureBlocks(input.root, input.origin, input.zoomFactor, input.planOf, false, input.carryOf);
    const effectiveSections = input.sectionInfos?.length ? liveSections(input.sectionInfos, measured.blocks, measured.sectBreaks, input.deletedSectionBreaks) : undefined;
    if (effectiveSections) assignSections(measured.blocks, effectiveSections);
    // 功能：继承尾注和浮动对象末尾溢出的容量计算，让对应纸张真实存在；时间：2026-11-09；作者：lq。
    const endnotes = appendEndnotesBlock(measured.blocks, measured.totalHeight, input.endnotes ?? [], FOOTNOTE_SEPARATOR_H);
    const lastSection = effectiveSections?.[effectiveSections.length - 1]?.settings;
    const totalHeight = appendFloatSpillBlock(measured.blocks, endnotes?.totalHeight ?? measured.totalHeight, measured.floats, (lastSection?.marginBottom ?? 0) / 15) ?? endnotes?.totalHeight ?? measured.totalHeight;
    const geometries = effectiveSections ? sectionGeoms(effectiveSections, input.headerFooterHeights, input.evenOddHeaders) : input.sections;
    if (geometries.some(section => !(section.contentHeight > 0))) throw new Error('Word 排版包含无效的分节几何');
    // 功能：开启上游表格行高、分行和浮动对象输出，供正文宿主统一应用；时间：2026-11-09；作者：lq。
    const outputs: SliceOutputs = { rowFills: [], rowSplits: [], floatVShifts: [], floatFlows: [], floatSplits: [], oversizeClips: [] };
    const slices = sliceWithLineSplit(measured.blocks, geometries, totalHeight, input.zoomFactor, input.blockMetaOf, outputs);
    return { ...measured, totalHeight, slices, outputs, effectiveSections, geometries, endnotesTop: endnotes?.top };
  } finally {
    if (!hadNatural) input.root.classList.remove('measuring-natural');
    if (!hadMeasuring) input.root.classList.remove('measuring-columns');
  }
}
export type WordLayoutResult = ReturnType<typeof measureWordLayout>;

/** 功能：按会话调度增量排版，合并请求并限制装饰引起的反复重排；时间：2026-11-08；作者：lq。 */
export class WordLayoutController {
  private timer: ReturnType<typeof setTimeout> | undefined;
  private disposed = false;
  // 功能：沿用 GenOffice 的编辑占空比和后续轮次间隔，给输入及滚动保留主线程；时间：2026-11-11；作者：lq。
  private lastPassMs = 0;
  private measured = false;
  private readonly chain = new FollowUpChain();
  constructor(private readonly read: () => WordLayoutInput, private readonly apply: (result: WordLayoutResult) => void, private readonly onError: (error: unknown) => void,
    // 功能：宿主在同一未缩放自然流窗口内完成测量与锚点解析，最后恢复显示；时间：2026-11-11；作者：lq。
    private readonly inMeasure: (run: () => WordLayoutResult | undefined) => WordLayoutResult | undefined = run => run()) {}
  /** 功能：文档、字体或页面设置变化开启新排版链；时间：2026-11-08；作者：lq。 */
  request(): void { this.chain.reset(); this.enqueue(this.measured ? Math.min(2000, Math.max(300, this.lastPassMs * 3)) : 0, true); }
  /** 功能：装饰或尺寸补偿请求有界的后续测量，拒绝同因循环；时间：2026-11-08；作者：lq。 */
  followUp(reason: string, dirtyIndex: number | null = null): boolean {
    if (this.disposed || this.chain.request(reason, dirtyIndex) !== 'run') return false;
    this.enqueue(100);
    return true;
  }
  /** 功能：重排前读取最新状态，中文组合期间保留输入法拥有的 DOM；时间：2026-11-08；作者：lq。 */
  flush(): WordLayoutResult | undefined {
    if (this.disposed) return;
    if (this.timer !== undefined) clearTimeout(this.timer);
    this.timer = undefined;
    const started = performance.now();
    try { return this.inMeasure(() => {
      const input = this.read();
      if (input.composing) return;
      this.chain.beginPass();
      // 功能：保持自然流到宿主解析完所有页界，避免两次恢复/隐藏装饰使表格几何漂移；时间：2026-11-11；作者：lq。
      const natural = input.root.classList.contains('measuring-natural');
      input.root.classList.add('measuring-natural');
      try { const result = measureWordLayout(input); this.apply(result); this.measured = true; return result; }
      finally { if (!natural) input.root.classList.remove('measuring-natural'); }
    }); } finally { this.lastPassMs = performance.now() - started; }
  }
  /** 功能：会话隐藏/关闭时清理迟到的回调；时间：2026-11-08；作者：lq。 */
  dispose(): void { this.disposed = true; if (this.timer !== undefined) clearTimeout(this.timer); this.timer = undefined; }
  /** 功能：同一轮事务只安排一次测量，并把异常交给宿主明确处理；时间：2026-11-08；作者：lq。 */
  private enqueue(delay: number, restart = false): void {
    if (this.disposed) return;
    if (this.timer !== undefined) { if (!restart) return; clearTimeout(this.timer); }
    this.timer = setTimeout(() => { this.timer = undefined; try { this.flush(); } catch (error) { this.onError(error); } }, delay);
  }
}
