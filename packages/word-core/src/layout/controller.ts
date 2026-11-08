/** 功能：把原 App 的测量/重排调度抽成非 React 会话能力，继续使用上游分页算法；时间：2026-11-08；作者：lq。 */
import { measureBlocks } from '../upstream/apps/docs/src/renderer/pagination-measure';
import { sliceWithLineSplit } from '../upstream/apps/docs/src/renderer/pagination-lines';
import { FollowUpChain } from '../upstream/apps/docs/src/renderer/pagination-loop';
import type { SectionGeom, SliceOutputs } from '../upstream/apps/docs/src/renderer/pagination-types';
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
}

/** 功能：执行同源 DOM 测量、行盒补采样与分节切片，结果供我们的页面 UI 应用；时间：2026-11-08；作者：lq。 */
export function measureWordLayout(input: WordLayoutInput) {
  if (!(input.zoomFactor > 0) || (!input.sections.length && !input.sectionInfos?.length)) throw new Error('Word 排版缺少有效的缩放或分节几何');
  if (input.sections.length > 1 && !input.sectionInfos?.length) throw new Error('多节 Word 排版缺少来源范围映射');
  // 功能：承接上游 App 的单流测量状态，暂时去掉显示层列平移，完成后恢复；时间：2026-11-08；作者：lq。
  const hadMeasuring = input.root.classList.contains('measuring-columns');
  const neutralize = input.mirrorMargins || input.sections.some(section => (section.cols ?? 1) > 1) || input.sectionInfos?.some(section => section.settings.vAlign === 'center' || section.settings.vAlign === 'bottom' || !!sectionVertical(section.settings));
  if (neutralize) input.root.classList.add('measuring-columns');
  try {
    const measured = measureBlocks(input.root, input.origin, input.zoomFactor, input.planOf, false, input.carryOf);
    const effectiveSections = input.sectionInfos?.length ? liveSections(input.sectionInfos, measured.blocks, measured.sectBreaks, input.deletedSectionBreaks) : undefined;
    if (effectiveSections) assignSections(measured.blocks, effectiveSections);
    const geometries = effectiveSections ? sectionGeoms(effectiveSections, input.headerFooterHeights, input.evenOddHeaders) : input.sections;
    if (geometries.some(section => !(section.contentHeight > 0))) throw new Error('Word 排版包含无效的分节几何');
    const outputs: SliceOutputs = {};
    const slices = sliceWithLineSplit(measured.blocks, geometries, measured.totalHeight, input.zoomFactor, input.blockMetaOf, outputs);
    return { ...measured, slices, outputs, effectiveSections };
  } finally {
    if (!hadMeasuring) input.root.classList.remove('measuring-columns');
  }
}
export type WordLayoutResult = ReturnType<typeof measureWordLayout>;

/** 功能：按会话调度增量排版，合并请求并限制装饰引起的反复重排；时间：2026-11-08；作者：lq。 */
export class WordLayoutController {
  private timer: ReturnType<typeof setTimeout> | undefined;
  private disposed = false;
  private readonly chain = new FollowUpChain();
  constructor(private readonly read: () => WordLayoutInput, private readonly apply: (result: WordLayoutResult) => void, private readonly onError: (error: unknown) => void) {}
  /** 功能：文档、字体或页面设置变化开启新排版链；时间：2026-11-08；作者：lq。 */
  request(): void { this.chain.reset(); this.enqueue(); }
  /** 功能：装饰或尺寸补偿请求有界的后续测量，拒绝同因循环；时间：2026-11-08；作者：lq。 */
  followUp(reason: string, dirtyIndex: number | null = null): boolean {
    if (this.disposed || this.chain.request(reason, dirtyIndex) !== 'run') return false;
    this.enqueue();
    return true;
  }
  /** 功能：重排前读取最新状态，中文组合期间保留输入法拥有的 DOM；时间：2026-11-08；作者：lq。 */
  flush(): WordLayoutResult | undefined {
    if (this.disposed) return;
    if (this.timer !== undefined) clearTimeout(this.timer);
    this.timer = undefined;
    const input = this.read();
    if (input.composing) return;
    this.chain.beginPass();
    const result = measureWordLayout(input);
    this.apply(result);
    return result;
  }
  /** 功能：会话隐藏/关闭时清理迟到的回调；时间：2026-11-08；作者：lq。 */
  dispose(): void { this.disposed = true; if (this.timer !== undefined) clearTimeout(this.timer); this.timer = undefined; }
  /** 功能：同一轮事务只安排一次测量，并把异常交给宿主明确处理；时间：2026-11-08；作者：lq。 */
  private enqueue(): void {
    if (this.disposed || this.timer !== undefined) return;
    this.timer = setTimeout(() => { this.timer = undefined; try { this.flush(); } catch (error) { this.onError(error); } }, 0);
  }
}
