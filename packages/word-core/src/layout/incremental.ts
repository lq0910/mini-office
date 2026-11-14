/** 功能：提取 GenOffice App.tsx 的续排边界与前缀采样复用，保持同源切片规则；时间：2026-11-13；作者：lq。 */
import { isResumePage, type PassResume } from '../upstream/apps/docs/src/renderer/pagination-lines';
import { pageOfY } from '../upstream/apps/docs/src/renderer/pagination-slices';
import type { BlockBox, PageSlice, SectionGeom, SliceOutputs } from '../upstream/apps/docs/src/renderer/pagination-types';

/** 功能：缓存只携带上轮自然流几何；外部字体、纸张及样式变化不提供此缓存；时间：2026-11-13；作者：lq。 */
export interface WordLayoutPrefix {
  blocks: BlockBox[];
  slices: PageSlice[];
  outputs: SliceOutputs;
  geometries: SectionGeom[];
  sectionSignature: string;
  dirtyElement?: Element | null;
}

/** 功能：与上游一致比较几何身份，文本改变即使同高也从事务块之前重新采样；时间：2026-11-13；作者：lq。 */
function sameBox(a: BlockBox, b: BlockBox): boolean {
  return a.el === b.el && Math.abs(a.top - b.top) < 0.01 && Math.abs(a.height - b.height) < 0.01
    && a.section === b.section && a.docxIndex === b.docxIndex && !!a.floated === !!b.floated
    && !!a.emptyPara === !!b.emptyPara && (a.spaceBeforePx ?? 0) === (b.spaceBeforePx ?? 0);
}

/** 功能：仅在未变前缀中寻找块首起页，并向前保留一页满足孤行/与下段同页约束；时间：2026-11-13；作者：lq。 */
export function wordLayoutResume(blocks: BlockBox[], geometries: SectionGeom[], previous?: WordLayoutPrefix, sectionSignature = ''): PassResume | undefined {
  if (!previous?.dirtyElement || previous.slices.length < 2 || !previous.blocks.length
    || sectionSignature !== previous.sectionSignature || JSON.stringify(geometries) !== JSON.stringify(previous.geometries)) return;
  const transactionBlock = blocks.findIndex(block => block.el === previous.dirtyElement);
  if (transactionBlock < 0) return;
  const count = Math.min(blocks.length, previous.blocks.length);
  let dirty = 0;
  while (dirty < count && sameBox(blocks[dirty], previous.blocks[dirty])) dirty++;
  dirty = Math.min(dirty, transactionBlock, blocks.length - 1);
  if (dirty <= 0) return;
  const dirtyPage = pageOfY(previous.slices, blocks[dirty].top);
  for (let page = dirtyPage - 1; page > 0; page--) {
    if (!isResumePage(previous.slices, page)) continue;
    const start = previous.slices[page].start;
    const block = blocks.findIndex((item, index) => index < dirty && Math.abs(item.top - start) < 0.5);
    if (block < 0 || blocks[block].floated) continue;
    for (let index = 0; index < block; index++) {
      const from = previous.blocks[index], to = blocks[index];
      if (from.lineBoxes) to.lineBoxes = from.lineBoxes;
      if (from.lineLeadPx !== undefined) to.lineLeadPx = from.lineLeadPx;
      if (from.tableRows) to.tableRows = from.tableRows;
      if (from.colWraps) to.colWraps = from.colWraps;
      if (from.oversizeLineH !== undefined) to.oversizeLineH = from.oversizeLineH;
    }
    return { prevSlices: previous.slices, page, block, prevOut: previous.outputs };
  }
}
