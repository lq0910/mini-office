/** 功能：跨页合并格的长文字按完整行放入实际正文片段，页缝不吞字且不改变来源对齐；时间：2026-11-11；作者：lq。 */
export interface WordCellLine { top: number; bottom: number; }
/** 功能：优先保留原始位置；余下空间不足时从首片段重排，只有所有行能完整容纳才返回位移；时间：2026-11-11；作者：lq。 */
export function wordCellLineShifts(lines: readonly WordCellLine[], segments: readonly WordCellLine[]): number[] | undefined {
  if (!lines.length || !segments.length) return;
  const place = (initial: number): number[] | undefined => {
    const shifts: number[] = []; let segment = 0, shift = initial;
    for (const line of lines) {
      while (segment < segments.length) {
        const slot = segments[segment];
        shift = Math.max(shift, slot.top - line.top);
        if (line.bottom + shift <= slot.bottom + 0.1) break;
        segment++;
      }
      if (segment === segments.length) return;
      shifts.push(shift);
    }
    return shifts;
  };
  return place(0) ?? place(segments[0].top - lines[0].top);
}
