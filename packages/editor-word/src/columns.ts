import type { DocBlock } from '@mini-office/document-model';

/**
 * 功能：把一个多栏分节内的块分配到各栏——有显式分栏符（段首 w:br type="column"）时按分栏符切栏，
 * 否则按块数均衡均分（对齐 Word 连续分节的栏平衡）；渲染与高度测量共用同一规则，避免两者不一致；
 * 时间：2026-10-30；作者：hou0。
 */
export function distributeColumns(blocks: DocBlock[], count: number): DocBlock[][] {
  const columns = Math.max(1, count);
  if (columns === 1) return [blocks];
  // 功能：节尾只用来挂 w:sectPr 的空段落不占行——Word 导出 PDF 实测勾选行→下一段 = 42px，
  // 恰为「一行正文 + 后续空段一行」，不含分节标记行；把它排进栏会让锚定其后的浮动方框低出一行；
  // 时间：2026-10-30；作者：hou0。
  let effective = blocks;
  while (effective.length > 1) {
    const last = effective[effective.length - 1];
    if (last.type === 'paragraph' && last.sectionEnd && !paragraphVisible(last)) effective = effective.slice(0, -1);
    else break;
  }
  const segments: DocBlock[][] = [[]];
  for (const block of effective) {
    const breaksBefore = block.type === 'paragraph' && block.style.columnBreakBefore === true;
    if (breaksBefore && segments[segments.length - 1].length > 0) segments.push([]);
    segments[segments.length - 1].push(block);
  }
  // 功能：分栏符切出的段数不足或超出栏数时退回均衡均分，避免空栏或内容溢出栏外；时间：2026-10-30；作者：hou0。
  if (segments.length > 1 && segments.length <= columns) {
    return Array.from({ length: columns }, (_, index) => segments[index] ?? []);
  }
  const perColumn = Math.ceil(effective.length / columns);
  return Array.from({ length: columns }, (_, index) => effective.slice(index * perColumn, (index + 1) * perColumn));
}

/** 功能：段落是否有可见内容（纯分节标记段落没有）；时间：2026-10-30；作者：hou0。 */
function paragraphVisible(paragraph: Extract<DocBlock, { type: 'paragraph' }>): boolean {
  if (paragraph.list) return true;
  return paragraph.runs.some(
    (run) =>
      (run.text ?? '').trim().length > 0 ||
      run.image ||
      run.control ||
      run.tab ||
      run.lineBreak ||
      run.pageBreak ||
      run.mergeField,
  );
}
