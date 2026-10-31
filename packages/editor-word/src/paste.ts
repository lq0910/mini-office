import type { DocBlock, DocParagraph, DocRun, RunStyle } from '@mini-office/document-model';

/**
 * 功能：把粘贴文本拆成多段 run——按换行分段，与 Word 粘贴多行文本产生独立段落一致；
 * 空行产出空段落；时间：2026-10-30；作者：hou0。
 */
export function splitPasteLines(text: string, style: RunStyle): DocRun[][] {
  const normalized = text.replace(/\r\n?/g, '\n');
  return normalized
    .split('\n')
    .map((line) => (line ? [{ text: line, style: { ...style } }] : []));
}

/**
 * 功能：把段落并回同层级的上一段落（递归进表格单元格）——单元格内首段不并出单元格，
 * 与 Word 一致；返回合并后的块数组与插入点信息，null 表示不可合并；时间：2026-10-30；作者：hou0。
 */
export function mergeIntoPreviousParagraph(
  blocks: DocBlock[],
  paragraphId: string,
): { blocks: DocBlock[]; caretParagraphId: string; caretOffset: number } | null {
  const index = blocks.findIndex((block) => block.type === 'paragraph' && block.id === paragraphId);
  if (index > 0) {
    const previous = blocks[index - 1];
    const current = blocks[index];
    if (previous.type !== 'paragraph' || current.type !== 'paragraph') return null;
    const caretOffset = modelLength(previous.runs);
    const next = blocks
      .map((block, position) => (position === index - 1 ? { ...block, runs: [...previous.runs, ...current.runs] } : block))
      .filter((_, position) => position !== index);
    return { blocks: next, caretParagraphId: previous.id, caretOffset };
  }
  for (let position = 0; position < blocks.length; position += 1) {
    const block = blocks[position];
    if (block.type !== 'table') continue;
    for (let row = 0; row < block.rows.length; row += 1) {
      for (let column = 0; column < block.rows[row].cells.length; column += 1) {
        const nested = mergeIntoPreviousParagraph(block.rows[row].cells[column].blocks, paragraphId);
        if (!nested) continue;
        const nextTable: DocBlock = {
          ...block,
          rows: block.rows.map((r, ri) =>
            ri !== row
              ? r
              : { ...r, cells: r.cells.map((c, ci) => (ci !== column ? c : { ...c, blocks: nested.blocks })) },
          ),
        };
        return {
          blocks: blocks.map((b, bi) => (bi === position ? nextTable : b)),
          caretParagraphId: nested.caretParagraphId,
          caretOffset: nested.caretOffset,
        };
      }
    }
  }
  return null;
}

/**
 * 功能：把多行粘贴文本接在锚定段落的插入点后——首行并入锚定段（替换 [start,end) 选区），
 * 其余行成为独立段落（继承锚定段的段落样式、列表与样式关联），尾段承接原插入点之后的 runs；
 * 返回新的顶层块数组、末段 id 与光标偏移；时间：2026-10-30；作者：hou0。
 */
export function pasteMultiline(
  blocks: DocBlock[],
  anchorParagraph: DocParagraph,
  start: number,
  end: number,
  text: string,
  runStyle: RunStyle,
  newId: (prefix: string) => string,
): { blocks: DocBlock[]; lastParagraphId: string; caretOffset: number } | null {
  const lines = splitPasteLines(text, runStyle);
  if (!lines.length) return null;
  // 功能：sliceRuns 保留区间外内容——前段取 [start, len) 的补集即 [0,start)，尾段取 [0,end) 的补集即 [end, len)；时间：2026-10-30；作者：hou0。
  const runsBefore = sliceRuns(anchorParagraph.runs, start, modelLength(anchorParagraph.runs));
  const runsAfter = sliceRuns(anchorParagraph.runs, 0, end);
  const paragraphs: DocParagraph[] = [];
  lines.forEach((lineRuns, index) => {
    const runs = index === 0 ? [...runsBefore, ...lineRuns] : [...lineRuns];
    if (index === lines.length - 1) runs.push(...runsAfter);
    paragraphs.push({
      type: 'paragraph',
      id: index === 0 ? anchorParagraph.id : newId('p'),
      runs,
      style: { ...anchorParagraph.style },
      list: anchorParagraph.list ? { ...anchorParagraph.list } : undefined,
      styleId: anchorParagraph.styleId,
    });
  });
  const replaced = replaceParagraph(blocks, anchorParagraph.id, paragraphs);
  if (!replaced) return null;
  const last = paragraphs[paragraphs.length - 1];
  return {
    blocks: replaced,
    lastParagraphId: last.id,
    caretOffset:
      lines[lines.length - 1].reduce((sum, run) => sum + modelLength([run]), 0) +
      runsAfter.reduce((sum, run) => sum + modelLength([run]), 0),
  };
}

function replaceParagraph(blocks: DocBlock[], paragraphId: string, replacements: DocParagraph[]): DocBlock[] | null {
  const result: DocBlock[] = [];
  let replaced = false;
  for (const block of blocks) {
    if (block.type === 'paragraph' && block.id === paragraphId && !replaced) {
      result.push(...replacements);
      replaced = true;
      continue;
    }
    if (block.type === 'table' && !replaced) {
      let tableReplaced = false;
      const table: DocBlock = {
        ...block,
        rows: block.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => {
            if (tableReplaced) return cell;
            const nested = replaceParagraph(cell.blocks, paragraphId, replacements);
            if (nested) {
              tableReplaced = true;
              return { ...cell, blocks: nested };
            }
            return cell;
          }),
        })),
      };
      if (tableReplaced) {
        result.push(table);
        replaced = true;
        continue;
      }
    }
    result.push(block);
  }
  return replaced ? result : null;
}

function sliceRuns(runs: DocRun[], start: number, end: number): DocRun[] {
  const output: DocRun[] = [];
  let position = 0;
  for (const run of runs) {
    const length = modelLength([run]);
    const runEnd = position + length;
    if (runEnd <= start || position >= end) {
      if (length > 0) output.push(run);
    } else if (run.tab || run.image || run.lineBreak || run.mergeField) {
      // 原子 run 落在替换区间内：丢弃
    } else {
      const keepFrom = Math.max(0, start - position);
      const keepTo = Math.min(length, end - position);
      const text = run.text ?? '';
      if (keepFrom > 0) output.push({ ...run, text: text.slice(0, keepFrom) });
      if (keepTo < length) output.push({ ...run, text: text.slice(keepTo) });
    }
    position = runEnd;
  }
  return output.filter((run) => (run.text ?? '').length > 0 || run.tab || run.image || run.lineBreak || run.mergeField);
}

function modelLength(runs: DocRun[]): number {
  return runs.reduce((sum, run) => sum + (run.tab || run.image || run.lineBreak || run.mergeField ? 1 : (run.text ?? '').length), 0);
}
