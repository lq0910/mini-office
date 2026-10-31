import type { DocBlock, DocComment, DocModel, DocRun } from '@mini-office/document-model';
import { mapParagraphs } from './doc-utils';

/** 功能：run 的模型长度（原子 run 计 1）；时间：2026-10-31；作者：hou0。 */
function modelLength(runs: DocRun[]): number {
  return runs.reduce((sum, run) => sum + (run.tab || run.image || run.lineBreak || run.mergeField ? 1 : (run.text ?? '').length), 0);
}

/** 功能：按 [start,end) 切 run，保留区间外/内的部分按需截断；时间：2026-10-31；作者：hou0。 */
function sliceRuns(runs: DocRun[], start: number, end: number, keepInside: boolean): DocRun[] {
  const output: DocRun[] = [];
  let position = 0;
  for (const run of runs) {
    const length = modelLength([run]);
    const runEnd = position + length;
    const inside = runEnd > start && position < end;
    if (!inside) {
      if (length > 0 || run.tab || run.image) output.push(run);
    } else if (run.tab || run.image || run.lineBreak || run.mergeField || run.field) {
      if (keepInside) output.push(run);
    } else {
      const text = run.text ?? '';
      const from = Math.max(0, start - position);
      const to = Math.min(length, end - position);
      if (keepInside) {
        const piece = text.slice(from, to);
        if (piece) output.push({ ...run, text: piece });
      } else {
        if (from > 0) output.push({ ...run, text: text.slice(0, from) });
        if (to < length) output.push({ ...run, text: text.slice(to) });
      }
    }
    position = runEnd;
  }
  return output;
}

/** 功能：给 [start,end) 范围内的 run 打上批注 id（自动切分边界 run）；时间：2026-10-31；作者：hou0。 */
export function attachCommentRange(blocks: DocBlock[], paragraphId: string, start: number, end: number, commentId: string): DocBlock[] {
  return mapParagraphs(blocks, (paragraph) => {
    if (paragraph.id !== paragraphId) return paragraph;
    const runs: DocRun[] = [];
    let position = 0;
    for (const run of paragraph.runs) {
      const length = modelLength([run]);
      const runEnd = position + length;
      const intersects = runEnd > start && position < end;
      if (!intersects) {
        runs.push(run);
      } else {
        const ids = [...(run.commentIds ?? []), commentId];
        const text = run.text ?? '';
        const isAtomic = run.tab || run.image || run.lineBreak || run.mergeField || run.field;
        if (isAtomic) {
          runs.push({ ...run, commentIds: ids });
        } else {
          const from = Math.max(0, start - position);
          const to = Math.min(length, end - position);
          if (from > 0) runs.push({ ...run, text: text.slice(0, from) });
          runs.push({ ...run, text: text.slice(from, to), commentIds: ids });
          if (to < length) runs.push({ ...run, text: text.slice(to) });
        }
      }
      position = runEnd;
    }
    return { ...paragraph, runs: runs.filter((run) => (run.text ?? '').length > 0 || run.tab || run.image || run.lineBreak || run.mergeField || run.field || run.control) };
  });
}

/** 功能：从所有 run 上摘除某条批注的锚点；时间：2026-10-31；作者：hou0。 */
export function stripCommentAnchors(blocks: DocBlock[], commentId: string): DocBlock[] {
  return mapParagraphs(blocks, (paragraph) => {
    if (!paragraph.runs.some((run) => run.commentIds?.includes(commentId))) return paragraph;
    return {
      ...paragraph,
      runs: paragraph.runs.map((run) => {
        if (!run.commentIds?.includes(commentId)) return run;
        const rest = run.commentIds.filter((id) => id !== commentId);
        return rest.length ? { ...run, commentIds: rest } : { ...run, commentIds: undefined };
      }),
    };
  });
}

/** 功能：聚合批注范围的当前文本（批注引用变化后同步面板）；时间：2026-10-31；作者：hou0。 */
export function gatherCommentText(blocks: DocBlock[], commentId: string): string {
  const parts: string[] = [];
  const visit = (items: DocBlock[]) => {
    for (const block of items) {
      if (block.type !== 'paragraph') {
        if (block.type === 'table') for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
        continue;
      }
      for (const run of block.runs) {
        if (run.commentIds?.includes(commentId) && run.text) parts.push(run.text);
      }
    }
  };
  visit(blocks);
  return parts.join('');
}

/** 功能：下一批注 id（正文已有锚点的最大数字 +1）；时间：2026-10-31；作者：hou0。 */
export function nextCommentId(doc: DocModel): string {
  let max = 0;
  for (const comment of doc.comments ?? []) {
    const numeric = Number.parseInt(comment.id, 10);
    if (Number.isFinite(numeric) && numeric > max) max = numeric;
  }
  return String(max + 1);
}

/**
 * 功能：接受/拒绝一处修订——插入：接受去标记/拒绝删 run；删除：接受删 run/拒绝去标记恢复；
 * 时间：2026-10-31；作者：hou0。
 */
export function applyRevisionDecision(blocks: DocBlock[], paragraphId: string, runIndex: number, accept: boolean): DocBlock[] {
  return mapParagraphs(blocks, (paragraph) => {
    if (paragraph.id !== paragraphId) return paragraph;
    if (runIndex < 0 || runIndex >= paragraph.runs.length) return paragraph;
    const run = paragraph.runs[runIndex];
    if (!run.revision) return paragraph;
    const runs = [...paragraph.runs];
    const isInsert = run.revision.type === 'insert';
    if ((isInsert && accept) || (!isInsert && !accept)) {
      runs.splice(runIndex, 1, { ...run, revision: undefined });
    } else {
      runs.splice(runIndex, 1);
    }
    return { ...paragraph, runs: runs.filter((r) => r.revision?.type === 'delete' || (r.text ?? '').length > 0 || r.tab || r.image || r.lineBreak || r.mergeField || r.field || r.control) };
  });
}

/** 功能：扫描全文档的修订 run（审阅窗格数据），按文档顺序；时间：2026-10-31；作者：hou0。 */
export function collectRevisions(blocks: DocBlock[]): Array<{ paragraphId: string; runIndex: number; type: 'insert' | 'delete'; author?: string; text: string }> {
  const result: Array<{ paragraphId: string; runIndex: number; type: 'insert' | 'delete'; author?: string; text: string }> = [];
  const visit = (items: DocBlock[]) => {
    for (const block of items) {
      if (block.type !== 'paragraph') {
        if (block.type === 'table') for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
        continue;
      }
      block.runs.forEach((run, runIndex) => {
        if (run.revision) {
          result.push({ paragraphId: block.id, runIndex, type: run.revision.type, ...(run.revision.author ? { author: run.revision.author } : {}), text: run.text ?? '' });
        }
      });
    }
  };
  visit(blocks);
  return result;
}
