import type {
  DocBlock,
  DocImage,
  DocModel,
  DocParagraph,
  DocRun,
  DocTable,
  RunStyle,
} from '@mini-office/document-model';
import { uid } from '@mini-office/shared';

export function mapParagraphs(
  blocks: DocBlock[],
  fn: (paragraph: DocParagraph) => DocParagraph | null,
): DocBlock[] {
  const result: DocBlock[] = [];
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      const next = fn(block);
      if (next) result.push(next);
      continue;
    }
    if (block.type === 'table') {
      const table: DocTable = {
        ...block,
        rows: block.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({
            ...cell,
            blocks: mapParagraphs(cell.blocks, fn),
          })),
        })),
      };
      result.push(table);
      continue;
    }
    result.push(block);
  }
  return result;
}

export function updateParagraph(
  doc: DocModel,
  paragraphId: string,
  updater: (paragraph: DocParagraph) => DocParagraph,
): DocModel {
  return {
    ...doc,
    blocks: mapParagraphs(doc.blocks, (paragraph) =>
      paragraph.id === paragraphId ? updater(paragraph) : paragraph,
    ),
  };
}

export function findParagraph(doc: DocModel, paragraphId: string): DocParagraph | undefined {
  let found: DocParagraph | undefined;
  const visit = (blocks: DocBlock[]) => {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        if (block.id === paragraphId) found = block;
      } else if (block.type === 'table' && !found) {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
      }
    }
  };
  visit(doc.blocks);
  return found;
}

export function collectParagraphs(doc: DocModel): DocParagraph[] {
  const result: DocParagraph[] = [];
  const visit = (blocks: DocBlock[]) => {
    for (const block of blocks) {
      if (block.type === 'paragraph') result.push(block);
      else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
      }
    }
  };
  visit(doc.blocks);
  return result;
}

export function allRuns(paragraph: DocParagraph): DocRun[] {
  return paragraph.runs;
}

export function paragraphImages(paragraph: DocParagraph): Map<string, DocImage> {
  const map = new Map<string, DocImage>();
  for (const run of paragraph.runs) {
    if (run.image?.id) map.set(run.image.id, run.image);
  }
  return map;
}

export function computeDocStats(doc: DocModel): { words: number; characters: number } {
  let words = 0;
  let characters = 0;
  for (const paragraph of collectParagraphs(doc)) {
    const text = paragraph.runs
      .map((run) => (run.image ? '' : run.tab ? '\t' : run.text ?? ''))
      .join('');
    const cjk = text.match(/[\u3400-\u9fff\uf900-\ufaff]/g)?.length ?? 0;
    const latin = text
      .replace(/[\u3400-\u9fff\uf900-\ufaff]/g, ' ')
      .trim()
      .split(/\s+/)
      .filter(Boolean).length;
    words += cjk + latin;
    characters += text.replace(/\s/g, '').length;
  }
  return { words, characters };
}

/**
 * 功能：run 的模型长度口径——图片/制表符/换行/分页符/合并域各计 1，其余按字符数；
 * 与 splitRunsAtOffset 保持一致，供分页器二分断点使用；时间：2026-10-30；作者：lq。
 */
export function runLength(run: DocRun): number {
  if (run.image || run.tab || run.lineBreak || run.mergeField || run.pageBreak) return 1;
  return run.text?.length ?? 0;
}

/** 功能：模型 runs 的总长度；时间：2026-10-30；作者：lq。 */
export function runsLength(runs: DocRun[]): number {
  return runs.reduce((sum, run) => sum + runLength(run), 0);
}

/** 功能：段落是否含有可见内容；空首片表示"此处仅断页"；时间：2026-10-30；作者：lq。 */
export function paragraphHasContent(paragraph: DocParagraph): boolean {
  return paragraph.runs.some((run) => Boolean(run.text || run.image || run.tab || run.lineBreak || run.pageBreak));
}

/**
 * 功能：跨页首片编辑回写模型——用片段中的新内容替换模型段落的前缀，保留被切到续页的尾部；
 * 没有边界信息时退化为整体替换；时间：2026-10-30；作者：lq。
 */
export function mergeFragmentRuns(
  modelRuns: DocRun[],
  headEnd: number,
  headRuns: DocRun[],
  fallback: RunStyle,
): DocRun[] {
  const [, after] = splitRunsAtOffset(modelRuns, headEnd, fallback);
  return [...headRuns, ...after];
}

export function splitRunsAtOffset(
  runs: DocRun[],
  offset: number,
  fallback: RunStyle,
): [DocRun[], DocRun[]] {
  const before: DocRun[] = [];
  const after: DocRun[] = [];
  let consumed = 0;
  for (const run of runs) {
    if (run.image || run.tab || run.lineBreak || run.mergeField || run.pageBreak) {
      if (consumed < offset) before.push(run);
      else after.push(run);
      consumed += 1;
      continue;
    }
    const text = run.text ?? '';
    const length = text.length;
    if (consumed + length <= offset) {
      before.push(run);
    } else if (consumed >= offset) {
      after.push(run);
    } else {
      const local = offset - consumed;
      before.push({ ...run, text: text.slice(0, local) });
      after.push({ ...run, text: text.slice(local) });
    }
    consumed += length;
  }
  if (after.length === 0) after.push({ text: '', style: fallback });
  if (before.length === 0) before.push({ text: '', style: fallback });
  return [before, after];
}

export function makeParagraph(
  runs: DocRun[] = [{ text: '', style: {} }],
  base?: Partial<DocParagraph>,
): DocParagraph {
  return {
    type: 'paragraph',
    id: uid('p'),
    runs,
    style: { ...(base?.style ?? {}) },
    list: base?.list,
    styleId: base?.styleId,
  };
}

export function totalTableColumns(table: DocTable): number {
  return Math.max(
    table.columnWidthsPx.length,
    ...table.rows.map((row) => row.cells.reduce((sum, cell) => sum + cell.colSpan, 0)),
  );
}

export interface CellPlacement {
  cell: DocTable['rows'][number]['cells'][number];
  col: number;
  row: number;
}

export function tableCellPlacements(table: DocTable): CellPlacement[] {
  const placements: CellPlacement[] = [];
  const active: Array<{ start: number; span: number; remaining: number }> = [];
  const columnCount = totalTableColumns(table);
  table.rows.forEach((row, rowIndex) => {
    let column = 0;
    let cellIndex = 0;
    while (column < columnCount || cellIndex < row.cells.length) {
      const spanning = active.find((item) => item.start === column && item.remaining > 0);
      if (spanning) {
        spanning.remaining -= 1;
        column += spanning.span;
        continue;
      }
      const cell = row.cells[cellIndex];
      if (!cell) break;
      placements.push({ cell, col: column, row: rowIndex });
      if (cell.rowSpan > 1) {
        active.push({ start: column, span: cell.colSpan, remaining: cell.rowSpan - 1 });
      }
      column += cell.colSpan;
      cellIndex += 1;
    }
  });
  return placements;
}
