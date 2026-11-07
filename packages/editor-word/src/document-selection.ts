/** 功能：跨段落替换在模型层完成，保护分页 DOM 和表格结构；时间：2026-10-18；作者：lq。 */
import type { DocBlock, DocModel, DocParagraph, DocRun } from '@mini-office/document-model';
import { collectParagraphs, makeParagraph, mapParagraphs, splitRunsAtOffset } from './doc-utils';
import { runsPlainText } from './editing';
import { renderRunsHtml } from './html';

export interface ParagraphSlice { paragraphId: string; start: number; end: number }

/** 功能：提取正文选区文本，复制时排除页眉页脚和界面文字；时间：2026-10-18；作者：lq。 */
export function selectionText(doc: DocModel, slices: ParagraphSlice[]): string {
  const paragraphs = new Map(collectParagraphs(doc).map((paragraph) => [paragraph.id, paragraph]));
  return slices.map((slice) => runsPlainText(paragraphs.get(slice.paragraphId)?.runs ?? []).slice(slice.start, slice.end)).join('\n');
}

/** 功能：复制时附带选中文字样式，剔除分页容器和界面结构；时间：2026-10-18；作者：lq。 */
export function selectionHtml(doc: DocModel, slices: ParagraphSlice[]): string {
  const paragraphs = new Map(collectParagraphs(doc).map((paragraph) => [paragraph.id, paragraph]));
  return slices.map((slice) => {
    const paragraph = paragraphs.get(slice.paragraphId);
    if (!paragraph) return '';
    const [, fromStart] = splitRunsAtOffset(paragraph.runs, slice.start, {});
    const [selected] = splitRunsAtOffset(fromStart, slice.end - slice.start, {});
    return `<p>${renderRunsHtml(selected)}</p>`;
  }).join('');
}

/** 功能：截取选区外的原始 run，保留文字样式和图片等原子内容；时间：2026-10-18；作者：lq。 */
function outsideRuns(paragraph: DocParagraph, slice: ParagraphSlice): [DocRun[], DocRun[]] {
  const fallback = paragraph.runs[0]?.style ?? {};
  const [before] = splitRunsAtOffset(paragraph.runs, slice.start, fallback);
  const [, after] = splitRunsAtOffset(paragraph.runs, slice.end, fallback);
  return [before, after];
}

/** 功能：同一正文/单元格内合并选区端点；跨单元格仅替换内容以保留有效表格；时间：2026-10-18；作者：lq。 */
export function replaceDocumentSelection(doc: DocModel, slices: ParagraphSlice[], text: string, rich?: DocParagraph[]): { doc: DocModel; paragraphId: string; offset: number } | null {
  const paragraphs = collectParagraphs(doc);
  const byId = new Map(paragraphs.map((paragraph) => [paragraph.id, paragraph]));
  const first = slices[0];
  const last = slices.at(-1);
  if (!first || !last || !byId.has(first.paragraphId) || !byId.has(last.paragraphId)) return null;
  const start = byId.get(first.paragraphId)!;
  const end = byId.get(last.paragraphId)!;
  const style = start.runs[0]?.style ?? {};
  const [before] = outsideRuns(start, first);
  const [, after] = outsideRuns(end, last);
  // 功能：富文本粘贴复用选区替换，逐段保留安全解析后的样式；时间：2026-11-07；作者：lq。
  const lines = rich ? rich.map((paragraph) => runsPlainText(paragraph.runs)) : text.replace(/\r\n?/g, '\n').split('\n');
  const insertedRuns = (index: number): DocRun[] => rich?.[index]?.runs ?? [{ text: lines[index], style }];
  let paragraphId = start.id;
  let offset = first.start + lines[0].length;
  const selected = new Map(slices.map((slice) => [slice.paragraphId, slice]));
  const coversAll = paragraphs.length === slices.length && paragraphs.every((paragraph) => {
    const slice = selected.get(paragraph.id);
    return slice && slice.start === 0 && slice.end >= runsPlainText(paragraph.runs).length;
  });
  // 全文替换移除正文表格，但保留页面设置和只读页眉页脚；时间：2026-10-18；作者：lq。
  if (coversAll) {
    const blocks = lines.map((line, index) => makeParagraph(insertedRuns(index), rich?.[index] ?? start));
    return { doc: { ...doc, blocks }, paragraphId: blocks.at(-1)!.id, offset: lines.at(-1)!.length };
  }
  let replaced = false;
  const visit = (blocks: DocBlock[]): DocBlock[] => {
    const startIndex = blocks.findIndex((block) => block.id === first.paragraphId);
    const endIndex = blocks.findIndex((block) => block.id === last.paragraphId);
    if (startIndex >= 0 && endIndex >= startIndex) {
      replaced = true;
      const replacement = lines.map((line, index): DocParagraph => ({
        ...makeParagraph([], rich?.[index] ?? start),
        id: index === 0 ? start.id : makeParagraph().id,
        runs: [...(index === 0 ? before : []), ...insertedRuns(index), ...(index === lines.length - 1 ? after : [])],
      }));
      paragraphId = replacement.at(-1)!.id;
      offset = lines.length === 1 ? first.start + lines[0].length : lines.at(-1)!.length;
      return [...blocks.slice(0, startIndex), ...replacement, ...blocks.slice(endIndex + 1)];
    }
    return blocks.map((block) => block.type !== 'table' ? block : {
      ...block, rows: block.rows.map((row) => ({ ...row, cells: row.cells.map((cell) => ({ ...cell, blocks: visit(cell.blocks) })) })),
    });
  };
  let blocks = visit(doc.blocks);
  if (!replaced) {
    blocks = mapParagraphs(doc.blocks, (paragraph) => {
      const slice = selected.get(paragraph.id);
      if (!slice) return paragraph;
      const [prefix, suffix] = outsideRuns(paragraph, slice);
      return { ...paragraph, runs: [...prefix, ...(paragraph.id === start.id ? (rich ? rich.flatMap((part, index) => [...(index ? [{ text: '', style, lineBreak: true }] : []), ...part.runs]) : [{ text, style }]) : []), ...suffix] };
    });
    offset = first.start + text.length;
  }
  return { doc: { ...doc, blocks }, paragraphId, offset };
}
