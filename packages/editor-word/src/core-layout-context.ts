/** 功能：承接固定 GenOffice 的页眉页脚和注释度量，让分页容量与显示使用同一份数据；时间：2026-11-09；作者：lq。 */
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import { readPageColor, type NoteInfo, type SectionInfo } from '../../word-core/src/index';
import { resolveHf, type HfKind } from '../../word-core/src/upstream/apps/docs/src/renderer/hf-sections';
import { sectionHfHeights } from '../../word-core/src/upstream/apps/docs/src/renderer/hf-heights';
import { hfLayoutResolved, hfPhantomSpec } from '../../word-core/src/upstream/apps/docs/src/renderer/hf-phantom';
import { effectiveTopPx, effectiveBottomPx, sectionPageBox } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-sections';
import { hfVariantOf, hfVariantReservedPx, type HfVariant } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-hf';
import { hfStripGeom, makeGapHfEl } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/hf-dom';
import { measureNoteHeightDom, noteRowMetrics, makeGapNotesEl } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/note-dom';
import { resolveNoteStyle, noteRunStyle, estimateFootnoteHeight, FOOTNOTE_SEPARATOR_H } from '../../word-core/src/upstream/apps/docs/src/renderer/line-metrics';
import { paraPaginationMeta, imageParaPaginationMeta } from '../../word-core/src/upstream/apps/docs/src/renderer/editor/para-flags';
import { pageOfY, type BlockBox, type PageSlice, type PageNoteItem, type BlockMeta } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination';

/** 功能：按当前事务构建注释引用，不把已删除的脚注继续留在页面底部；时间：2026-11-09；作者：lq。 */
export function wordLayoutContext(session: WordCoreSession, sections: SectionInfo[] = session.sections()) {
  const parsed = session.parsed;
  // 功能：读取事务中的页眉页脚草稿，撤销及重做与正文分页同步；时间：2026-11-10；作者：lq。
  const hfState = { ...session.headerFooterState(), sections };
  const hf = (section: number, kind: HfKind, variant: HfVariant) => hfLayoutResolved(kind, resolveHf(hfState, section, kind, variant), readPageColor(parsed) ? hfPhantomSpec(parsed) : null);
  const heights = sectionHfHeights(sections, hf, { titlePg: !!sections[0]?.titlePg, evenOddHf: !!session.evenAndOddHeaders() });
  const firstVariant = hfVariantOf(sections[0]?.titlePg, true, !!session.evenAndOddHeaders(), sections[0]?.pageNumberStart ?? 1);
  const firstReserved = hfVariantReservedPx(heights[0], firstVariant);
  const top = effectiveTopPx(sections[0].settings, firstReserved.headerPx);
  const bottom = effectiveBottomPx(sections[0].settings, heights[0].footerPx);
  const references = new Map<number, Array<{ kind: 'footnote' | 'endnote'; id: string; no: number }>>();
  const endReferences: Array<{ id: string; no: number }> = [];
  session.editor.state.doc.forEach(node => {
    const refs: Array<{ kind: 'footnote' | 'endnote'; id: string; no: number }> = [];
    node.descendants(child => {
      if (child.type.name !== 'docNoteRef') return;
      const ref = { kind: child.attrs.kind as 'footnote' | 'endnote', id: String(child.attrs.id), no: Number(child.attrs.num) || 1 };
      refs.push(ref);
      if (ref.kind === 'endnote') endReferences.push(ref);
    });
    if (typeof node.attrs.docxIndex === 'number') references.set(node.attrs.docxIndex, refs);
  });
  const noteCache = new Map<string, PageNoteItem>();
  /** 功能：复用上游样式解析和真实 DOM 高度探测，测量与页底内容保持一致；时间：2026-11-09；作者：lq。 */
  const note = (kind: 'footnote' | 'endnote', id: string, no: number, section: number): PageNoteItem => {
    const key = `${kind}:${id}:${no}:${section}`;
    const cached = noteCache.get(key);
    if (cached) return cached;
    const definition: NoteInfo | undefined = (kind === 'footnote' ? parsed.footnotes : parsed.endnotes).find(note => note.id === id);
    const settings = sections[section]?.settings ?? sections[0].settings;
    const style = { ...resolveNoteStyle(parsed, definition?.styleId, definition?.spacing), ...noteRunStyle(definition?.richParas) };
    const metrics = noteRowMetrics(style, settings.docGrid);
    const width = sectionPageBox(settings).contentWidth;
    const item = { id, no, text: definition?.text ?? '', richParas: definition?.richParas, ...(definition?.noRefMark ? { noRefMark: true as const } : {}), ...metrics };
    const height = measureNoteHeightDom(item, kind, width, metrics.lineHeightPx, metrics.fontSizePt, metrics.fontFamily)
      ?? estimateFootnoteHeight(item.text, width, settings.docGrid, undefined, style, item.richParas);
    const value = { ...item, height };
    noteCache.set(key, value);
    return value;
  };
  /** 功能：把段落约束和脚注预留交给同源切片器，而不是显示后覆盖正文；时间：2026-11-09；作者：lq。 */
  const blockMetaOf = (docxIndex: number): BlockMeta | undefined => {
    const block = parsed.blocks.find(block => block.docxIndex === docxIndex);
    const section = Math.max(0, sections.findIndex(section => docxIndex <= section.lastBlockIndex));
    const flags = block?.format ? paraPaginationMeta(block.format) : imageParaPaginationMeta(block?.originalXml, id => parsed.styles.get(id ?? '')?.display, parsed.docDefaults);
    const bands = (references.get(docxIndex) ?? []).filter(ref => ref.kind === 'footnote').map(ref => ({ heightPx: note('footnote', ref.id, ref.no, section).height }));
    return { ...flags, ...(bands.length ? { footnoteBands: bands, footnoteExtraPx: bands.reduce((sum, band) => sum + band.heightPx, 0) } : {}) };
  };
  /** 功能：脚注跟随实际引用行所在页，跨段编辑后页底重新归属；时间：2026-11-09；作者：lq。 */
  const pageNotes = (blocks: BlockBox[], slices: PageSlice[]): PageNoteItem[][] => {
    const pages: PageNoteItem[][] = slices.map(() => []);
    for (const block of blocks) {
      const refs = (block.docxIndex === undefined ? [] : references.get(block.docxIndex) ?? []).filter(ref => ref.kind === 'footnote');
      refs.forEach((ref, index) => {
        const y = block.top + (block.spaceBeforePx ?? 0) + (block.noteBands?.[index]?.offset ?? 0);
        const page = pageOfY(slices, y);
        const item = note('footnote', ref.id, ref.no, slices[page]?.section ?? 0);
        if (pages[page] && !pages[page].some(existing => existing.id === item.id)) pages[page].push(item);
      });
    }
    return pages;
  };
  /** 功能：页眉页脚变体和页码由解析模型决定，复用上游富内容 DOM 构造器；时间：2026-11-09；作者：lq。 */
  const strip = (section: number, kind: HfKind, variant: HfVariant, pageNo: string, total: number, sectionPages: number) => {
    const resolved = hf(section, kind, variant);
    if (!resolved.value && !resolved.images?.length) return undefined;
    return makeGapHfEl({ kind, value: resolved.value ?? { text: '' }, images: resolved.images, pageNo, pageTotal: total, sectionPages, geom: hfStripGeom(sections[section].settings) });
  };
  /** 功能：页底注释与末页注释区域使用相同渲染器；M3 尚未开放注释编辑入口；时间：2026-11-09；作者：lq。 */
  const notesElement = (items: PageNoteItem[], section: number, kind: 'footnote' | 'endnote' = 'footnote') => {
    if (!items.length) return undefined;
    const settings = sections[section].settings;
    const height = items.reduce((sum, item) => sum + item.height, 0) + FOOTNOTE_SEPARATOR_H;
    const element = makeGapNotesEl(items, settings.marginLeft / 15, sectionPageBox(settings).contentWidth, height, items[0].lineHeightPx ?? 16, () => {}, kind);
    element.contentEditable = 'false';
    return { element, height };
  };
  return { sections, heights, top, bottom, hf, strip, blockMetaOf, pageNotes, notesElement,
    endnotes: endReferences.map(ref => note('endnote', ref.id, ref.no, sections.length - 1)) };
}
export type WordLayoutContext = ReturnType<typeof wordLayoutContext>;
