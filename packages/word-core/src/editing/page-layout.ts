// 功能：区域内容与样式定义以同一主文档事务提交，保证一次撤销；时间：2026-11-11；作者：lq。
import type { StyleCatalog } from './styles';
/** 功能：页面和页眉页脚设置共用正文事务、历史与保存版本；时间：2026-11-10；作者：lq。 */
import { Extension, type Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import type { Node } from '@tiptap/pm/model';
import { readSections, applyPageNumType, applyTitlePg, type ParsedDocFull, type SectionSettings, type SectionInfo, type HeaderFooter, type SaveBlock, type SaveOptions } from '../upstream/packages/docx-engine/src/index';
import { hfVariantsFromParsed } from '../upstream/apps/docs/src/renderer/doc-state';
import { hfSaveOptions, resolveHf, type HfSectionState } from '../upstream/apps/docs/src/renderer/hf-sections';
// 功能：提交前复用分页测量，避免页眉页脚占满整张纸后才暴露排版错误；时间：2026-11-10；作者：lq。
import { sectionHfHeights } from '../upstream/apps/docs/src/renderer/hf-heights';
import { effectiveTopPx, effectiveBottomPx } from '../upstream/apps/docs/src/renderer/pagination-sections';
import { hfVariantReservedPx } from '../upstream/apps/docs/src/renderer/pagination-hf';

/** 功能：只存用户修改字段，保留源分节、栏和未知 XML；时间：2026-11-10；作者：lq。 */
export type WordPagePatch = Partial<Pick<SectionSettings, 'pageWidth' | 'pageHeight' | 'orientation' | 'marginTop' | 'marginRight' | 'marginBottom' | 'marginLeft' | 'headerDist' | 'footerDist'>> & { pageNumberStart?: number; titlePg?: boolean };
/** 功能：首页按节、奇偶页按文档保存，隐藏变体保留其内容以便再次启用；时间：2026-11-11；作者：lq。 */
export interface WordPageLayout { sections: Record<string, WordPagePatch>; edits: Record<string, HeaderFooter>; links: Record<string, true>; evenAndOddHeaders?: boolean; }
const EMPTY: WordPageLayout = { sections: {}, edits: {}, links: {} };
export const wordPageLayout = Extension.create({
  name: 'miniOfficePageLayout',
  addGlobalAttributes: () => [{ types: ['doc'], attributes: { miniOfficePageLayout: { default: null, rendered: false } } }],
});
/** 功能：只读取当前不可变文档快照的页面目录；时间：2026-11-10；作者：lq。 */
export function wordPageLayoutOf(doc: Node): WordPageLayout { return doc.attrs.miniOfficePageLayout ?? EMPTY; }
/** 功能：分节标识沿用源段落索引，设置不会覆盖其他节；时间：2026-11-10；作者：lq。 */
export function wordSections(doc: Node, parsed: ParsedDocFull): SectionInfo[] {
  const catalog = wordPageLayoutOf(doc);
  return readSections(parsed).map(section => {
    const patch = catalog.sections[section.lastBlockIndex];
    return patch ? { ...section, settings: { ...section.settings, ...patch }, titlePg: patch.titlePg ?? section.titlePg, pageNumberStart: patch.pageNumberStart ?? section.pageNumberStart } : section;
  });
}
/** 功能：分页、预览和保存共同采用原生页眉页脚继承规则；时间：2026-11-10；作者：lq。 */
export function wordHeaderFooterState(doc: Node, parsed: ParsedDocFull, sections = wordSections(doc, parsed), catalog = wordPageLayoutOf(doc)): HfSectionState {
  return { sections, edits: catalog.edits, links: catalog.links, parsed,
    docStrips: { header: parsed.headerText || parsed.headerParas?.length || parsed.headerHasPageNumber ? { text: parsed.headerText ?? '', paras: parsed.headerParas ?? undefined, pageNumber: parsed.headerHasPageNumber } : null,
      footer: parsed.footerText || parsed.footerParas?.length || parsed.footerHasPageNumber ? { text: parsed.footerText ?? '', paras: parsed.footerParas ?? undefined, pageNumber: parsed.footerHasPageNumber } : null,
      variants: hfVariantsFromParsed(parsed) } };
}
/** 功能：页面草稿一次提交，可与正文编辑交错撤销；拒绝只读、输入法和无效尺寸；时间：2026-11-10；作者：lq。 */
export function commitWordPageLayout(editor: Editor, parsed: ParsedDocFull, next: WordPageLayout, historyGroup?: number, styles?: StyleCatalog, closeGroup = false): boolean {
  if (!editor.isEditable || editor.view.composing) return false;
  const sections = readSections(parsed);
  for (const [id, patch] of Object.entries(next.sections)) {
    const section = sections.find(item => String(item.lastBlockIndex) === id);
    if (!section) throw new Error('页面设置所属分节已不存在');
    const s = { ...section.settings, ...patch };
    for (const [key, value] of Object.entries(patch)) {
      if (key === 'titlePg') { if (typeof value !== 'boolean') throw new Error('无效的首页设置'); continue; }
      if (key === 'orientation') { if (value !== 'portrait' && value !== 'landscape') throw new Error('无效的纸张方向'); continue; }
      if (!Number.isInteger(value) || Number(value) < 0 || Number(value) > (key === 'pageNumberStart' ? 32767 : 31680)) throw new Error('页面设置超出有效范围');
    }
    if ((s.gutter ?? 0) > (s.gutterAtTop ? s.marginTop : s.marginLeft)) throw new Error('页边距不能小于原有装订线宽度');
    if (s.pageWidth - s.marginLeft - s.marginRight < 720 || s.pageHeight - s.marginTop - s.marginBottom < 720) throw new Error('页边距过大，请至少为正文保留 12.7 mm 空间');
    if ((s.headerDist ?? 0) >= s.pageHeight || (s.footerDist ?? 0) >= s.pageHeight) throw new Error('页眉页脚距纸边距离超出纸张高度');
  }
  if (JSON.stringify(wordPageLayoutOf(editor.state.doc)) === JSON.stringify(next) && (!styles || JSON.stringify(editor.state.doc.attrs.miniOfficeStyles) === JSON.stringify(styles))) return false;
  if (next.evenAndOddHeaders !== undefined && typeof next.evenAndOddHeaders !== 'boolean') throw new Error('无效的奇偶页设置');
  const pendingSections = sections.map(section => ({ ...section, titlePg: next.sections[section.lastBlockIndex]?.titlePg ?? section.titlePg, settings: { ...section.settings, ...next.sections[section.lastBlockIndex] } }));
  const hfState = wordHeaderFooterState(editor.state.doc, parsed, pendingSections, next);
  const evenOdd = next.evenAndOddHeaders ?? !!parsed.evenAndOddHeaders;
  const heights = sectionHfHeights(pendingSections, (index, kind, variant) => resolveHf(hfState, index, kind, variant), { titlePg: !!pendingSections[0]?.titlePg, evenOddHf: evenOdd });
  pendingSections.forEach((section, index) => {
    const variants: Array<'default' | 'first' | 'even'> = ['default', ...(section.titlePg ? ['first' as const] : []), ...(evenOdd ? ['even' as const] : [])];
    for (const variant of variants) {
      const reserved = hfVariantReservedPx(heights[index], variant);
      if (effectiveTopPx(section.settings, reserved.headerPx) + effectiveBottomPx(section.settings, reserved.footerPx) > section.settings.pageHeight / 15 - 48) throw new Error('页眉页脚占用空间过大，请减少文字、字号或距纸边距离');
    }
  });
  // 功能：DocAttrStep 没有相邻范围，区域显式边界必须关闭历史组后再设置组合身份；时间：2026-11-11；作者：lq。
  const transaction = historyGroup === undefined || closeGroup ? closeHistory(editor.state.tr) : editor.state.tr;
  if (historyGroup !== undefined) transaction.setMeta('composition', historyGroup);
  transaction.setDocAttribute('miniOfficePageLayout', structuredClone(next));
  if (styles) transaction.setDocAttribute('miniOfficeStyles', structuredClone(styles));
  editor.view.dispatch(transaction);
  return true;
}
/** 功能：只补写页面尺寸、边距和页码字段，保留分节未知属性及子元素顺序；时间：2026-11-10；作者：lq。 */
export function patchWordSection(xml: string, patch: WordPagePatch, source?: SectionSettings): string {
  let result = xml || '<w:sectPr/>';
  if (/<w:sectPr\b[^>]*\/>$/.test(result)) result = result.replace(/\/>$/, '></w:sectPr>');
  const groups = { pgSz: { pageWidth: 'w', pageHeight: 'h', orientation: 'orient' }, pgMar: { marginTop: 'top', marginRight: 'right', marginBottom: 'bottom', marginLeft: 'left', headerDist: 'header', footerDist: 'footer' } };
  for (const [name, fields] of Object.entries(groups)) {
    const updates = Object.entries(fields).filter(([key]) => key in patch);
    if (!updates.length) continue;
    const expression = new RegExp(`<w:${name}\\b[^>]*\\/>`);
    let tag = expression.exec(result)?.[0] ?? `<w:${name}/>`;
    for (const [key, attribute] of updates) {
      let value = patch[key as keyof WordPagePatch];
      // 功能：解析边距含装订线，保存需还原原始值；负边距的固定正文语义继续保留；时间：2026-11-10；作者：lq。
      if (source && typeof value === 'number') {
        if (key === 'marginLeft') value = Math.max(0, value - (source.gutterAtTop ? 0 : source.gutter ?? 0));
        if (key === 'marginTop') value = Math.max(0, value - (source.gutterAtTop ? source.gutter ?? 0 : 0)) * (source.marginTopFixed ? -1 : 1);
        if (key === 'marginBottom' && source.marginBottomFixed) value = -value;
      }
      const attr = new RegExp(`\\s+w:${attribute}\\s*=\\s*(["']).*?\\1`);
      tag = tag.replace(attr, '').replace(/\/>$/, ` w:${attribute}="${value}"/>`);
    }
    if (expression.test(result)) result = result.replace(expression, () => tag);
    else {
      // 功能：CT_SectPr 中引用在前，页面尺寸在边距前，不能把新标签随意置顶；时间：2026-11-10；作者：lq。
      const follower = new RegExp(`<w:${name === 'pgSz' ? '(?:pgMar|paperSrc|pgBorders|lnNumType|pgNumType|cols|formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)' : '(?:paperSrc|pgBorders|lnNumType|pgNumType|cols|formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)'}\\b`);
      const at = follower.exec(result)?.index ?? result.indexOf('</w:sectPr>');
      result = result.slice(0, at) + tag + result.slice(at);
    }
  }
  if (patch.pageNumberStart !== undefined) {
    const fmt = /<w:pgNumType\b[^>]*\bw:fmt=["']([^"']*)/.exec(result)?.[1];
    result = applyPageNumType(result, fmt, patch.pageNumberStart);
  }
  if (patch.titlePg !== undefined) result = applyTitlePg(result, patch.titlePg);
  return result;
}
/** 功能：将同一快照的页面目录应用到实际保存计划，正文改写后的分节属性也能保留；时间：2026-11-10；作者：lq。 */
export function pageLayoutSavePlan(parsed: ParsedDocFull, catalog: WordPageLayout, blocks: SaveBlock[]): { blocks: SaveBlock[]; options: SaveOptions } {
  const sections = readSections(parsed);
  const last = sections.at(-1)!;
  const options: SaveOptions = hfSaveOptions({ sections, edits: catalog.edits, links: catalog.links });
  if (catalog.evenAndOddHeaders !== undefined) options.evenAndOddHeaders = catalog.evenAndOddHeaders;
  if (catalog.sections[last.lastBlockIndex]) options.trailingSectPr = patchWordSection(last.sectPrXml, catalog.sections[last.lastBlockIndex], last.settings);
  return { options, blocks: blocks.map(block => {
    if (!('docxIndex' in block) || block.docxIndex == null || block.docxIndex === last.lastBlockIndex) return block;
    const patch = catalog.sections[block.docxIndex];
    if (!patch) return block;
    const replace = (xml: string) => xml.replace(/<w:sectPr\b[^>]*\/>|<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/, value => patchWordSection(value, patch, sections.find(item => item.lastBlockIndex === block.docxIndex)?.settings));
    if (block.kind === 'original') return { ...block, kind: 'xml' as const, xml: replace(parsed.blocks.find(item => item.docxIndex === block.docxIndex)!.originalXml ?? '') };
    if (block.kind === 'xml') return { ...block, xml: replace(block.xml) };
    if (block.kind === 'generated' && 'rawPPr' in block.block && typeof block.block.rawPPr === 'string') return { ...block, block: { ...block.block, rawPPr: replace(block.block.rawPPr) } };
    throw new Error('无法定位分节属性，页面设置尚未保存');
  }) };
}
