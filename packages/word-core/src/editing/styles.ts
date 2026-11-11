/** 功能：将样式定义放入文档原生属性，创建、改名、删除与正文共用撤销及快照；时间：2026-11-10；作者：lq。 */
import { Extension, type Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import type { Node } from '@tiptap/pm/model';
import JSZip from 'jszip';
import { PPR_CHILD_ORDER, RPR_CHILD_ORDER, splitXmlChildren, inlineRunsXml } from '../upstream/packages/docx-engine/src/generate';
import { inlineToRuns } from '../upstream/apps/docs/src/renderer/editor/convert';
import type { ParsedDocFull, StyleInfo, StyleDisplay, StyleUpsert } from '../upstream/packages/docx-engine/src/index';
import { appendWordFormat, type WordFormatSnapshot } from './format-painter';
import { selectedParagraphs } from './formatting';
import { wordXmlEscape, xmlSpans, wordXmlNode, wordXmlAttr, applyXmlEdits, type XmlEdit } from '../serialization/xml-spans';

/** 功能：保存原生定义及用于新建样式的完整格式，改名不重建导入样式属性；时间：2026-11-10；作者：lq。 */
export interface WordStyleDefinition { upsert: StyleUpsert; format?: WordFormatSnapshot; }
export interface StyleCatalog { definitions: WordStyleDefinition[]; deleted: string[]; retained?: Record<string, string>; }
const EMPTY: StyleCatalog = { definitions: [], deleted: [] };

/** 功能：文档属性由 ProseMirror 的 DocAttrStep 记录和反演，不另外维护样式历史；时间：2026-11-10；作者：lq。 */
export const wordStyleCatalog = Extension.create({
  name: 'miniOfficeStyleCatalog',
  addGlobalAttributes: () => [{ types: ['doc'], attributes: { miniOfficeStyles: { default: null, rendered: false } } }],
});

/** 功能：只读取不可变事务快照，返回的样式目录不由界面直接修改；时间：2026-11-10；作者：lq。 */
export function wordStyleCatalogOf(doc: Node): StyleCatalog { return doc.attrs.miniOfficeStyles ?? EMPTY; }

/** 功能：把本地完整格式映射为上游样式显示数据；时间：2026-11-10；作者：lq。 */
function displayOf(format: WordFormatSnapshot): StyleDisplay {
  const { text: t, paragraph: p } = format;
  return {
    font: t.font, eastAsiaFont: t.font, fontAscii: t.fontAscii, sizeHalfPoints: t.sizeHalfPoints,
    bold: t.bold, italic: t.italic, underline: t.underline, strike: t.strike, color: t.color?.replace(/^#/, '') ?? 'auto', shading: t.highlight?.replace(/^#/, ''),
    align: p.align as StyleDisplay['align'], indentLeftTwips: p.indentLeft ?? undefined, indentRightTwips: p.indentRight ?? undefined,
    indentFirstLineTwips: p.indentFirstLine ?? undefined, spaceBeforeTwips: p.spaceBefore ?? undefined, spaceAfterTwips: p.spaceAfter ?? undefined,
    lineRule: p.lineRule as StyleDisplay['lineRule'], lineRawTwips: p.lineRawTwips ?? undefined, lineSpacing: p.lineSpacing ?? undefined,
    keepNext: p.keepNext ?? undefined, keepLines: p.keepLines ?? undefined, pageBreakBefore: p.pageBreakBefore,
    widowControl: p.widowControl ?? undefined,
    borderSides: p.borderLines ? JSON.parse(p.borderLines) : undefined, shadingFill: p.shadingFill ?? undefined,
  };
}

/** 功能：导入样式优先保留原样，事务中的新增/改名覆盖目录元数据；时间：2026-11-10；作者：lq。 */
export function wordStyles(doc: Node, parsed: ParsedDocFull): Map<string, StyleInfo> {
  const result = new Map(parsed.styles);
  const catalog = wordStyleCatalogOf(doc);
  for (const id of catalog.deleted) result.delete(id);
  // 功能：删除导入样式保留隐藏的格式基底，页眉/脚注/派生样式及未知效果仍能继承；时间：2026-11-11；作者：lq。
  for (const [id, alias] of Object.entries(catalog.retained ?? {})) { const original = parsed.styles.get(id); if (original) result.set(alias, { ...original, styleId: alias, semiHidden: true, qFormat: false, unhideWhenUsed: false }); }
  for (const { upsert: up, format } of catalog.definitions) {
    const original = result.get(up.styleId);
    result.set(up.styleId, {
      ...original, styleId: up.styleId, name: up.name ?? original?.name ?? up.styleId,
      type: original?.type ?? up.type ?? 'paragraph', custom: original?.custom ?? !up.builtIn,
      basedOn: up.basedOn === null ? undefined : up.basedOn ?? original?.basedOn,
      headingLevel: up.pPr?.outlineLevel === null ? undefined : up.pPr?.outlineLevel ?? original?.headingLevel,
      numPr: up.pPr?.numPr === null ? 'none' : up.pPr?.numPr ?? original?.numPr,
      display: format ? displayOf(format) : original?.display,
    });
  }
  return result;
}

/** 功能：将现有样式的已解析格式交给摘要和事务；未知属性继续经原样式引用继承；时间：2026-11-10；作者：lq。 */
export function wordStyleFormat(style: StyleInfo, parsed: ParsedDocFull): WordFormatSnapshot {
  const d = { ...parsed.docDefaults, ...style.display };
  return {
    text: { font: d.font ?? d.eastAsiaFont, fontAscii: d.fontAscii ?? parsed.docDefaults?.asciiFont,
      sizeHalfPoints: d.sizeHalfPoints ?? 22, bold: d.bold ?? false, italic: d.italic ?? false,
      underline: d.underline ?? false, strike: d.strike ?? false, color: d.color && d.color !== 'auto' ? `#${d.color}` : undefined,
      highlight: d.shading ? `#${d.shading}` : undefined },
    paragraph: { align: d.align ?? 'left', indentLeft: d.indentLeftTwips ?? 0, indentRight: d.indentRightTwips ?? 0,
      indentFirstLine: d.indentFirstLineTwips ?? 0, spaceBefore: d.spaceBeforeTwips ?? 0, spaceAfter: d.spaceAfterTwips ?? 0,
      lineRule: d.lineRule ?? 'auto', lineSpacing: (d.lineRule ?? 'auto') === 'auto' ? d.lineSpacing ?? 1 : null,
      lineRawTwips: d.lineRawTwips ?? null, keepNext: d.keepNext ?? false, keepLines: d.keepLines ?? false,
      pageBreakBefore: d.pageBreakBefore ?? false, widowControl: d.widowControl ?? true,
      borders: d.borderSides ? Object.entries(d.borderSides).filter(([, line]) => line && !line.none).map(([side]) => side).join('') : '',
      borderLines: d.borderSides ? JSON.stringify(d.borderSides) : null, shadingFill: d.shadingFill ?? null,
      tabStops: d.tabStops ? JSON.stringify(d.tabStops) : null },
    effects: { charSpacingTwips: d.charSpacingTwips ?? null, caps: d.caps ?? null, bdr: d.bdr ? JSON.stringify(d.bdr) : null, textOutline: d.textOutline ? JSON.stringify(d.textOutline) : null },
    ...(style.numPr && style.numPr !== 'none' ? { list: { numId: style.numPr.numId, level: style.numPr.ilvl, kind: parsed.numbering.get(style.numPr.numId)?.levels[style.numPr.ilvl]?.numFmt === 'bullet' ? 'bullet' as const : 'ordered' as const } } : {}),
  };
}

/** 功能：完整格式生成属性补丁，额外固定行距与底纹由局部 XML 补写持久化；时间：2026-11-10；作者：lq。 */
export function defineWordStyle(id: string, name: string, format: WordFormatSnapshot, headingLevel?: number): WordStyleDefinition {
  const { text: t, paragraph: p } = format;
  return { format: structuredClone(format), upsert: { styleId: id, name, type: 'paragraph', basedOn: null, next: 'Normal',
    pPr: { align: p.align as 'left', indentLeftTwips: p.indentLeft, indentRightTwips: p.indentRight,
      firstLineTwips: p.indentFirstLine, spaceBeforeTwips: p.spaceBefore, spaceAfterTwips: p.spaceAfter,
      lineSpacing: p.lineRule === 'auto' ? p.lineSpacing : null, keepNext: p.keepNext ?? false, keepLines: p.keepLines ?? false,
      pageBreakBefore: p.pageBreakBefore ?? false, outlineLevel: headingLevel ?? null,
      ...(format.list?.numId ? { numPr: { numId: format.list.numId, ilvl: format.list.level } } : {}) },
    rPr: { font: t.fontAscii ?? t.font ?? null, eastAsiaFont: t.font ?? null, sizeHalfPoints: t.sizeHalfPoints,
      bold: t.bold ?? false, italic: t.italic ?? false, underline: t.underline ?? false, strike: t.strike ?? false,
      color: t.color?.replace(/^#/, '') ?? 'auto' },
  } };
}

/** 功能：定义与段落应用合为一个事务，空正文创建或仅改名也会进入原生撤销；时间：2026-11-10；作者：lq。 */
export function applyWordStyle(editor: Editor, parsed: ParsedDocFull, id: string, definition?: WordStyleDefinition, apply = true): boolean {
  if (!editor.isEditable || editor.view.composing || !editor.schema.nodes.docParagraph) return false;
  const catalog = wordStyleCatalogOf(editor.state.doc);
  const next = definition ? { ...catalog, definitions: [...catalog.definitions.filter(item => item.upsert.styleId !== id), structuredClone(definition)], deleted: catalog.deleted.filter(value => value !== id) } : catalog;
  const tr = closeHistory(editor.state.tr);
  if (definition) tr.setDocAttribute('miniOfficeStyles', next);
  const style = wordStyles(tr.doc, parsed).get(id);
  if (!style || style.type !== 'paragraph') return false;
  if (apply) {
    const format = next.definitions.find(item => item.upsert.styleId === id)?.format ?? wordStyleFormat(style, parsed);
    appendWordFormat(tr, editor, format, true);
    for (const { node, pos } of selectedParagraphs(editor)) {
      // 功能：表格单元格使用其允许的段落类型，列表保留编号语义，正文标题正确切换节点；时间：2026-11-10；作者：lq。
      const parent = editor.state.doc.resolve(pos).parent;
      const current = tr.doc.nodeAt(pos)!;
      const type = current.type.name === 'docListItem' ? current.type : style.headingLevel
        ? (editor.schema.nodes.docHeading ?? editor.schema.nodes.docParagraph) : editor.schema.nodes.docParagraph;
      const index = editor.state.doc.resolve(pos).index();
      const allowed = parent.canReplaceWith(index, index + 1, type) ? type : editor.schema.nodes.docParagraph;
      tr.setNodeMarkup(pos, allowed, { ...current.attrs, styleId: id, level: style.headingLevel ?? null });
    }
  }
  if (!tr.docChanged) return false;
  editor.view.dispatch(tr);
  return true;
}

/** 功能：一次导入整组定义，校验完成后原子提交并与前后正文操作分隔撤销；时间：2026-11-10；作者：lq。 */
export function importWordStyles(editor: Editor, definitions: WordStyleDefinition[]): boolean {
  if (!editor.isEditable || editor.view.composing || !editor.schema.nodes.docParagraph || !definitions.length) return false;
  const catalog = wordStyleCatalogOf(editor.state.doc);
  editor.view.dispatch(closeHistory(editor.state.tr).setDocAttribute('miniOfficeStyles', {
    ...catalog, definitions: [...catalog.definitions, ...structuredClone(definitions)],
  }));
  return true;
}

/** 功能：仅改样式名称，不重新生成导入样式的格式 XML；时间：2026-11-10；作者：lq。 */
export function renameWordStyle(editor: Editor, parsed: ParsedDocFull, id: string, name: string): boolean {
  const trimmed = name.trim();
  if (!trimmed || trimmed.length > 40 || [...wordStyles(editor.state.doc, parsed).values()].some(item => item.styleId !== id && item.name === trimmed)) return false;
  const original = wordStyleCatalogOf(editor.state.doc).definitions.find(item => item.upsert.styleId === id);
  return applyWordStyle(editor, parsed, id, { ...original, upsert: { ...original?.upsert, styleId: id, name: trimmed } }, false);
}

/** 功能：删除用户可见样式时保留实际格式，导入定义的跨部件引用改向隐藏基底；时间：2026-11-11；作者：lq。 */
export function deleteWordStyle(editor: Editor, parsed: ParsedDocFull, id: string): boolean {
  if (!editor.isEditable || editor.view.composing || id === 'Normal' || parsed.styles.get(id)?.isDefault) return false;
  const catalog = wordStyleCatalogOf(editor.state.doc);
  if (!wordStyles(editor.state.doc, parsed).has(id)) return false;
  const retained = { ...catalog.retained };
  if (parsed.styles.has(id) && !retained[id]) {
    const stem = 'MoRetained_' + Array.from(id).map(char => char.codePointAt(0)!.toString(16)).join('_');
    const taken = wordStyles(editor.state.doc, parsed); let alias = stem, suffix = 1;
    while (taken.has(alias)) alias = stem + '_' + suffix++;
    retained[id] = alias;
  }
  const replacement = retained[id] ?? null;
  const tr = closeHistory(editor.state.tr).setDocAttribute('miniOfficeStyles', {
    ...catalog, retained, definitions: catalog.definitions.filter(item => item.upsert.styleId !== id), deleted: [...new Set([...catalog.deleted, id])],
  });
  // 功能：只改已建模的 styleId，不匹配普通文字/未知值；附属区域仍属于同一文档事务；时间：2026-11-11；作者：lq。
  const rewrite = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(rewrite);
    if (!value || typeof value !== 'object') return value;
    return Object.fromEntries(Object.entries(value).map(([key, child]) => [key, key === 'styleId' && child === id ? replacement : rewrite(child)]));
  };
  editor.state.doc.descendants((node, pos) => {
    if (node.isTextblock && node.attrs.styleId === id) tr.setNodeMarkup(pos, undefined, { ...node.attrs, styleId: replacement });
    if (node.attrs.textboxes) tr.setNodeMarkup(pos, undefined, { ...node.attrs, textboxes: rewrite(node.attrs.textboxes) });
  });
  if (editor.state.doc.attrs.miniOfficePageLayout) tr.setDocAttribute('miniOfficePageLayout', rewrite(editor.state.doc.attrs.miniOfficePageLayout));
  editor.view.dispatch(tr); return true;
}

/** 功能：仅将明确的 Word 样式引用重定向，所有其他 XML 子串保持不变；时间：2026-11-11；作者：lq。 */
async function retainDeletedWordStyles(zip: JSZip, catalog: StyleCatalog): Promise<void> {
  const retained = catalog.retained ?? {}; if (!Object.keys(retained).length) return;
  const referenceTags = ['pStyle', 'rStyle', 'tblStyle', 'basedOn', 'next', 'link', 'styleLink', 'numStyleLink', 'style'];
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir || !path.endsWith('.xml')) continue;
    const original = await file.async('string'); let xml = original;
    if (path === 'word/styles.xml') {
      const clones: string[] = [];
      for (const node of xmlSpans(xml)) {
        if (!wordXmlNode(node, 'style')) continue;
        const id = wordXmlAttr(node, 'styleId')!; const alias = retained[id]; if (!alias) continue;
        const prefix = node.prefix || 'w';
        let clone = xml.slice(node.start, node.end).replace(/\bw:styleId=(['"])[^'"]*\1/, `w:styleId="${wordXmlEscape(alias)}"`);
        clone = clone.replace(/<w:(?:qFormat|semiHidden|unhideWhenUsed)\b[^>]*(?:\/>|>[\s\S]*?<\/w:(?:qFormat|semiHidden|unhideWhenUsed)>)/g, '');
        const follower = new RegExp(`<${prefix}:(?:unhideWhenUsed|qFormat|locked|personal|personalCompose|personalReply|rsid|pPr|rPr|tblPr|trPr|tcPr|tblStylePr)\\b`).exec(clone)?.index ?? clone.indexOf(`</${prefix}:style>`);
        clone = clone.slice(0, follower) + `<${prefix}:semiHidden/>` + clone.slice(follower);
        clones.push(clone);
      }
      const nodes = xmlSpans(xml); const remove = nodes.filter(node => wordXmlNode(node, 'style') && catalog.deleted.includes(wordXmlAttr(node, 'styleId') ?? ''));
      xml = applyXmlEdits(xml, remove.map(node => ({ start: node.start, end: node.end, text: '' })));
      xml = xml.replace('</w:styles>', clones.join('') + '</w:styles>');
    }
    const edits: XmlEdit[] = [];
    for (const node of xmlSpans(xml)) {
      if (!referenceTags.some(tag => wordXmlNode(node, tag))) continue;
      const id = wordXmlAttr(node, 'val') ?? ''; const alias = catalog.deleted.includes(id) ? retained[id] : undefined; if (!alias) continue;
      const tag = xml.slice(node.start, node.openEnd); const attr = new RegExp(`(\\b${node.prefix || 'w'}:val\\s*=\\s*)(['"])[^'"]*\\2`);
      edits.push({ start: node.start, end: node.openEnd, text: tag.replace(attr, (_all, key: string) => key + '"' + wordXmlEscape(alias) + '"') });
    }
    xml = applyXmlEdits(xml, edits); if (xml !== original) zip.file(path, xml);
  }
}

/** 功能：仅补写本地新建定义里上游补丁尚未覆盖的行距、底纹和孤行控制，原样式原字节保留；时间：2026-11-10；作者：lq。 */
export async function finishWordStyles(bytes: Uint8Array, catalog: StyleCatalog): Promise<Uint8Array> {
  const extras = catalog.definitions.filter(item => item.format);
  if (!extras.length && !Object.keys(catalog.retained ?? {}).length) return bytes;
  const zip = await JSZip.loadAsync(bytes);
  let xml = await zip.file('word/styles.xml')!.async('string');
  for (const { upsert, format } of extras) {
    const { text: t, paragraph: p } = format!;
    xml = xml.replace(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g, style => {
      if (!style.includes(`w:styleId="${upsert.styleId}"`)) return style;
      let result = style;
      // 功能：边框、底纹和制表位保存样式原生定义，效果通过同源 Run 转换写回；时间：2026-11-11；作者：lq。
      const setChild = (parent: string, tag: string, content: string) => {
        const child = new RegExp(`<w:${tag}\\b[^>]*(?:/>|>[\\s\\S]*?</w:${tag}>)`, 'g');
        result = result.replace(new RegExp(`<w:${parent}(?:\\s[^>]*)?>([\\s\\S]*?)</w:${parent}>`), (_all, inner: string) => `<w:${parent}>${inner.replace(child, '')}${content}</w:${parent}>`);
      };
      if (p.borderLines) {
        const lines = JSON.parse(p.borderLines);
        const border = Object.entries({ t: 'top', l: 'left', b: 'bottom', r: 'right' }).map(([side, name]) => {
          const line = lines[side]; if (!line) return '';
          return `<w:${name} w:val="${wordXmlEscape(line.none ? 'nil' : line.val ?? 'single')}" w:sz="${Math.round((line.szPt ?? 0.5) * 8)}" w:space="${wordXmlEscape(String(line.spacePt ?? 0))}" w:color="${wordXmlEscape(line.color ?? 'auto')}"/>`;
        }).join('');
        setChild('pPr', 'pBdr', `<w:pBdr>${border}</w:pBdr>`);
      }
      if (p.shadingFill) setChild('pPr', 'shd', `<w:shd w:val="clear" w:fill="${wordXmlEscape(p.shadingFill)}"/>`);
      if (p.tabStops) setChild('pPr', 'tabs', `<w:tabs>${JSON.parse(p.tabStops).map((tab: { pos: number; val: string; leader?: string }) => `<w:tab w:val="${wordXmlEscape(tab.val)}" w:pos="${wordXmlEscape(String(tab.pos))}"${tab.leader ? ` w:leader="${wordXmlEscape(tab.leader)}"` : ''}/>`).join('')}</w:tabs>`);
      if (format?.list?.numId) setChild('pPr', 'numPr', `<w:numPr><w:ilvl w:val="${format.list.level}"/><w:numId w:val="${wordXmlEscape(format.list.numId)}"/></w:numPr>`);
      if (format?.effects) {
        const runs = inlineToRuns([{ type: 'text', text: 'x', marks: [{ type: 'docTextStyle', attrs: format.effects }] }]);
        const rPr = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(inlineRunsXml(runs))?.[1] ?? '';
        for (const child of splitXmlChildren(rPr)) setChild('rPr', child.name.replace('w:', ''), child.xml);
      }
      if (p.lineRule && p.lineRule !== 'auto' && p.lineRawTwips != null) {
        result = result.replace(/<w:spacing\b[^>]*\/>/, spacing => spacing.replace(/\s+w:line(?:Rule)?="[^"]*"/g, '').replace('/>', ` w:line="${p.lineRawTwips}" w:lineRule="${p.lineRule}"/>`));
        if (!/<w:spacing\b/.test(result)) result = result.replace('</w:pPr>', `<w:spacing w:line="${p.lineRawTwips}" w:lineRule="${p.lineRule}"/></w:pPr>`);
      }
      if (p.widowControl != null) result = result.replace('</w:pPr>', `<w:widowControl w:val="${p.widowControl ? 1 : 0}"/></w:pPr>`);
      if (t.highlight && /^#[\da-f]{6}$/i.test(t.highlight)) result = result.replace('</w:rPr>', `<w:shd w:val="clear" w:fill="${t.highlight.slice(1)}"/></w:rPr>`);
      // 功能：补写节点后沿用同源 OOXML 子元素顺序，避免 Word 修复或丢弃样式属性；时间：2026-11-10；作者：lq。
      for (const [tag, order] of [['pPr', PPR_CHILD_ORDER], ['rPr', RPR_CHILD_ORDER]] as const) {
        result = result.replace(new RegExp(`<w:${tag}>([\\s\\S]*?)</w:${tag}>`), (_whole, inner: string) => {
          const rank = (name: string) => order.indexOf(name) < 0 ? order.length : order.indexOf(name);
          const children = splitXmlChildren(inner).sort((a, b) => rank(a.name) - rank(b.name));
          return `<w:${tag}>${children.map(item => item.xml).join('')}</w:${tag}>`;
        });
      }
      return result;
    });
  }
  zip.file('word/styles.xml', xml);
  await retainDeletedWordStyles(zip, catalog);
  return zip.generateAsync({ type: 'uint8array' });
}
