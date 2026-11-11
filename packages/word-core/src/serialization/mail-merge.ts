/** 功能：直接合并原生保存快照的 OOXML 部件，保持样式/图片/表格/分节与未知 XML，不经过旧 DocModel；时间：2026-11-11；作者：lq。 */
import type JSZip from 'jszip';
import { loadDocxZip } from '../upstream/packages/docx-engine/src/zip-load';
import { resolveRelationshipTargetPath } from '../upstream/packages/docx-engine/src/parse-package';
import type { MergeRecord } from '../upstream/packages/docx-engine/src/mail-merge';
import { inspectWordMergeXml, mergeWordXml } from './merge-fields';
import { applyXmlEdits, wordXmlAttr, wordXmlEscape, wordXmlNode, xmlSpans, type XmlEdit, type XmlSpan } from './xml-spans';

const WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
export interface NativeWordMergeOptions { isCancelled?: () => boolean; onProgress?: (current: number, total: number) => void; }
export interface WordMergePackageInfo { fields: string[]; auxiliaryFields: string[]; staticAuxiliaryFields: string[]; nextBlocks: number; panelState?: string; }
/** 功能：只读取真实 Word 内容部件的域和面板文档变量，不从字体或样式的同名字符串误判字段；时间：2026-11-11；作者：lq。 */
export async function inspectWordMergePackage(bytes: Uint8Array): Promise<WordMergePackageInfo> {
  const zip = await loadDocxZip(bytes); const fields = new Set<string>(), auxiliary = new Set<string>(), fixed = new Set<string>(); let nextBlocks = 0;
  for (const path of Object.keys(zip.files).filter(path => /^word\/(?:document|header[^/]*|footer[^/]*|footnotes|endnotes|comments)\.xml$/.test(path))) {
    const info = inspectWordMergeXml(await zip.file(path)!.async('string')); info.fields.forEach(name => fields.add(name));
    if (path !== 'word/document.xml') info.fields.forEach(name => auxiliary.add(name));
    // 功能：页眉页脚摘要由当前模型负责，未开放编辑的注释部件仍读取来源；时间：2026-11-11；作者：lq。
    if (/^word\/(?:footnotes|endnotes|comments)\.xml$/.test(path)) info.fields.forEach(name => fixed.add(name));
    if (path === 'word/document.xml' && info.next) nextBlocks = info.next + 1;
  }
  const settings = await zip.file('word/settings.xml')?.async('string');
  const state = settings ? xmlSpans(settings).find(node => wordXmlNode(node, 'docVar') && wordXmlAttr(node, 'name') === 'mailMergePanelState') : undefined;
  return { fields: [...fields], auxiliaryFields: [...auxiliary], staticAuxiliaryFields: [...fixed], nextBlocks, panelState: state ? wordXmlAttr(state, 'val') : undefined };
}
/** 功能：仅替换目标属性值，未修改的属性顺序、引号和扩展属性原样保留；时间：2026-11-11；作者：lq。 */
function replaceAttribute(tag: string, name: string, value: string): string {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return tag.replace(new RegExp(`(${escaped}\\s*=\\s*)(["'])([\\s\\S]*?)\\2`), (_match, before, quote) => `${before}${quote}${wordXmlEscape(value)}${quote}`);
}
/** 功能：关系 ID 和部件路径由现有包避让分配，重复输出不会覆盖模板资源；时间：2026-11-11；作者：lq。 */
function allocator(existing: Iterable<string>, prefix: string): () => string {
  const used = new Set(existing); let value = 0;
  return () => { let id: string; do { id = `${prefix}${++value}`; } while (used.has(id)); used.add(id); return id; };
}
/** 功能：每份记录的书签/注释/脚注/绘图/段落 ID 分别重分配，保留引用对应关系；时间：2026-11-11；作者：lq。 */
function recordIdentity(pass: number, number: () => number, numberIds: Map<string, string>, preservedNotes = new Set<string>(), usedNames = new Set<string>()) {
  const ids = new Map<string, string>(), names = new Map<string, string>();
  const id = (kind: string, old: string) => { const key = `${kind}:${old}`; if (!ids.has(key)) ids.set(key, String(number())); return ids.get(key)!; };
  const bookmark = (name: string) => { if (!names.has(name)) { let mapped: string; do { mapped = `_mo${pass}_${number()}_${name}`.slice(0, 40); } while (usedNames.has(mapped)); usedNames.add(mapped); names.set(name, mapped); } return names.get(name)!; };
  return (xml: string) => {
    const nodes = xmlSpans(xml); const edits: XmlEdit[] = [];
    for (const node of nodes) {
      let tag = xml.slice(node.start, node.openEnd); const before = tag;
      if (wordXmlNode(node, 'bookmarkStart') || wordXmlNode(node, 'bookmarkEnd')) {
        const old = wordXmlAttr(node, 'id'); if (old !== undefined) tag = replaceAttribute(tag, `${node.prefix}:id`, id('bookmark', old));
        const name = wordXmlAttr(node, 'name'); if (name) tag = replaceAttribute(tag, `${node.prefix}:name`, bookmark(name));
      } else if (wordXmlNode(node, 'hyperlink') && wordXmlAttr(node, 'anchor')) tag = replaceAttribute(tag, `${node.prefix}:anchor`, bookmark(wordXmlAttr(node, 'anchor')!));
      else if (/^(footnote|endnote|comment)(Reference|RangeStart|RangeEnd)?$/.test(node.local) && wordXmlNode(node, node.local)) {
        const old = wordXmlAttr(node, 'id'); if (old !== undefined && Number(old) >= (node.local.startsWith('comment') ? 0 : 1)) { const kind = node.local.replace(/Reference|RangeStart|RangeEnd/g, ''); if (!preservedNotes.has(`${kind}:${old}`)) tag = replaceAttribute(tag, `${node.prefix}:id`, id(kind, old)); }
      } else if (wordXmlNode(node, 'numId')) {
        const old = wordXmlAttr(node, 'val'); if (old && old !== '0') { if (!numberIds.has(old)) numberIds.set(old, String(number())); tag = replaceAttribute(tag, `${node.prefix}:val`, numberIds.get(old)!); }
      } else if (['ins', 'del', 'moveFrom', 'moveTo', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd', 'id'].some(name => wordXmlNode(node, name))) {
        const attribute = node.local === 'id' ? 'val' : 'id'; const old = wordXmlAttr(node, attribute); if (old !== undefined) tag = replaceAttribute(tag, `${node.prefix}:${attribute}`, id('revision', old));
      }
      if (['docPr', 'cNvPr'].includes(node.local) && node.attrs.id !== undefined) tag = replaceAttribute(tag, 'id', String(number()));
      for (const [attribute, old] of Object.entries(node.attrs)) if (/^(?:w14|w15|w16cid|w16cex|wp14):(?:paraId|paraIdParent|textId|anchorId|editId|durableId)$/.test(attribute) && old !== '00000000') { const kind = attribute.split(':')[1].replace('paraIdParent', 'paraId'); tag = replaceAttribute(tag, attribute, Number(id(kind, old.toUpperCase())).toString(16).padStart(8, '0').toUpperCase()); }
      if (tag !== before) edits.push({ start: node.start, end: node.openEnd, text: tag });
      if (wordXmlNode(node, 'instrText')) {
        const text = xml.slice(node.openEnd, node.closeStart);
        const updated = text.replace(/(\b(?:REF|PAGEREF|NOTEREF)\s+)(?:"([^"]+)"|([^\s<]+))/gi, (_match, lead, quoted, bare) => `${lead}${quoted ? '"' : ''}${bookmark(quoted ?? bare)}${quoted ? '"' : ''}`);
        if (updated !== text) edits.push({ start: node.openEnd, end: node.closeStart, text: updated });
      }
    }
    return applyXmlEdits(xml, edits);
  };
}
/** 功能：插入新部件类型和关系时沿用原根命名空间，其他包配置保持不变；时间：2026-11-11；作者：lq。 */
function appendRoot(xml: string, fragment: string): string {
  const root = xmlSpans(xml).find(node => !node.parent); if (!root) throw new Error('缺少 OOXML 根元素。');
  if (root.openEnd === root.end) return xml.slice(0, root.start) + xml.slice(root.start, root.end).replace(/\/\s*>$/, '>') + fragment + `</${root.name}>` + xml.slice(root.end);
  return xml.slice(0, root.closeStart) + fragment + xml.slice(root.closeStart);
}
/** 功能：输出清除模板邮件合并连接及私有面板状态，未知设置和其他外部关系继续保留；时间：2026-11-11；作者：lq。 */
async function clearMergeSettings(zip: JSZip): Promise<void> {
  const settings = await zip.file('word/settings.xml')?.async('string');
  if (settings) {
    const nodes = xmlSpans(settings); const removable = nodes.filter(node => wordXmlNode(node, 'mailMerge') || wordXmlNode(node, 'recipientData') || (wordXmlNode(node, 'docVar') && wordXmlAttr(node, 'name') === 'mailMergePanelState'));
    const roots = removable.filter(node => !removable.some(parent => parent !== node && parent.start <= node.start && parent.end >= node.end));
    zip.file('word/settings.xml', applyXmlEdits(settings, roots.map(node => ({ start: node.start, end: node.end, text: '' }))));
  }
  const rels = await zip.file('word/_rels/settings.xml.rels')?.async('string');
  if (rels) {
    const nodes = xmlSpans(rels).filter(node => node.local === 'Relationship' && /\/(mailMergeSource|recipientData)$/.test(node.attrs.Type));
    zip.file('word/_rels/settings.xml.rels', applyXmlEdits(rels, nodes.map(node => ({ start: node.start, end: node.end, text: '' }))));
  }
}
/** 功能：每份模板使用独立节末定义，边界改成下一页分节；保留原纸张/边距/页码和节内连续分节；时间：2026-11-11；作者：lq。 */
function recordBreak(section: string): string {
  if (/\/\s*>$/.test(section)) section = section.replace(/\/\s*>$/, '></w:sectPr>');
  const nodes = xmlSpans(section); const type = nodes.find(node => wordXmlNode(node, 'type'));
  if (type) section = applyXmlEdits(section, [{ start: type.start, end: type.end, text: '<w:type w:val="nextPage"/>' }]);
  else { const root = nodes[0]; section = section.slice(0, root.openEnd) + '<w:type w:val="nextPage"/>' + section.slice(root.openEnd); }
  return `<w:p><w:pPr>${section}</w:pPr></w:p>`;
}
/** 功能：附属部件移到 word 根目录后重写相对附件路径，不移动或覆盖原图片；时间：2026-11-11；作者：lq。 */
function relativePartTarget(from: string, to: string): string {
  const base = from.split('/').slice(0, -1), target = to.split('/');
  while (base.length && target.length && base[0] === target[0]) { base.shift(); target.shift(); }
  return [...base.map(() => '..'), ...target].join('/');
}
/** 功能：将段落继承的样式编号显式化，使每份记录独立续号/重启，同时保持样式其他属性不变；时间：2026-11-11；作者：lq。 */
function materializeStyleNumbering(xml: string, stylesXml?: string): string {
  if (!stylesXml) return xml;
  const styles = xmlSpans(stylesXml).filter(node => wordXmlNode(node, 'style'));
  const resolve = (id: string | undefined, seen = new Set<string>()): { id?: string; level?: string } => {
    if (!id || seen.has(id)) return {}; seen.add(id);
    const style = styles.find(node => wordXmlAttr(node, 'styleId') === id); if (!style) return {};
    const inherited = resolve(wordXmlAttr(style.children.find(node => wordXmlNode(node, 'basedOn')) ?? style, 'val'), seen);
    const pPr = style.children.find(node => wordXmlNode(node, 'pPr'));
    const numPr = pPr?.children.find(node => wordXmlNode(node, 'numPr'));
    const number = numPr?.children.find(node => wordXmlNode(node, 'numId')), level = numPr?.children.find(node => wordXmlNode(node, 'ilvl'));
    return { id: number ? wordXmlAttr(number, 'val') : inherited.id, level: level ? wordXmlAttr(level, 'val') : inherited.level };
  };
  const defaultStyle = styles.find(node => wordXmlAttr(node, 'type') === 'paragraph' && ['1', 'true', 'on'].includes(wordXmlAttr(node, 'default') ?? ''));
  const edits: XmlEdit[] = [];
  for (const paragraph of xmlSpans(xml).filter(node => wordXmlNode(node, 'p'))) {
    const properties = paragraph.children.find(node => wordXmlNode(node, 'pPr'));
    const style = properties?.children.find(node => wordXmlNode(node, 'pStyle'));
    const numPr = properties?.children.find(node => wordXmlNode(node, 'numPr'));
    if (numPr?.children.some(node => wordXmlNode(node, 'numId'))) continue;
    const inherited = resolve(style ? wordXmlAttr(style, 'val') : defaultStyle ? wordXmlAttr(defaultStyle, 'styleId') : undefined);
    if (!inherited.id || inherited.id === '0') continue;
    const number = `<w:numId w:val="${wordXmlEscape(inherited.id)}"/>`;
    if (numPr) { const fragment = xml.slice(numPr.start, numPr.end); edits.push({ start: numPr.start, end: numPr.end, text: appendRoot(fragment, number) }); }
    else {
      const numberProperties = `<w:numPr>${inherited.level ? `<w:ilvl w:val="${wordXmlEscape(inherited.level)}"/>` : ''}${number}</w:numPr>`;
      if (properties) {
        if (properties.end === properties.openEnd) edits.push({ start: properties.start, end: properties.end, text: appendRoot(xml.slice(properties.start, properties.end), numberProperties) });
        else {
          const preceding = new Set(['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl']);
          const at = properties.children.find(node => !preceding.has(node.local))?.start ?? properties.closeStart;
          edits.push({ start: at, end: at, text: numberProperties });
        }
      } else if (paragraph.end === paragraph.openEnd) edits.push({ start: paragraph.start, end: paragraph.end, text: appendRoot(xml.slice(paragraph.start, paragraph.end), `<w:pPr>${numberProperties}</w:pPr>`) });
      else edits.push({ start: paragraph.openEnd, end: paragraph.openEnd, text: `<w:pPr>${numberProperties}</w:pPr>` });
    }
  }
  return applyXmlEdits(xml, edits);
}
/** 功能：按真实记录和 NEXT 游标生成单份可编辑 DOCX，取消/失败只丢弃临时包，源快照永不修改；时间：2026-11-11；作者：lq。 */
export async function mergeWordPackage(bytes: Uint8Array, records: readonly MergeRecord[], options: NativeWordMergeOptions = {}): Promise<Uint8Array> {
  if (!records.length || records.length > 500) throw new Error('请选择 1 至 500 条合并记录。');
  const cancelled = () => { if (options.isCancelled?.()) throw new Error('已取消生成，本次结果未保存。'); }; cancelled();
  const zip = await loadDocxZip(bytes); cancelled();
  if (Object.keys(zip.files).some(path => /(?:vbaProject\.bin$|^_xmlsignatures\/)/i.test(path))) throw new Error('宏或数字签名文档不能生成合并编辑副本。');
  const document = await zip.file('word/document.xml')?.async('string'); if (!document) throw new Error('缺少 Word 主文档。');
  const info = await inspectWordMergePackage(bytes); if (!info.fields.length) throw new Error('模板尚未插入合并字段。');
  const rawRels = await zip.file('word/_rels/document.xml.rels')?.async('string') ?? `<Relationships xmlns="${PKG_REL}"></Relationships>`;
  const relations = xmlSpans(rawRels).filter(node => node.local === 'Relationship'); const newRel = allocator(relations.map(node => node.attrs.Id), 'rIdMoMerge');
  let rels = rawRels, types = await zip.file('[Content_Types].xml')!.async('string');
  const pathId = allocator(Object.keys(zip.files), 'word/mo-merge-');
  // 功能：避让原包全部十进制及十六进制标识，不依赖固定起点或假设原 ID 较小；时间：2026-11-11；作者：lq。
  const usedIds = new Set<number>(), usedNames = new Set<string>();
  for (const path of Object.keys(zip.files).filter(path => /^word\/.*\.xml$/.test(path))) for (const node of xmlSpans(await zip.file(path)!.async('string'))) {
    for (const [attribute, value] of Object.entries(node.attrs)) {
      const local = attribute.split(':').at(-1)!;
      if (['id', 'numId', 'abstractNumId'].includes(local) && /^\d+$/.test(value)) usedIds.add(Number(value));
      if (['paraId', 'paraIdParent', 'textId', 'anchorId', 'editId', 'durableId'].includes(local) && /^[0-9a-f]{8}$/i.test(value)) usedIds.add(parseInt(value, 16));
    }
    if (wordXmlNode(node, 'bookmarkStart') && wordXmlAttr(node, 'name')) usedNames.add(wordXmlAttr(node, 'name')!);
  }
  let serial = 0; const number = () => { do { serial++; } while (usedIds.has(serial)); if (serial > 2147483647) throw new Error('合并标识超出 Word 范围。'); usedIds.add(serial); return serial; };
  const newPart = (kind: string) => { let path: string; do { path = `${pathId()}-${kind}.xml`; } while (zip.file(path)); return path; };
  const preservedNotes = new Set<string>();
  const blankIds = new Map<string, string>();
  const originalRefs = xmlSpans(document).filter(node => wordXmlNode(node, 'headerReference') || wordXmlNode(node, 'footerReference'));
  const variants = [...new Set(originalRefs.map(node => `${wordXmlNode(node, 'headerReference') ? 'header' : 'footer'}:${wordXmlAttr(node, 'type') ?? 'default'}`))];
  const numbering = await zip.file('word/numbering.xml')?.async('string'); const newNums: string[] = [];
  const styledDocument = materializeStyleNumbering(document, await zip.file('word/styles.xml')?.async('string'));
  const extensions = new Map<string, { xml: string; items: XmlSpan[]; copies: string[] }>();
  for (const [path, itemName] of [['commentsExtended', 'commentEx'], ['commentsIds', 'commentId'], ['commentsExtensible', 'commentExtensible']]) {
    const xml = await zip.file(`word/${path}.xml`)?.async('string');
    if (xml) extensions.set(path, { xml, items: xmlSpans(xml).filter(node => node.local === itemName), copies: [] });
  }
  const notes = new Map<string, { xml: string; root: XmlSpan; items: XmlSpan[]; copies: string[] }>();
  for (const kind of ['footnotes', 'endnotes', 'comments']) {
    const xml = await zip.file(`word/${kind}.xml`)?.async('string'); if (!xml) continue;
    const nodes = xmlSpans(xml); const root = nodes.find(node => !node.parent)!;
    notes.set(kind, { xml, root, items: root.children.filter(node => wordXmlNode(node, kind === 'comments' ? 'comment' : kind === 'footnotes' ? 'footnote' : 'endnote') && !['separator', 'continuationSeparator', 'continuationNotice'].includes(wordXmlAttr(node, 'type') ?? '')), copies: [] });
  }
  const parts: Array<{ content: string; section: string }> = []; let record = 0, pass = 0;
  while (record < records.length) {
    cancelled(); const start = record; const cursor = { index: start, skipped: false }; const numberIds = new Map<string, string>();
    const identify = recordIdentity(++pass, number, numberIds, preservedNotes, usedNames);
    // 功能：每个 NEXT 记录独立生成注释/脚注缓存，同时保留相同记录内的引用和线程映射；时间：2026-11-11；作者：lq。
    const contexts = new Map<number, { ids: Map<string, string>; identify: ReturnType<typeof recordIdentity> }>();
    const context = (index: number) => { if (!contexts.has(index)) contexts.set(index, { ids: new Map(), identify: recordIdentity(pass, number, numberIds, preservedNotes, usedNames) }); return contexts.get(index)!; };
    const allocateNote = (kind: string, old: string, index: number) => {
      if (!notes.get(kind === 'comment' ? 'comments' : kind === 'footnote' ? 'footnotes' : 'endnotes')?.items.some(item => wordXmlAttr(item, 'id') === old)) throw new Error('合并模板引用了缺失的注释部件。');
      const ids = context(index).ids, key = `${kind}:${old}`;
      if (!ids.has(key)) { const id = String(number()); ids.set(key, id); preservedNotes.add(`${kind}:${id}`); }
      return ids.get(key)!;
    };
    const commentOrigins = new Map<string, number>();
    const reference = { id: (kind: string, old: string, index: number) => {
      if (kind === 'comment') { if (!commentOrigins.has(old)) commentOrigins.set(old, index); index = commentOrigins.get(old)!; }
      return allocateNote(kind, old, index);
    } };
    let merged = mergeWordXml(styledDocument, records, cursor, true, reference);
    record = cursor.index + 1;
    if (cursor.skipped) { options.onProgress?.(Math.min(record, records.length), records.length); continue; }
    // 功能：每份模板的首节补齐缺省空白变体，避免继承上一份末节的页眉页脚；时间：2026-11-11；作者：lq。
    let firstSection = xmlSpans(merged).find(node => wordXmlNode(node, 'sectPr'));
    if (firstSection && variants.length) {
      let section = merged.slice(firstSection.start, firstSection.end);
      if (/\/\s*>$/.test(section)) section = section.replace(/\/\s*>$/, '></w:sectPr>');
      for (const variant of variants) {
        const [kind, type] = variant.split(':'); const nodes = xmlSpans(section), root = nodes[0];
        if (root.children.some(node => wordXmlNode(node, `${kind}Reference`) && (wordXmlAttr(node, 'type') ?? 'default') === type)) continue;
        if (!blankIds.has(kind)) {
          const id = newRel(), path = newPart(kind); blankIds.set(kind, id);
          zip.file(path, `<w:${kind === 'header' ? 'hdr' : 'ftr'} xmlns:w="${WORD}"><w:p/></w:${kind === 'header' ? 'hdr' : 'ftr'}>`);
          rels = appendRoot(rels, `<Relationship Id="${id}" Type="${REL}/${kind}" Target="${path.slice(5)}"/>`);
          types = appendRoot(types, `<Override PartName="/${path}" ContentType="${CONTENT}${kind}+xml"/>`);
        }
        const at = kind === 'header' ? root.openEnd : root.children.find(node => !wordXmlNode(node, 'headerReference'))?.start ?? root.closeStart;
        section = section.slice(0, at) + `<w:${kind}Reference w:type="${type}" r:id="${blankIds.get(kind)}"/>` + section.slice(at);
      }
      merged = applyXmlEdits(merged, [{ start: firstSection.start, end: firstSection.end, text: section }]);
    }
    const refs = xmlSpans(merged).filter(node => wordXmlNode(node, 'headerReference') || wordXmlNode(node, 'footerReference'));
    const referenceIds = new Map<string, string>();
    for (const ref of refs) {
      const originalId = ref.attrs['r:id']; if (referenceIds.has(originalId)) continue; if ([...blankIds.values()].includes(originalId)) { referenceIds.set(originalId, originalId); continue; }
      const relation = relations.find(node => node.attrs.Id === originalId); if (!relation || relation.attrs.TargetMode === 'External') throw new Error('页眉页脚关系不完整，已取消合并。');
      const path = resolveRelationshipTargetPath('word/document.xml', relation.attrs.Target); if (!path) throw new Error('页眉页脚路径无效。');
      const source = await zip.file(path)?.async('string'); if (!source) throw new Error('页眉页脚部件缺失。');
      const kind = wordXmlNode(ref, 'headerReference') ? 'header' : 'footer';
      const newPath = newPart(kind), id = newRel(); referenceIds.set(originalId, id);
      zip.file(newPath, identify(mergeWordXml(source, records, { index: start, skipped: false }, false, reference)));
      const sourceRels = path.replace(/([^/]+)$/, '_rels/$1.rels'); const sourceRelations = zip.file(sourceRels);
      if (sourceRelations) {
        const xml = await sourceRelations.async('string');
        const nodes = xmlSpans(xml).filter(node => node.local === 'Relationship' && node.attrs.TargetMode !== 'External');
        zip.file(newPath.replace(/([^/]+)$/, '_rels/$1.rels'), applyXmlEdits(xml, nodes.map(node => {
          const resolved = resolveRelationshipTargetPath(path, node.attrs.Target); if (!resolved) throw new Error('页眉页脚附件路径无效。');
          return { start: node.start, end: node.openEnd, text: replaceAttribute(xml.slice(node.start, node.openEnd), 'Target', relativePartTarget(newPath, resolved)) };
        })));
      }
      rels = appendRoot(rels, `<Relationship Id="${id}" Type="${REL}/${kind}" Target="${newPath.slice(5)}"/>`);
      types = appendRoot(types, `<Override PartName="/${newPath}" ContentType="${CONTENT}${kind}+xml"/>`);
    }
    merged = applyXmlEdits(merged, refs.map(ref => ({ start: ref.start, end: ref.openEnd, text: replaceAttribute(merged.slice(ref.start, ref.openEnd), 'r:id', referenceIds.get(ref.attrs['r:id'])!) })));
    merged = identify(merged);
    for (const [index, current] of contexts) for (const [kind, note] of notes) {
      // 功能：批注回复可能只有线程关系而无正文引用，整组按相同数据记录复制；时间：2026-11-11；作者：lq。
      const singular = kind === 'footnotes' ? 'footnote' : kind === 'endnotes' ? 'endnote' : 'comment';
      for (const item of note.items) {
        const old = wordXmlAttr(item, 'id')!, id = allocateNote(singular, old, index);
        const fragment = replaceAttribute(note.xml.slice(item.start, item.end), `${item.prefix}:id`, id);
        const wrapper = note.xml.slice(0, note.root.openEnd) + fragment + note.xml.slice(note.root.closeStart);
        const identified = current.identify(mergeWordXml(wrapper, records, { index, skipped: false }, false));
        const root = xmlSpans(identified).find(node => !node.parent)!;
        note.copies.push(identified.slice(root.openEnd, root.closeStart));
      }
    }
    // 功能：现代线程批注的段落/父批注/持久标识沿用同一记录映射，避免回复指向模板旧 ID；时间：2026-11-11；作者：lq。
    for (const current of contexts.values()) for (const extension of extensions.values()) {
      const root = xmlSpans(extension.xml).find(node => !node.parent)!;
      for (const item of extension.items) {
        const identified = current.identify(extension.xml.slice(0, root.openEnd) + extension.xml.slice(item.start, item.end) + extension.xml.slice(root.closeStart));
        const node = xmlSpans(identified).find(node => !node.parent)!;
        extension.copies.push(identified.slice(node.openEnd, node.closeStart));
      }
    }
    if (numbering && numberIds.size) {
      const nums = xmlSpans(numbering).filter(node => wordXmlNode(node, 'num'));
      for (const [old, id] of numberIds) {
        const source = nums.find(node => wordXmlAttr(node, 'numId') === old); if (!source) throw new Error('合并模板引用了缺失的编号定义。');
        newNums.push(replaceAttribute(numbering.slice(source.start, source.end), `${source.prefix}:numId`, id));
      }
    }
    const nodes = xmlSpans(merged); const body = nodes.find(node => wordXmlNode(node, 'body'))!;
    const section = body.children.find(node => wordXmlNode(node, 'sectPr'));
    parts.push({ content: merged.slice(body.openEnd, section?.start ?? body.closeStart), section: section ? merged.slice(section.start, section.end) : '<w:sectPr></w:sectPr>' });
    options.onProgress?.(Math.min(record, records.length), records.length);
    await new Promise<void>(resolve => setTimeout(resolve, 0));
  }
  cancelled(); if (!parts.length) throw new Error('所有记录均被 SKIPIF 排除，没有可输出的报告。');
  const nodes = xmlSpans(document); const body = nodes.find(node => wordXmlNode(node, 'body'))!;
  const content = parts.map((part, index) => part.content + (index === parts.length - 1 ? part.section : recordBreak(part.section))).join('');
  zip.file('word/document.xml', document.slice(0, body.openEnd) + content + document.slice(body.closeStart));
  zip.file('word/_rels/document.xml.rels', rels); zip.file('[Content_Types].xml', types);
  if (numbering && newNums.length) zip.file('word/numbering.xml', appendRoot(numbering, newNums.join('')));
  for (const [kind, note] of notes) {
    const edits = note.items.map(item => ({ start: item.start, end: item.end, text: '' }));
    zip.file(`word/${kind}.xml`, appendRoot(applyXmlEdits(note.xml, edits), note.copies.join('')));
  }
  for (const [path, extension] of extensions) zip.file(`word/${path}.xml`, appendRoot(applyXmlEdits(extension.xml, extension.items.map(item => ({ start: item.start, end: item.end, text: '' }))), extension.copies.join('')));
  await clearMergeSettings(zip); cancelled();
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }, () => cancelled());
}
