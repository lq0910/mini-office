/** 功能：复用同源标题规则和 TOC 域生成，目录操作共享正文历史与保存；时间：2026-11-11；作者：lq。
 * 参考固定 GenOffice e1f50e90 的 ribbon-references-tab；来源及许可见 packages/word-core/UPSTREAM.md。
 */
import type { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { TextSelection } from '@tiptap/pm/state';
import type { PmNode } from '../upstream/apps/docs/src/renderer/editor/convert';
import { collectHeadings } from '../upstream/apps/docs/src/renderer/editor/headings';
import { generateTocFieldXml, type ParsedDocFull, type FieldDisplay } from '../upstream/packages/docx-engine/src/index';
export { collectHeadings as wordHeadings } from '../upstream/apps/docs/src/renderer/editor/headings';
export type WordHeadingPages = ReadonlyMap<number, string>;

/** 功能：定位可完整替换的独立目录行；拒绝夹带正文、分节或复杂容器的导入目录；时间：2026-11-11；作者：lq。 */
function tocRegion(editor: Editor, parsed: ParsedDocFull) {
  const lines: Array<{ from: number; to: number; xml: string }> = [];
  let unsafe = false;
  editor.state.doc.forEach((node, from) => {
    const original = parsed.blocks.find(block => block.docxIndex === node.attrs.docxIndex)?.originalXml ?? '';
    const xml = String(node.attrs.genXml ?? original);
    if (node.attrs.fieldDisplay?.kind !== 'tocLine') {
      if (/\bTOC[\s\\]/.test(xml) && /<w:instrText\b/.test(xml)) unsafe = true;
      return;
    }
    if (node.type.name !== 'docProtected' || /<w:(?:sectPr|sdt|tbl|drawing|br)\b/.test(xml)) unsafe = true;
    lines.push({ from, to: from + node.nodeSize, xml });
  });
  if (!lines.length) { if (unsafe) throw new Error('此目录含复杂结构，暂保留原内容；请在 Word 中更新。'); return undefined; }
  if (unsafe || lines.some((line, i) => i > 0 && lines[i-1].to !== line.from)) throw new Error('目录包含多个区域或复杂结构，暂不能安全更新。');
  const xml = lines.map(line => line.xml).join('');
  if ((xml.match(/w:fldCharType=["']begin["']/g) ?? []).length !== 1 || (xml.match(/w:fldCharType=["']end["']/g) ?? []).length !== 1) throw new Error('目录域边界不完整，暂不能安全更新。');
  // 功能：域边界之外的正文不能随目录更新被删除；仅接受独立段落里的普通缓存；时间：2026-11-11；作者：lq。
  const tree = new DOMParser().parseFromString(`<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${xml}</root>`, 'application/xml');
  if (tree.querySelector('parsererror') || Array.from(tree.documentElement.children).some(node => node.localName !== 'p')) throw new Error('目录结构无法完整识别，原内容保持不变。');
  let depth = 0;
  for (const element of Array.from(tree.documentElement.getElementsByTagName('*'))) {
    if (['bookmarkStart','bookmarkEnd','commentRangeStart','commentRangeEnd'].includes(element.localName)) throw new Error('目录包含跨区域引用，暂不能安全更新。');
    if (element.localName === 'fldChar') {
      const kind = element.getAttribute('w:fldCharType');
      if (kind === 'begin') depth++;
      if (kind === 'end') depth--;
      if (depth < 0) throw new Error('目录域边界不完整，原内容保持不变。');
    }
    if (element.localName === 't' && depth === 0 && element.textContent?.trim()) throw new Error('目录段落夹带正文，暂不能安全更新。');
  }
  return { from: lines[0].from, to: lines[lines.length-1].to };
}
/** 功能：在正文段落前插入目录或原位更新完整目录，整批仅一步撤销；时间：2026-11-11；作者：lq。 */
export function insertWordToc(editor: Editor, parsed: ParsedDocFull, update = false): boolean {
  if (!editor.isEditable || editor.view.composing) return false;
  if (editor.state.selection.$from.depth > 1) throw new Error('请先将光标放到正文段落，再插入目录。');
  const entries = collectHeadings(editor.state.doc, parsed.styles).filter(entry => entry.level <= 3);
  if (!entries.length) throw new Error('请先应用标题 1–3 样式，再插入目录。');
  const region = tocRegion(editor, parsed);
  if (region && !update) throw new Error('文档已有目录，请使用“更新目录”。');
  if (!region && update) throw new Error('文档尚无目录，请先插入目录。');
  const nodes = generateTocFieldXml(entries).map((xml, i) => editor.schema.nodes.docProtected.create({
    docxIndex: null, blockType: 'passthrough', label: '目录域', genXml: xml.replace('<w:tab/></w:r>', '<w:tab/></w:r><w:r><w:t/></w:r>'),
    fieldDisplay: { kind: 'tocLine', left: entries[i].text, right: '', level: entries[i].level },
  }));
  const from = region?.from ?? (editor.state.selection.$from.depth ? editor.state.selection.$from.before(1) : editor.state.selection.from);
  const tr = closeHistory(editor.state.tr).replaceWith(from, region?.to ?? from, nodes);
  tr.setSelection(TextSelection.near(tr.doc.resolve(from + nodes.reduce((sum, node) => sum + node.nodeSize, 0))));
  editor.view.dispatch(tr); return true;
}
/** 功能：真实分页结果仅补入保存快照，排版本身不产生 dirty、事务或历史；时间：2026-11-11；作者：lq。 */
export function wordTocSnapshot(document: PmNode, headings: ReturnType<typeof collectHeadings>, pages: WordHeadingPages): PmNode {
  const byTitle = new Map<string, Array<string | undefined>>();
  for (const heading of headings) { const key = `${heading.level}:${heading.text.trim()}`; byTitle.set(key, [...byTitle.get(key) ?? [], pages.get(heading.pos)]); }
  const seen = new Map<string, number>();
  return { ...document, content: document.content?.map(node => {
    const field = node.attrs?.fieldDisplay as FieldDisplay | undefined;
    if (node.type !== 'docProtected' || field?.kind !== 'tocLine') return node;
    const key = `${field.level ?? 1}:${String(field.left ?? '').trim()}`; const nth = seen.get(key) ?? 0; seen.set(key, nth + 1);
    const right = byTitle.get(key)?.[nth];
    return right === undefined ? node : { ...node, attrs: { ...node.attrs, fieldDisplay: { ...field, right } } };
  }) };
}

/** 功能：导入目录优先跳到原书签，新目录按级别及同名标题顺序定位；导航只改变选区和视图；时间：2026-11-11；作者：lq。 */
export function followWordToc(editor: Editor, parsed: ParsedDocFull, lineIndex: number): boolean {
  if (editor.view.composing) return false;
  let index = 0; let target: { text: string; level: number; nth: number; anchor?: string } | undefined;
  const seen = new Map<string, number>();
  editor.state.doc.forEach(node => {
    const field = node.attrs.fieldDisplay as FieldDisplay | undefined;
    if (field?.kind !== 'tocLine') return;
    const text = String(field.left ?? '').trim(); const level = Number(field.level ?? 1);
    const key = `${level}:${text}`; const nth = seen.get(key) ?? 0; seen.set(key, nth+1);
    if (index++ === lineIndex) target = { text, level, nth, anchor: field.anchor };
  });
  if (!target) return false;
  // 功能：Word 的 TOC 链接常指向普通段落的隐藏 _Toc 书签，不能要求所有目标都有标题样式；时间：2026-11-11；作者：lq。
  let position: number | undefined;
  if (target.anchor) {
    // 功能：同时检查 bookmarks 和 hiddenBookmarks，以及直接比较锚点名；时间：2026-11-13；作者：lq。
    // 在闭包外固定 anchor 值，避免 TypeScript 无法在回调内收窄 target.anchor 的 undefined 类型；
    const anchor = target.anchor;
    const anchorNorm = anchor.replace(/^_/, '').toLowerCase();
    editor.state.doc.descendants((node, pos) => {
      if (position !== undefined || node.attrs.fieldDisplay?.kind === 'tocLine') return false;
      const bookmarks: string[] = [...(node.attrs.bookmarks ?? []), ...(node.attrs.hiddenBookmarks ?? [])];
      // 精确匹配优先
      if (bookmarks.includes(anchor)) { position = pos; return false; }
      // 容错：忽略前缀下划线和大小写差异（某些 doc2docx 转换可能改变格式）
      if (bookmarks.some(b => b.replace(/^_/, '').toLowerCase() === anchorNorm)) { position = pos; return false; }
    });
  }
  if (position === undefined) {
    const normalize = (text: string) => text.replace(/\s+/gu, '').normalize('NFKC');
    const heading = collectHeadings(editor.state.doc, parsed.styles).filter(item => normalize(item.text) === normalize(target!.text) && item.level === target!.level)[target.nth];
    position = heading?.pos;
    if (position === undefined) {
      // 功能：旧 DOC 转换可能缺失书签/标题样式，先匹配独立正文标题；不拿来源缓存页码猜测位置；时间：2026-11-11；作者：lq。
      const paragraphs: Array<{ text: string; raw: string; pos: number }> = [];
      editor.state.doc.forEach((node, pos) => { if (node.isTextblock && !node.attrs.fieldDisplay) paragraphs.push({ text: normalize(node.textContent), raw: node.textContent, pos }); });
      const title = normalize(target.text);
      const exact = paragraphs.filter(item => item.text === title);
      position = exact[target.nth]?.pos;
      if (position === undefined && /^第[一二三四五六七八九十百零〇两\d]+[章节篇部]/u.test(title)) {
        // 功能：旧 DOC 将章标题接在上一句末尾时，仅接受完整章名、句末边界和唯一目标，并定位标题字符；时间：2026-11-11；作者：lq。
        const attached = paragraphs.filter(item => item.text.endsWith(title) && /[。！？；]$/u.test(item.text.slice(0, -title.length)));
        if (attached.length === 1) {
          const item = attached[0], prefixLength = item.text.length - title.length;
          let offset = 0;
          while (offset < item.raw.length && normalize(item.raw.slice(0, offset)).length < prefixLength) offset++;
          while (offset < item.raw.length && /\s/u.test(item.raw[offset])) offset++;
          position = item.pos + offset;
        }
      }
      if (position === undefined) {
        // 功能：列表生成的章节号可能不在文字中；仅接受唯一的短标题后缀，拒绝模糊匹配正文句子；时间：2026-11-11；作者：lq。
        const plain = title.replace(/^第[一二三四五六七八九十百零〇两\d]+[章节篇部][、.．:：\-—]*/u, '');
        const candidates = plain.length >= 4 ? paragraphs.filter(item => item.text.length <= Math.max(80, plain.length * 3) && !/[。！？；]/u.test(item.text) && item.text.endsWith(plain)) : [];
        if (candidates.length === 1) position = candidates[0].pos;
      }
    }
  }
  if (position === undefined) return false;
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(position+1))).scrollIntoView());
  editor.view.focus(); return true;
}
