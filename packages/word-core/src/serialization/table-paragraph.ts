/** 功能：表格重建时局部改写文字和格式，保留原段落外壳、书签、链接及未知 XML；时间：2026-11-11；作者：lq。 */
import type { ParaFormat, Run } from '../upstream/packages/docx-engine/src/types';
import { inlineRunsXml, mergePPrFormat, PPR_CHILD_ORDER, splitXmlChildren } from '../upstream/packages/docx-engine/src/generate';
import { decodeEntities } from '../upstream/packages/docx-engine/src/parse-xml-text';
import { escapeXmlAttr } from '../upstream/packages/docx-engine/src/xml-utils';

/** 功能：来源只存不可变段落和格式上下文，不产生另一份可编辑正文；时间：2026-11-11；作者：lq。 */
export interface WordCellParagraphSource { xml: string; format?: ParaFormat; styleId?: string | null; }
export function patchWordCellParagraph(source: WordCellParagraphSource, format: ParaFormat | undefined, runs: Run[], styleId: string | null, list?: { numId: string; ilvl: number }): string | undefined {
  const xml = source.xml;
  // 功能：域、绘图、修订与数学原子沿用各自的同源专用序列化，不能把缓存文字当普通 run 改写；时间：2026-11-11；作者：lq。
  if (/<(?:w:(?:fldChar|instrText|drawing|pict|object|ruby|sym|delText|ins|del)|m:)/.test(xml)) return undefined;
  const open = /^<w:p\b[^>]*>/.exec(xml)?.[0];
  if (!open || /\/>$/.test(open)) return undefined;
  const originalText = [...xml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)].map(match => decodeEntities(match[1])).join('');
  const text = runs.map(run => run.text).join('');
  if (/[\t\n\r\f\v]/.test(text) || /<w:(?:tab|br)\b/.test(xml)) return undefined;
  let prefix = 0; let suffix = 0;
  while (prefix < Math.min(originalText.length, text.length) && originalText[prefix] === text[prefix]) prefix++;
  while (suffix < Math.min(originalText.length, text.length) - prefix && originalText[originalText.length - suffix - 1] === text[text.length - suffix - 1]) suffix++;
  /** 功能：书签等注释边界跟随相同前后缀；替换内部的边界保持在替换区内；时间：2026-11-11；作者：lq。 */
  const mapOffset = (offset: number) => offset === 0 ? 0 : offset === originalText.length ? text.length : offset <= prefix ? offset : offset >= originalText.length - suffix ? text.length - (originalText.length - offset) : Math.min(text.length - suffix, offset);
  let cursor = 0;
  const render = (from: number, to: number, insideLink: boolean, runOpen: string) => {
    let offset = 0; let result = '';
    for (const run of runs) {
      const end = offset + run.text.length;
      const a = Math.max(offset, from); const b = Math.min(end, to);
      if (a < b) result += inlineRunsXml([{ ...run, text: run.text.slice(a - offset, b - offset), commentIds: undefined, ...(insideLink ? { link: undefined } : {}) }]).replace(/<w:r>/g, runOpen);
      offset = end;
    }
    return result;
  };
  /** 功能：只替换原文字 run，其他子节点和任意容器属性沿用原 XML；时间：2026-11-11；作者：lq。 */
  const walk = (inner: string, insideLink = false): string => splitXmlChildren(inner).map(child => {
    if (child.name === 'w:r') {
      const old = [...child.xml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)].map(match => decodeEntities(match[1])).join('');
      if (!old) return child.xml;
      const from = cursor; cursor += old.length;
      const start = /^<w:r\b[^>]*>/.exec(child.xml)?.[0] ?? '<w:r>';
      const unknown = splitXmlChildren(child.xml.slice(start.length, -'</w:r>'.length)).filter(item => !['w:rPr', 'w:t'].includes(item.name)).map(item => item.xml).join('');
      return render(mapOffset(from), mapOffset(cursor), insideLink, start) + (unknown ? `${start}${unknown}</w:r>` : '');
    }
    if (child.name === 'w:pPr') return '';
    if (/<w:t\b/.test(child.xml)) {
      const tag = /^<[^>]*>/.exec(child.xml)?.[0]; const end = `</${child.name}>`;
      return tag ? tag + walk(child.xml.slice(tag.length, -end.length), insideLink || child.name === 'w:hyperlink') + end : child.xml;
    }
    return child.xml;
  }).join('');
  const rawPPr = splitXmlChildren(xml.slice(open.length, -'</w:p>'.length)).find(child => child.name === 'w:pPr')?.xml ?? '';
  let pPr = mergePPrFormat(rawPPr, format, source.format);
  /** 功能：样式和编号只改目标组，未目标段落属性保留；时间：2026-11-11；作者：lq。 */
  const replaceProperty = (name: string, value: string) => {
    pPr = pPr.replace(new RegExp(`<w:${name}\\b[^>]*\\/>|<w:${name}\\b[^>]*>[\\s\\S]*?<\\/w:${name}>`, 'g'), '');
    if (!value) return;
    const start = /^<w:pPr\b[^>]*>/.exec(pPr)?.[0] ?? '<w:pPr>';
    const children = splitXmlChildren(pPr ? pPr.slice(start.length, -'</w:pPr>'.length) : '');
    // 功能：按 CT_PPr 序列插入目标属性，未知扩展仍保持原相对顺序；时间：2026-11-11；作者：lq。
    const rank = PPR_CHILD_ORDER.indexOf(`w:${name}`);
    const at = children.findIndex(child => PPR_CHILD_ORDER.indexOf(child.name) > rank);
    children.splice(at < 0 ? children.length : at, 0, { name: `w:${name}`, xml: value });
    pPr = start + children.map(child => child.xml).join('') + '</w:pPr>';
  };
  if (styleId !== (source.styleId ?? null)) replaceProperty('pStyle', styleId ? `<w:pStyle w:val="${escapeXmlAttr(styleId)}"/>` : '');
  replaceProperty('numPr', list ? `<w:numPr><w:ilvl w:val="${list.ilvl}"/><w:numId w:val="${escapeXmlAttr(list.numId)}"/></w:numPr>` : '');
  let content = walk(xml.slice(open.length, -'</w:p>'.length));
  // 功能：空格插入文字时在原注释结束边界前生成 run，不丢失首次输入；时间：2026-11-11；作者：lq。
  if (!originalText && text) {
    const boundary = content.search(/<w:(?:bookmarkEnd|commentRangeEnd)\b/);
    const at = boundary < 0 ? content.length : boundary;
    content = content.slice(0, at) + render(0, text.length, false, '<w:r>') + content.slice(at);
  }
  if (cursor !== originalText.length) return undefined;
  return open + pPr + content + '</w:p>';
}
