/** 功能：把剪贴板 HTML 转为编辑器段落，只读取内容和格式，不插入外部脚本或事件；时间：2026-11-07；作者：lq。 */
import type { DocParagraph, ParagraphStyle, RunStyle } from '@mini-office/document-model';
import { makeParagraph } from './doc-utils';
import { serializeDomRuns } from './editing';

export type PasteMode = 'source' | 'destination' | 'text';

/** 功能：源格式保留字体与强调，匹配目标保留强调，纯文本只采用目标样式；时间：2026-11-07；作者：lq。 */
export function clipboardParagraphs(html: string, text: string, mode: PasteMode, destination: RunStyle, paragraphStyle: ParagraphStyle): DocParagraph[] {
  if (!html || mode === 'text') return text.replace(/\r\n?/g, '\n').split('\n').map((line) => makeParagraph([{ text: line, style: { ...destination } }], { style: { ...paragraphStyle } }));
  const root = document.createElement('div');
  root.innerHTML = html;
  root.querySelectorAll('script,style,link,meta,iframe,object,embed,img,svg').forEach((node) => node.remove());
  // 功能：保留块边界并避免嵌套段落重复采集，表格单元格按段落读取；时间：2026-11-07；作者：lq。
  const groups: HTMLElement[] = [];
  let inline = document.createElement('div');
  const flush = () => { if (inline.childNodes.length) { groups.push(inline); inline = document.createElement('div'); } };
  const visit = (parent: HTMLElement) => {
    for (const node of Array.from(parent.childNodes)) {
      if (node instanceof HTMLElement && /^(P|DIV|LI|H[1-6]|BLOCKQUOTE|TABLE|TBODY|THEAD|TR|TD|TH|UL|OL|SECTION|ARTICLE)$/.test(node.tagName)) {
        flush();
        // 功能：继承外层剪贴板容器的文字格式，保留网页整块复制时的字体；时间：2026-11-07；作者：lq。
        node.style.cssText = [parent.style.cssText, node.style.cssText].filter(Boolean).join(' ');
        if (node.querySelector('p,div,li,h1,h2,h3,h4,h5,h6,table,tr,td,th')) visit(node);
        else groups.push(node);
      } else if (node.nodeType !== Node.TEXT_NODE || (node.textContent?.trim() || !/[\r\n]/.test(node.textContent ?? ''))) inline.appendChild(node.cloneNode(true));
    }
  };
  visit(root); flush();
  if (!groups.length) return clipboardParagraphs('', text, 'text', destination, paragraphStyle);
  return groups.map((element) => {
    // 功能：将段落本身的行内样式纳入 run 解析，避免丢失整段字体和颜色；时间：2026-11-07；作者：lq。
    const wrapper = document.createElement('div');
    const span = document.createElement('span');
    span.style.cssText = element.style.cssText;
    span.innerHTML = element.innerHTML;
    wrapper.appendChild(span);
    const runs = serializeDomRuns(wrapper, destination, new Map()).map((run) => ({
      ...run,
      style: mode === 'source' ? run.style : { ...destination, bold: run.style.bold, italic: run.style.italic, underline: run.style.underline, strike: run.style.strike },
    }));
    const alignment = element.style.textAlign;
    return makeParagraph(runs, { style: { ...paragraphStyle, ...(mode === 'source' && /^(left|center|right|justify)$/.test(alignment) ? { alignment: alignment as ParagraphStyle['alignment'] } : {}) } });
  });
}
