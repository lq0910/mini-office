/** 功能：段落套用样式/转换列表时只替换语义属性，分节、书签相关属性及未知 XML 保留；时间：2026-11-11；作者：lq。 */
import type { GeneratedBlock } from '../upstream/packages/docx-engine/src/types';
import { splitXmlChildren, PPR_CHILD_ORDER } from '../upstream/packages/docx-engine/src/generate';
import { wordXmlEscape } from './xml-spans';

/** 功能：身份变更不重建整个 pPr，避免工具栏样式操作丢失节边界和其他来源属性；时间：2026-11-11；作者：lq。 */
export function patchWordParagraphStructure(raw: string, generated: GeneratedBlock): string {
  const open = /^<w:pPr\b[^>]*>/.exec(raw)?.[0] ?? '<w:pPr>';
  const inner = raw.endsWith('</w:pPr>') ? raw.slice(open.length, -'</w:pPr>'.length) : '';
  const children = splitXmlChildren(inner).filter(child => !['w:pStyle', 'w:numPr', 'w:outlineLvl'].includes(child.name));
  if (generated.styleId) children.push({ name: 'w:pStyle', xml: `<w:pStyle w:val="${wordXmlEscape(generated.styleId)}"/>` });
  if (generated.type === 'listItem' && generated.list) children.push({ name: 'w:numPr', xml: `<w:numPr><w:ilvl w:val="${Math.max(0, Math.min(8, generated.list.ilvl))}"/><w:numId w:val="${wordXmlEscape(generated.list.numId)}"/></w:numPr>` });
  if (generated.type === 'heading') children.push({ name: 'w:outlineLvl', xml: `<w:outlineLvl w:val="${Math.max(0, Math.min(8, (generated.level ?? 1) - 1))}"/>` });
  const rank = (name: string) => PPR_CHILD_ORDER.indexOf(name) < 0 ? PPR_CHILD_ORDER.length : PPR_CHILD_ORDER.indexOf(name);
  return `${open}${children.sort((a, b) => rank(a.name) - rank(b.name)).map(child => child.xml).join('')}</w:pPr>`;
}
