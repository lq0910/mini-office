/** 功能：格式刷只搬运已支持的字符效果 XML，保留目标语言、链接及其余未知属性；时间：2026-11-11；作者：lq。 */
import { splitXmlChildren, RPR_CHILD_ORDER } from '../upstream/packages/docx-engine/src/generate';
import { wordXmlEscape } from './xml-spans';
const EFFECTS = new Set(['w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:bdr', 'w:em', 'w14:textOutline', 'w14:glow']);
/** 功能：复制来源原始效果以保留线宽/主题/透明度，继承的边框和阴影可由摘要补出；时间：2026-11-11；作者：lq。 */
export function wordRunEffectsXml(attrs: Record<string, unknown>): string {
  const raw = String(attrs.rawRPr ?? '').replace(/^<w:rPr\b[^>]*>/, '').replace(/<\/w:rPr>$/, '');
  const effects = splitXmlChildren(raw).filter(child => EFFECTS.has(child.name));
  if (attrs.textEffect && !effects.some(child => child.name === `w:${attrs.textEffect}`) && ['outline', 'shadow', 'emboss', 'imprint'].includes(String(attrs.textEffect))) effects.push({ name: `w:${attrs.textEffect}`, xml: `<w:${attrs.textEffect}/>` });
  if (attrs.em && !effects.some(child => child.name === 'w:em')) effects.push({ name: 'w:em', xml: `<w:em w:val="${wordXmlEscape(String(attrs.em))}"/>` });
  if (attrs.bdr && !effects.some(child => child.name === 'w:bdr')) {
    const bdr = typeof attrs.bdr === 'string' ? JSON.parse(attrs.bdr) : attrs.bdr;
    effects.push({ name: 'w:bdr', xml: `<w:bdr${Object.entries(bdr as Record<string, unknown>).filter(([, value]) => value != null).map(([key, value]) => ` w:${key}="${wordXmlEscape(String(value))}"`).join('')}/>` });
  }
  return effects.map(child => child.xml.replace(/<(w14:[^\s/>]+)/, '<$1 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')).join('');
}
/** 功能：显式效果摘要覆盖目标对应节点，非目标 rPr 子节点保持原字节；时间：2026-11-11；作者：lq。 */
export function patchWordRunEffects(rawRPr: string | undefined, effectXml: string): string {
  const raw = (rawRPr ?? '').replace(/^<w:rPr\b[^>]*\/>$/, '').replace(/^<w:rPr\b[^>]*>/, '').replace(/<\/w:rPr>$/, '');
  const effects = splitXmlChildren(effectXml).filter(child => EFFECTS.has(child.name));
  const children = [...splitXmlChildren(raw).filter(child => !EFFECTS.has(child.name)), ...effects];
  const rank = (name: string) => { const at = RPR_CHILD_ORDER.indexOf(name); return at < 0 ? RPR_CHILD_ORDER.length : at; };
  children.sort((a, b) => rank(a.name) - rank(b.name));
  return children.length ? `<w:rPr>${children.map(child => child.xml).join('')}</w:rPr>` : '';
}
