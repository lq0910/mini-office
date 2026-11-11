/** 功能：为局部 OOXML 替换保留原始字符区间、属性和命名空间，不重新序列化未修改节点；时间：2026-11-11；作者：lq。 */
import { decodeEntities } from '../upstream/packages/docx-engine/src/parse-xml-text';

export interface XmlSpan { name: string; local: string; prefix: string; attrs: Record<string, string>; start: number; openEnd: number; closeStart: number; end: number; parent?: XmlSpan; children: XmlSpan[]; }
export interface XmlEdit { start: number; end: number; text: string; }
/** 功能：扫描带引号属性、注释及 CDATA 的 XML，并拒绝不平衡结构与 DTD；时间：2026-11-11；作者：lq。 */
export function xmlSpans(xml: string): XmlSpan[] {
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) throw new Error('邮件合并不接受包含 DTD 的 XML。');
  const all: XmlSpan[] = [], stack: XmlSpan[] = [];
  const tags = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(?:[^>"']|"[^"]*"|'[^']*')*>/g;
  for (const match of xml.matchAll(tags)) {
    const tag = match[0], at = match.index!;
    if (/^<[/!?]/.test(tag)) {
      if (!tag.startsWith('</')) continue;
      const node = stack.pop(); if (!node || tag.slice(2, -1).trim() !== node.name) throw new Error('OOXML 结构不完整，已取消合并。');
      node.closeStart = at; node.end = at + tag.length; continue;
    }
    const name = /^<([^\s/>]+)/.exec(tag)?.[1]; if (!name) continue;
    const colon = name.indexOf(':'); const attrs: Record<string, string> = {};
    for (const attr of tag.matchAll(/([^\s=<>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) attrs[attr[1]] = decodeEntities(attr[2] ?? attr[3]);
    const node: XmlSpan = { name, local: colon < 0 ? name : name.slice(colon + 1), prefix: colon < 0 ? '' : name.slice(0, colon), attrs, start: at, openEnd: at + tag.length, closeStart: at + tag.length, end: at + tag.length, parent: stack.at(-1), children: [] };
    node.parent?.children.push(node); all.push(node); if (!/\/\s*>$/.test(tag)) stack.push(node);
  }
  if (stack.length) throw new Error('OOXML 结构不完整，已取消合并。');
  return all;
}
/** 功能：命名空间判定沿祖先解析，支持 Word 的其他合法前缀；时间：2026-11-11；作者：lq。 */
export function wordXmlNode(node: XmlSpan, local: string): boolean {
  if (node.local !== local) return false;
  const key = node.prefix ? `xmlns:${node.prefix}` : 'xmlns';
  for (let p: XmlSpan | undefined = node; p; p = p.parent) if (p.attrs[key]) return /\/wordprocessingml\/(?:2006\/main|main)$/.test(p.attrs[key]) || p.attrs[key] === 'http://purl.oclc.org/ooxml/wordprocessingml/main';
  return node.prefix === 'w';
}
/** 功能：使用元素自己的前缀读取 Word 属性，避免把未知命名空间同名属性当作正文；时间：2026-11-11；作者：lq。 */
export const wordXmlAttr = (node: XmlSpan, name: string) => node.attrs[`${node.prefix || 'w'}:${name}`] ?? node.attrs[name];
/** 功能：倒序提交互不重叠的局部替换；重叠说明计划有误，不能静默输出损坏文档；时间：2026-11-11；作者：lq。 */
export function applyXmlEdits(xml: string, edits: XmlEdit[]): string {
  const sorted = [...edits].sort((a, b) => b.start - a.start || b.end - a.end); let boundary = xml.length;
  for (const edit of sorted) {
    if (edit.start < 0 || edit.end < edit.start || edit.end > boundary) throw new Error('邮件合并 XML 修改区间重叠，已取消输出。');
    xml = xml.slice(0, edit.start) + edit.text + xml.slice(edit.end); boundary = edit.start;
  }
  return xml;
}
/** 功能：数据作为文本写入 XML，不允许字段值引入标签或控制字符；时间：2026-11-11；作者：lq。 */
export function wordXmlEscape(value: string): string {
  return value.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}
