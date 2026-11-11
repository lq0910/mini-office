/** 功能：标尺读取原生段落与继承制表位，拖动结果一次提交到正文历史；时间：2026-11-11；作者：lq。 */
import type { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import type { ParsedDocFull, TabStop } from '../upstream/packages/docx-engine/src/index';
import { selectedParagraphs } from './formatting';

/** 功能：合并样式和直接制表位，保留 clear/bar 等不在旧标尺显示的原生语义；时间：2026-11-11；作者：lq。 */
export function wordRulerTabs(attrs: Record<string, unknown>, parsed: ParsedDocFull): TabStop[] {
  const inherited = parsed.styles.get(String(attrs.styleId ?? 'Normal'))?.display?.tabStops ?? [];
  let direct: TabStop[] = [];
  try { direct = JSON.parse(String(attrs.tabStops ?? '[]')); } catch { /* 功能：损坏显示缓存按无直接制表位读取；时间：2026-11-11；作者：lq。 */ }
  if (!Array.isArray(direct)) direct = [];
  const overrides = new Set(direct.map(stop => stop.pos));
  return [...inherited.filter(stop => !overrides.has(stop.pos)), ...direct].filter(stop => stop.val !== 'clear' && !stop.rel).sort((a,b) => a.pos-b.pos);
}
/** 功能：单次提交缩进或可见制表位，移除继承位时写入 clear，保留未展示的 bar；时间：2026-11-11；作者：lq。 */
export function applyWordRuler(editor: Editor, parsed: ParsedDocFull, patch: { indentLeft?: number; indentRight?: number; indentFirstLine?: number; tabs?: TabStop[] }): boolean {
  if (!editor.isEditable || editor.view.composing) return false;
  const paragraphs = selectedParagraphs(editor);
  // 功能：正文、文本框及同一表格格内段落共用标尺事务，多格选区由 UI 明确禁用；时间：2026-11-11；作者：lq。
  const parent = paragraphs[0] && editor.state.doc.resolve(paragraphs[0].pos).parent;
  if (!Object.keys(patch).length || !paragraphs.length || paragraphs.some(({ node, pos }) => editor.state.doc.resolve(pos).parent !== parent || !['docParagraph', 'docHeading', 'docListItem'].includes(node.type.name))) return false;
  for (const [key, value] of Object.entries(patch)) {
    if (key === 'tabs') continue;
    if (!Number.isFinite(value) || Math.abs(Number(value)) > 31680 || key !== 'indentFirstLine' && Number(value) < 0) return false;
  }
  if (patch.tabs?.some(stop => !Number.isInteger(stop.pos) || stop.pos < 0 || stop.pos > 31680 || !['left','center','right','decimal'].includes(stop.val))) return false;
  const tr = closeHistory(editor.state.tr);
  for (const { node, pos } of paragraphs) {
    const attrs = { ...node.attrs };
    for (const key of ['indentLeft', 'indentRight', 'indentFirstLine'] as const) if (key in patch) attrs[key] = Math.round(patch[key]!);
    if ('indentFirstLine' in patch) attrs.indentFirstLineChars = null;
    if (patch.tabs) {
      const inherited = parsed.styles.get(String(node.attrs.styleId ?? 'Normal'))?.display?.tabStops ?? [];
      const tabs: TabStop[] = [...wordRulerTabs(node.attrs, parsed).filter(stop => stop.val === 'bar'), ...patch.tabs];
      for (const stop of inherited) if (!tabs.some(item => item.pos === stop.pos)) tabs.push({ pos: stop.pos, val: 'clear' });
      attrs.tabStops = JSON.stringify(tabs);
    }
    if (JSON.stringify(attrs) !== JSON.stringify(node.attrs)) tr.setNodeMarkup(pos, undefined, attrs);
  }
  if (!tr.docChanged) return false;
  editor.view.dispatch(tr); return true;
}
