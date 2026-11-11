/** 功能：使用新核心位置查找跨格式文字，并在单个原生事务中替换；时间：2026-11-10；作者：lq。 */
import type { Editor } from '@tiptap/core';
import type { Node } from '@tiptap/pm/model';
import { closeHistory } from '@tiptap/pm/history';
import { Plugin } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { findInDoc, type FindMatch } from '../upstream/apps/docs/src/renderer/editor/find';
import { searchPluginKey } from '../upstream/apps/docs/src/renderer/editor/decoration-extensions';

/** 功能：结果携带不可变文档身份，拒绝把旧查找位置用于新正文；时间：2026-11-10；作者：lq。 */
export interface WordSearchResults { doc: Node; matches: FindMatch[]; }

/** 功能：沿用产品的文字/大小写查找语义，不把用户输入当作正则或特殊替换语法；时间：2026-11-10；作者：lq。 */
export function findWordText(editor: Editor, query: string, matchCase = false): WordSearchResults {
  const doc = editor.state.doc;
  const matches = query ? findInDoc(doc, { mode: 'literal', re: null, needle: query, anchorStart: false, anchorEnd: false }, { matchCase, wholeWord: false, matchWidth: true }) : [];
  return { doc, matches };
}

/** 功能：高亮仅是显示事务，不进入撤销或未保存修订；大量命中只装饰活动项；时间：2026-11-10；作者：lq。 */
export function highlightWordSearch(editor: Editor, results: WordSearchResults, active: number): void {
  if (editor.isDestroyed || editor.state.doc !== results.doc) return;
  // 功能：文本框精简编辑器按需采用同一高亮协议，避免查得到但没有可见命中；时间：2026-11-10；作者：lq。
  if (!searchPluginKey.getState(editor.state)) editor.registerPlugin(new Plugin<DecorationSet>({
    key: searchPluginKey,
    state: {
      init: () => DecorationSet.empty,
      apply(tr, old) {
        const hit = tr.getMeta(searchPluginKey) as { ranges: Array<{ from: number; to: number }>; activeIndex: number } | undefined;
        return hit ? DecorationSet.create(tr.doc, hit.ranges.map((range, i) => Decoration.inline(range.from, range.to, { class: i === hit.activeIndex ? 'search-hit search-hit-active' : 'search-hit' }))) : old.map(tr.mapping, tr.doc);
      },
    },
    props: { decorations(state) { return this.getState(state); } },
  }));
  const ranges = results.matches.length > 2000 ? results.matches.slice(active, active + 1) : results.matches;
  editor.view.dispatch(editor.state.tr.setMeta(searchPluginKey, { ranges, activeIndex: results.matches.length > 2000 ? 0 : active }).setMeta('addToHistory', false));
}

/** 功能：从后向前替换以保持命中位置稳定，替换文字继承命中首字符格式，空字符串代表删除；时间：2026-11-10；作者：lq。 */
export function replaceWordText(editor: Editor, results: WordSearchResults, replacement: string, active?: number): number {
  if (!editor.isEditable || editor.view.composing || editor.state.doc !== results.doc) return 0;
  const matches = active === undefined ? results.matches : results.matches.slice(active, active + 1);
  if (!matches.length) return 0;
  const tr = closeHistory(editor.state.tr);
  for (const match of [...matches].reverse()) {
    const marks = results.doc.resolve(match.from).nodeAfter?.marks ?? [];
    if (replacement) tr.replaceWith(match.from, match.to, editor.schema.text(replacement, marks));
    else tr.delete(match.from, match.to);
  }
  editor.view.dispatch(tr);
  return matches.length;
}
