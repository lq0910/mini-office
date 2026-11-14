/** 功能：将现有查找浮层接到当前 Word 会话的原生位置、高亮和替换历史；时间：2026-11-10；作者：lq。 */
import { useEffect, useMemo, useState } from 'react';
import type { WordCoreSession } from '../../word-core/src/index';
import { findWordText, highlightWordSearch, replaceWordText } from '../../word-core/src/editing/index';
import { FindBar } from './FindBar';
import './core-search.css';

/** 功能：焦点进入输入框后继续搜索当前活动编辑区，正文变化时丢弃旧位置并重新扫描；时间：2026-11-10；作者：lq。 */
export function WordCoreFindBar({ core, readonly, onClose }: { core: WordCoreSession; readonly: boolean; onClose: () => void }) {
  const [query, setQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [index, setIndex] = useState(0);
  const [, refresh] = useState(0);
  useEffect(() => core.subscribe(() => refresh(value => value + 1)), [core]);
  const editor = core.commandEditor();
  const doc = editor.state.doc;
  const results = useMemo(() => findWordText(editor, query, matchCase), [editor, doc, query, matchCase]);
  const active = Math.min(index, Math.max(0, results.matches.length - 1));
  useEffect(() => { highlightWordSearch(editor, results, active); }, [editor, results, active]);
  // 功能：关闭查找、切换文档或活动编辑区时移除旧高亮，不改变正文和 dirty 状态；时间：2026-11-10；作者：lq。
  useEffect(() => () => { if (!editor.isDestroyed) highlightWordSearch(editor, findWordText(editor, ''), 0); }, [editor]);
  /** 功能：导航只改变原生选区并滚动到结果，输入焦点保留在查找框；时间：2026-11-10；作者：lq。 */
  const navigate = (delta: number) => {
    if (!results.matches.length) return;
    const next = (active + delta + results.matches.length) % results.matches.length;
    setIndex(next);
    const match = results.matches[next];
    editor.commands.setTextSelection({ from: match.from, to: match.to });
    editor.view.dispatch(editor.state.tr.scrollIntoView());
  };
  /** 功能：替换遵守只读/组合输入并读取最新结果；整批替换保持一步撤销；时间：2026-11-10；作者：lq。 */
  const replace = (all = false) => {
    // 功能：全文替换等待来源加载完整，避免仅替换首屏后错误报告完成；时间：2026-11-13；作者：lq。
    if (!readonly && !core.contentPending()) replaceWordText(editor, results, replacement, all ? undefined : active);
  };
  return <FindBar query={query} replacement={replacement} caseSensitive={matchCase} count={results.matches.length} active={active} readonly={readonly || core.contentPending()}
    onQuery={value => { setQuery(value); setIndex(0); }} onReplacement={setReplacement}
    onCaseSensitive={value => { setMatchCase(value); setIndex(0); }} onNext={() => navigate(1)} onPrev={() => navigate(-1)}
    onReplace={() => replace()} onReplaceAll={() => replace(true)} onClose={() => { onClose(); editor.view.focus(); }} />;
}
