/** 功能：沿用 Mini Office 三按钮粘贴选项，操作当前编辑区的唯一正文；时间：2026-11-11；作者：lq。 */
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import { applyWordPasteMode, dismissWordPaste, wordPasteRecord } from '../../word-core/src/editing/paste';
export function WordCorePasteOptions({ core }: { core: WordCoreSession }) {
  const editor = core.commandEditor(); const record = wordPasteRecord(editor);
  if (!record) return null;
  return <div role="group" aria-label="粘贴选项" className="mo-no-print fixed bottom-12 right-6 z-[90] rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface)] p-2 shadow-pop" onMouseDown={event => event.preventDefault()}>
    <div className="mb-1 flex items-center justify-between px-2 text-[11px] text-[var(--mo-text-muted)]">粘贴选项<button type="button" aria-label="关闭粘贴选项" onClick={() => dismissWordPaste(editor)}>×</button></div>
    {([['source', '保留源格式'], ['merge', '匹配目标格式'], ['text', '仅保留文本']] as const).map(([mode, label]) => <button key={mode} type="button" aria-pressed={record.mode === mode} onClick={() => applyWordPasteMode(editor, mode)} className="flex w-full items-center gap-3 rounded-md px-3 py-2 text-left text-[12px] hover:bg-[var(--mo-surface-2)]"><span className="w-3 text-[var(--mo-accent)]">{record.mode === mode ? '✓' : ''}</span>{label}</button>)}
  </div>;
}
