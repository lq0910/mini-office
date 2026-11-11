/**
 * 功能：行距组合框，提供倍数预设与固定值/最小值的自由磅值输入；时间：2026-10-29；作者：lq。
 *
 * 旧实现只有“固定 18pt”一个死值，用户无法自定义固定行距。这里把三种行距模式
 * （多倍、固定值、最小值）都暴露出来，并保留进入完整段落设置的入口。
 * 精确行距只保留一个输入框 + 一个“应用”，模式用分段按钮切换，避免两行各带一个“应用”造成歧义。
 */
import { useEffect, useState } from 'react';
import { ChevronDown } from 'lucide-react';
import type { ParagraphStyle } from '@mini-office/document-model';
import { cn } from '@mini-office/ui';
import { ToolbarPopover } from './ToolbarPopover';
import { formatPt, lineSpacingLabel, readLineSpacing, writeLineSpacing } from './typography';

const MULTIPLES = [1, 1.15, 1.5, 2, 2.5, 3];
const PRECISE_MODES = [
  { mode: 'exact' as const, label: '固定值' },
  { mode: 'atLeast' as const, label: '最小值' },
];

export function LineSpacingPicker({
  style,
  mixed,
  disabled,
  onChange,
  onOpenOptions,
}: {
  style: ParagraphStyle | undefined;
  /** 功能：混合行距不误显示默认倍数或勾选预设；时间：2026-11-10；作者：lq。 */
  mixed?: boolean;
  disabled?: boolean;
  onChange: (patch: Partial<ParagraphStyle>) => void;
  onOpenOptions?: () => void;
}) {
  const state = readLineSpacing(style);
  const [preciseMode, setPreciseMode] = useState<'exact' | 'atLeast'>(
    state.mode === 'atLeast' ? 'atLeast' : 'exact',
  );
  const [draft, setDraft] = useState('');

  useEffect(() => {
    setPreciseMode(state.mode === 'atLeast' ? 'atLeast' : 'exact');
    setDraft(state.mode === 'multiple' ? '18' : formatPt(state.value));
    // 只在外部值变化时同步输入框，避免输入过程中被覆盖。
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [state.mode, state.value]);

  const apply = (close?: () => void) => {
    const value = Number(draft);
    if (!Number.isFinite(value) || value <= 0) return;
    onChange(writeLineSpacing({ mode: preciseMode, value }));
    close?.();
  };

  return (
    <ToolbarPopover
      width={232}
      disabled={disabled}
      trigger={({ open, toggle }) => (
        <button
          type="button"
          aria-label="行距"
          aria-expanded={open}
          disabled={disabled}
          onClick={toggle}
          title="行距：倍数、固定值或最小值"
          className={cn(
            'flex h-7 w-[104px] shrink-0 items-center justify-between gap-1 rounded border border-transparent px-2 text-[12px] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)] disabled:opacity-40',
            open && 'bg-[var(--mo-surface-2)]',
          )}
        >
          <span className="truncate">{mixed ? '混合行距' : lineSpacingLabel(state)}</span>
          <ChevronDown size={12} className="shrink-0 text-[var(--mo-text-muted)]" />
        </button>
      )}
    >
      {({ close }) => (
        <div className="space-y-1">
          <div className="px-2 pb-0.5 pt-1 text-[10px] font-semibold uppercase tracking-wide text-[var(--mo-text-faint)]">倍数行距</div>
          <div className="grid grid-cols-3 gap-0.5">
            {MULTIPLES.map((multiple) => (
              <button
                key={multiple}
                type="button"
                onClick={() => {
                  onChange(writeLineSpacing({ mode: 'multiple', value: multiple }));
                  close();
                }}
                className={cn(
                  'h-7 rounded text-[11px] tabular-nums text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]',
                  !mixed && state.mode === 'multiple' && Math.abs(state.value - multiple) < 0.01 && 'bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]',
                )}
              >
                {formatPt(multiple)}
              </button>
            ))}
          </div>
          <div className="mx-1 my-1 h-px bg-[var(--mo-line)]" />
          <div className="px-2 pb-0.5 text-[10px] font-semibold uppercase tracking-wide text-[var(--mo-text-faint)]">精确行距</div>
          <div className="mx-1 flex gap-0.5 rounded bg-[var(--mo-surface-2)] p-0.5">
            {PRECISE_MODES.map((item) => (
              <button
                key={item.mode}
                type="button"
                aria-pressed={preciseMode === item.mode}
                onClick={() => setPreciseMode(item.mode)}
                className={cn(
                  'h-6 flex-1 rounded text-[11px] transition-colors',
                  preciseMode === item.mode
                    ? 'bg-[var(--mo-surface-raised)] text-[var(--mo-text-primary)] shadow-sm'
                    : 'text-[var(--mo-text-muted)] hover:text-[var(--mo-text-primary)]',
                )}
              >
                {item.label}
              </button>
            ))}
          </div>
          <div className="flex items-center gap-1 px-1">
            <input
              value={draft}
              aria-label={preciseMode === 'exact' ? '固定值行距' : '最小值行距'}
              inputMode="decimal"
              onChange={(event) => setDraft(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') {
                  event.preventDefault();
                  apply(close);
                }
              }}
              className="h-7 min-w-0 flex-1 rounded border border-[var(--mo-line)] bg-[var(--mo-surface)] px-1.5 text-[11px] tabular-nums outline-none focus:border-[var(--mo-accent)]"
            />
            <span className="shrink-0 text-[10px] text-[var(--mo-text-faint)]">磅</span>
            <button
              type="button"
              onClick={() => apply(close)}
              className="h-7 shrink-0 rounded bg-[var(--mo-accent)] px-2 text-[11px] font-medium text-[var(--mo-accent-ink)] hover:opacity-90"
            >
              应用
            </button>
          </div>
          {onOpenOptions ? (
            <>
              <div className="mx-1 my-1 h-px bg-[var(--mo-line)]" />
              <button
                type="button"
                onClick={() => {
                  close();
                  onOpenOptions();
                }}
                className="flex h-7 w-full items-center rounded px-2 text-left text-[11px] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]"
              >
                行距与缩进选项…
              </button>
            </>
          ) : null}
        </div>
      )}
    </ToolbarPopover>
  );
}
