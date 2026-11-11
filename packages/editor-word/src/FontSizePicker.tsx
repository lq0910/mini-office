/**
 * 功能：字号组合框，按 Office / WPS 的方式同时提供中文字号与数字磅值，并允许手动输入；时间：2026-10-29；作者：lq。
 *
 * 交互约定：
 * - 点击展开，面板顶部是可直接输入的数字框（回车生效），下面是“中文字号”和“数字磅值”两组快捷项；
 * - 触发器显示当前值，命中中文字号时显示“三号 16”，否则只显示磅值；
 * - 任何一次选择都只触发一次 onChange，由编辑器统一记入撤销历史。
 */
import { useEffect, useRef, useState } from 'react';
import { ChevronDown } from 'lucide-react';
import { cn } from '@mini-office/ui';
import { ToolbarPopover } from './ToolbarPopover';
import {
  CN_FONT_SIZES,
  PT_FONT_SIZES,
  cnFontSizeName,
  fontSizeLabel,
  formatPt,
  parseFontSizeInput,
} from './typography';

export function FontSizePicker({
  value,
  mixed,
  disabled,
  onChange,
}: {
  value?: number;
  /** 功能：混合字号显示空状态，不将默认字号冒充选区格式；时间：2026-11-10；作者：lq。 */
  mixed?: boolean;
  disabled?: boolean;
  onChange: (pt: number) => void;
}) {
  const [draft, setDraft] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setDraft(!mixed && value ? formatPt(value) : '');
  }, [value, mixed]);

  const commit = (raw: string, close?: () => void) => {
    const parsed = parseFontSizeInput(raw);
    if (parsed === undefined) return false;
    onChange(parsed);
    close?.();
    return true;
  };

  return (
    <ToolbarPopover
      width={214}
      disabled={disabled}
      trigger={({ open, toggle }) => (
        <button
          type="button"
          aria-label="字号"
          aria-expanded={open}
          disabled={disabled}
          onClick={toggle}
          title="字号：支持中文字号与自定义磅值"
          className={cn(
            'flex h-7 w-[92px] shrink-0 items-center justify-between gap-1 rounded border border-transparent px-2 text-[12px] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)] disabled:opacity-40',
            open && 'bg-[var(--mo-surface-2)]',
          )}
        >
          <span className="truncate">{mixed ? '混合字号' : fontSizeLabel(value)}</span>
          <ChevronDown size={12} className="shrink-0 text-[var(--mo-text-muted)]" />
        </button>
      )}
    >
      {({ close }) => (
        <div className="space-y-1" data-testid="font-size-panel">
          <div className="flex items-center gap-1 px-1 pb-1">
            <input
              ref={inputRef}
              autoFocus
              value={draft}
              aria-label="字号磅值"
              placeholder="输入磅值或字号"
              onChange={(event) => setDraft(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') {
                  event.preventDefault();
                  commit(draft, close);
                }
              }}
              className="h-7 min-w-0 flex-1 rounded border border-[var(--mo-line)] bg-[var(--mo-surface)] px-2 text-[12px] tabular-nums outline-none focus:border-[var(--mo-accent)]"
            />
            <button
              type="button"
              onClick={() => commit(draft, close)}
              className="h-7 shrink-0 rounded bg-[var(--mo-accent)] px-2 text-[11px] font-medium text-[var(--mo-accent-ink)] hover:opacity-90"
            >
              应用
            </button>
          </div>
          <div className="px-2 pb-0.5 pt-1 text-[10px] font-semibold uppercase tracking-wide text-[var(--mo-text-faint)]">中文字号</div>
          <div className="grid grid-cols-2 gap-0.5">
            {CN_FONT_SIZES.map((item) => (
              <button
                key={item.name}
                type="button"
                onClick={() => {
                  onChange(item.pt);
                  close();
                }}
                className={cn(
                  'flex h-7 items-center justify-between rounded px-2 text-[11px] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]',
                  !mixed && value !== undefined && Math.abs(value - item.pt) < 0.01 && 'bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]',
                )}
              >
                <span>{item.name}</span>
                <span className="tabular-nums text-[10px] text-[var(--mo-text-faint)]">{formatPt(item.pt)}</span>
              </button>
            ))}
          </div>
          <div className="mx-1 my-1 h-px bg-[var(--mo-line)]" />
          <div className="px-2 pb-0.5 text-[10px] font-semibold uppercase tracking-wide text-[var(--mo-text-faint)]">数字磅值</div>
          <div className="grid grid-cols-3 gap-0.5">
            {PT_FONT_SIZES.map((pt) => (
              <button
                key={pt}
                type="button"
                onClick={() => {
                  onChange(pt);
                  close();
                }}
                className={cn(
                  'h-7 rounded text-[11px] tabular-nums text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]',
                  !mixed && value !== undefined && Math.abs(value - pt) < 0.01 && 'bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]',
                )}
              >
                {formatPt(pt)}
              </button>
            ))}
          </div>
          {!mixed && value && cnFontSizeName(value) ? (
            <p className="px-2 pt-1 text-[10px] text-[var(--mo-text-faint)]">
              当前：{cnFontSizeName(value)} · {formatPt(value)} 磅
            </p>
          ) : null}
        </div>
      )}
    </ToolbarPopover>
  );
}
