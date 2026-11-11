/** 功能：新旧 Word 核心共用现有查找/替换浮层，补齐读屏标签与组合输入保护；时间：2026-11-10；作者：lq。 */
import { useEffect, useRef } from 'react';
import { Search, Replace, ChevronUp, ChevronDown, X } from 'lucide-react';
import { Input, IconButton, Button } from '@mini-office/ui';

export function FindBar({
  query,
  replacement,
  caseSensitive,
  count,
  active,
  readonly,
  onQuery,
  onReplacement,
  onCaseSensitive,
  onNext,
  onPrev,
  onReplace,
  onReplaceAll,
  onClose,
}: {
  query: string;
  replacement: string;
  caseSensitive: boolean;
  count: number;
  active: number;
  readonly: boolean;
  onQuery: (value: string) => void;
  onReplacement: (value: string) => void;
  onCaseSensitive: (value: boolean) => void;
  onNext: () => void;
  onPrev: () => void;
  onReplace: () => void;
  onReplaceAll: () => void;
  onClose: () => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  useEffect(() => {
    inputRef.current?.focus();
  }, []);
  return (
    <div role="search" aria-label="查找替换" className="mo-pop-in absolute right-4 top-3 z-30 w-[330px] rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface)] p-2.5 shadow-pop">
      <div className="flex items-center gap-1.5">
        <div className="relative flex-1">
          <Search size={13} className="pointer-events-none absolute left-2 top-1/2 -translate-y-1/2 text-[var(--mo-text-faint)]" />
          <Input
            ref={inputRef}
            value={query}
            onChange={(event) => onQuery(event.target.value)}
            onKeyDown={(event) => {
              if (event.nativeEvent.isComposing || event.keyCode === 229) return;
              if (event.key === 'Enter') {
                event.preventDefault();
                if (event.shiftKey) onPrev();
                else onNext();
              }
              if (event.key === 'Escape') onClose();
            }}
            aria-label="查找内容"
            placeholder="查找内容"
            className="h-7 pl-7 text-[12px]"
          />
        </div>
        <span role="status" aria-label="查找结果" className="w-14 whitespace-nowrap text-center text-[10px] tabular-nums text-[var(--mo-text-muted)]">
          {count > 0 ? `${active + 1}/${count}` : '0/0'}
        </span>
        <IconButton label="上一个" onClick={onPrev} disabled={count === 0}>
          <ChevronUp size={14} />
        </IconButton>
        <IconButton label="下一个" onClick={onNext} disabled={count === 0}>
          <ChevronDown size={14} />
        </IconButton>
        <IconButton label="关闭" onClick={onClose}>
          <X size={14} />
        </IconButton>
      </div>
      {!readonly ? (
        <div className="mt-2 flex items-center gap-1.5">
          <div className="relative flex-1">
            <Replace size={13} className="pointer-events-none absolute left-2 top-1/2 -translate-y-1/2 text-[var(--mo-text-faint)]" />
            <Input
              value={replacement}
              onChange={(event) => onReplacement(event.target.value)}
              onKeyDown={(event) => {
                if (event.nativeEvent.isComposing || event.keyCode === 229) return;
              if (event.key === 'Enter') {
                  event.preventDefault();
                  onReplace();
                }
                if (event.key === 'Escape') onClose();
              }}
              aria-label="替换为"
              placeholder="替换为"
              className="h-7 pl-7 text-[12px]"
            />
          </div>
          <Button size="sm" onClick={onReplace} disabled={count === 0}>
            替换
          </Button>
          <Button size="sm" onClick={onReplaceAll} disabled={count === 0}>
            全部替换
          </Button>
        </div>
      ) : null}
      <label className="mt-2 flex cursor-pointer items-center gap-1.5 text-[11px] text-[var(--mo-text-muted)]">
        <input
          type="checkbox"
          checked={caseSensitive}
          onChange={(event) => onCaseSensitive(event.target.checked)}
          className="h-3 w-3 accent-[var(--mo-accent)]"
        />
        区分大小写
      </label>
    </div>
  );
}
