import { useEffect, useId, useLayoutEffect, useRef, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { X } from 'lucide-react';
import { cn } from './primitives';

export interface DialogProps {
  open: boolean;
  onClose: () => void;
  title: string;
  description?: ReactNode;
  icon?: ReactNode;
  children: ReactNode;
  footer?: ReactNode;
  width?: 'sm' | 'md' | 'lg' | 'xl' | '2xl' | 'print';
  closeOnBackdrop?: boolean;
  className?: string;
  /** 覆盖内容区样式；配合 bodyScroll=false 可让内容区自行管理滚动。 */
  bodyClassName?: string;
  /** 是否由弹窗内容区统一滚动，默认 true。 */
  bodyScroll?: boolean;
}

const WIDTHS = {
  sm: 'max-w-[400px]',
  md: 'max-w-[500px]',
  lg: 'max-w-[660px]',
  xl: 'max-w-[820px]',
  '2xl': 'max-w-[940px]',
  // 功能：打印预览需要更大的版面才能看清细节；时间：2026-11-01；作者：lq。
  print: 'max-w-[min(1240px,96vw)]',
};

export function Dialog({
  open,
  onClose,
  title,
  description,
  icon,
  children,
  footer,
  width = 'md',
  closeOnBackdrop = true,
  className,
  bodyClassName,
  bodyScroll = true,
}: DialogProps) {
  const titleId = useId();
  const panelRef = useRef<HTMLDivElement>(null);
  // 功能：保存最新关闭行为，避免表单输入生成新回调时重新执行弹窗聚焦；时间：2026-10-19；作者：lq。
  const onCloseRef = useRef(onClose);
  useLayoutEffect(() => {
    onCloseRef.current = onClose;
  }, [onClose]);

  useEffect(() => {
    if (!open) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.stopPropagation();
        onCloseRef.current();
      }
      if (event.key === 'Tab' && panelRef.current) {
        const focusables = panelRef.current.querySelectorAll<HTMLElement>(
          'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])',
        );
        if (focusables.length === 0) return;
        const first = focusables[0];
        const last = focusables[focusables.length - 1];
        if (!event.shiftKey && document.activeElement === last) {
          event.preventDefault();
          first.focus();
        } else if (event.shiftKey && document.activeElement === first) {
          event.preventDefault();
          last.focus();
        }
      }
    };
    document.addEventListener('keydown', onKeyDown, true);
    const timer = setTimeout(() => {
      // 功能：只在打开时初始化焦点，用户已点击表单时不抢走输入焦点；时间：2026-10-19；作者：lq。
      const panel = panelRef.current;
      if (!panel || panel.contains(document.activeElement)) return;
      const preferred = panel.querySelector<HTMLElement>('[data-autofocus]');
      (preferred ?? panel).focus({ preventScroll: true });
    }, 30);
    return () => {
      document.removeEventListener('keydown', onKeyDown, true);
      clearTimeout(timer);
    };
  }, [open]);

  if (!open || typeof document === 'undefined') return null;

  return createPortal(
    <div className="fixed inset-0 z-[70] flex items-end justify-center p-3 sm:items-center sm:p-6">
      <div
        className="mo-scrim mo-fade-in"
        onClick={closeOnBackdrop ? onClose : undefined}
        aria-hidden
      />
      <div
        ref={panelRef}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        tabIndex={-1}
        className={cn(
          'mo-glass mo-pop-in relative z-[71] flex max-h-[calc(100dvh-24px)] w-full flex-col overflow-hidden rounded-2xl border border-[var(--mo-line-strong)] bg-[var(--mo-surface-raised)] shadow-[0_28px_80px_-28px_rgba(16,24,40,0.5),0_10px_30px_-16px_rgba(16,24,40,0.28)] outline-none sm:max-h-[calc(100dvh-48px)]',
          WIDTHS[width],
          className,
        )}
      >
        <div className="flex items-start gap-3.5 px-5 pb-3.5 pt-4 sm:px-6 sm:pt-5">
          {icon ? (
            <div className="mt-0.5 flex h-9 w-9 shrink-0 items-center justify-center rounded-[11px] border border-[color-mix(in_srgb,var(--mo-accent)_12%,transparent)] bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]">
              {icon}
            </div>
          ) : null}
          <div className="min-w-0 flex-1">
            <h2 id={titleId} className="text-[16px] font-semibold leading-6 tracking-[-0.01em] text-[var(--mo-text-primary)]">
              {title}
            </h2>
            {description ? (
              <div className="mt-1 max-w-[60ch] text-[12px] leading-5 text-[var(--mo-text-muted)]">{description}</div>
            ) : null}
          </div>
          <button
            type="button"
            aria-label="关闭"
            data-tooltip="关闭"
            onClick={onClose}
            className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full text-[var(--mo-text-muted)] transition-colors duration-fast hover:bg-[var(--mo-surface-2)] hover:text-[var(--mo-text-primary)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--mo-accent)]/40"
          >
            <X size={16} />
          </button>
        </div>
        <div
          className={cn(
            'min-h-0 flex-1 text-[13px]',
            bodyScroll ? 'mo-scroll overflow-x-hidden px-5 pb-5 sm:px-6' : 'flex overflow-hidden',
            bodyClassName,
          )}
        >
          {children}
        </div>
        {footer ? (
          <div className="flex items-center justify-end gap-2 border-t border-[var(--mo-line)] bg-[var(--mo-surface-2)]/50 px-5 py-3.5 sm:px-6">
            {footer}
          </div>
        ) : null}
      </div>
    </div>,
    document.body,
  );
}
