import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type ReactNode,
} from 'react';
import { createPortal } from 'react-dom';
import { Check } from 'lucide-react';
import { cn } from './primitives';

export interface MenuItem {
  id?: string;
  label?: ReactNode;
  icon?: ReactNode;
  shortcut?: string;
  onSelect?: () => void;
  disabled?: boolean;
  danger?: boolean;
  checked?: boolean;
  type?: 'item' | 'separator' | 'label';
}

export function MenuList({
  items,
  onClose,
  className,
  minWidth = 200,
}: {
  items: MenuItem[];
  onClose?: () => void;
  className?: string;
  minWidth?: number;
}) {
  return (
    <div
      role="menu"
      style={{ minWidth }}
      className={cn(
        'mo-glass mo-pop-in overflow-hidden rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface-raised)] p-1 shadow-pop',
        className,
      )}
    >
      {items.map((item, index) => {
        if (item.type === 'separator') {
          return <div key={`sep-${index}`} className="my-1 h-px bg-[var(--mo-line)]" />;
        }
        if (item.type === 'label') {
          return (
            <div
              key={`label-${index}`}
              className="px-2 pb-1 pt-1.5 text-[10px] font-semibold uppercase tracking-wide text-[var(--mo-text-faint)]"
            >
              {item.label}
            </div>
          );
        }
        return (
          <button
            key={item.id ?? `item-${index}`}
            type="button"
            role="menuitem"
            disabled={item.disabled}
            onClick={() => {
              if (item.disabled) return;
              item.onSelect?.();
              onClose?.();
            }}
            className={cn(
              'flex h-8 w-full items-center gap-2 rounded-[5px] px-2 text-left text-[12px] transition-colors duration-fast',
              item.danger
                ? 'text-[var(--mo-danger)] hover:bg-[var(--mo-danger-soft)]'
                : 'text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]',
              item.disabled && 'pointer-events-none opacity-40',
            )}
          >
            <span className="flex h-4 w-4 shrink-0 items-center justify-center text-[var(--mo-text-muted)]">
              {item.checked ? <Check size={13} /> : item.icon}
            </span>
            <span className="min-w-0 flex-1 truncate">{item.label}</span>
            {item.shortcut ? (
              <span className="shrink-0 text-[10px] text-[var(--mo-text-faint)]">{item.shortcut}</span>
            ) : null}
          </button>
        );
      })}
    </div>
  );
}

export function Dropdown({
  trigger,
  items,
  align = 'left',
  minWidth = 200,
  className,
}: {
  trigger: (props: { open: boolean; toggle: () => void }) => ReactNode;
  items: MenuItem[];
  align?: 'left' | 'right';
  minWidth?: number;
  className?: string;
}) {
  const [open, setOpen] = useState(false);
  const anchorRef = useRef<HTMLDivElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState({ left: 0, top: 0 });

  const updatePosition = useCallback(() => {
    const anchor = anchorRef.current;
    if (!anchor) return;
    const rect = anchor.getBoundingClientRect();
    const width = menuRef.current?.offsetWidth ?? minWidth;
    const height = menuRef.current?.offsetHeight ?? 240;
    const margin = 8;
    let left = align === 'right' ? rect.right - width : rect.left;
    left = Math.min(Math.max(margin, left), window.innerWidth - width - margin);
    let top = rect.bottom + 4;
    if (top + height > window.innerHeight - margin) {
      top = Math.max(margin, rect.top - height - 4);
    }
    setPosition({ left, top });
  }, [align, minWidth]);

  useLayoutEffect(() => {
    if (!open) return;
    updatePosition();
    const frame = requestAnimationFrame(updatePosition);
    return () => cancelAnimationFrame(frame);
  }, [open, updatePosition]);

  useEffect(() => {
    if (!open) return;
    const onPointerDown = (event: PointerEvent) => {
      const target = event.target as Node;
      if (anchorRef.current?.contains(target) || menuRef.current?.contains(target)) return;
      setOpen(false);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setOpen(false);
    };
    const reposition = () => updatePosition();
    document.addEventListener('pointerdown', onPointerDown, true);
    document.addEventListener('keydown', onKeyDown);
    window.addEventListener('resize', reposition);
    window.addEventListener('scroll', reposition, true);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true);
      document.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('resize', reposition);
      window.removeEventListener('scroll', reposition, true);
    };
  }, [open, updatePosition]);

  return (
    <div ref={anchorRef} className={cn('relative inline-flex', className)}>
      {trigger({ open, toggle: () => setOpen((value) => !value) })}
      {open
        ? createPortal(
            <div ref={menuRef} className="fixed z-[95]" style={{ left: position.left, top: position.top }}>
              <MenuList items={items} onClose={() => setOpen(false)} minWidth={minWidth} />
            </div>,
            document.body,
          )
        : null}
    </div>
  );
}

/** 功能：上下文菜单在窄/矮窗口保留边缘并允许滚动到全部命令；时间：2026-11-11；作者：lq。 */
export function ContextMenu({
  x,
  y,
  items,
  onClose,
  minWidth = 190,
}: {
  x: number;
  y: number;
  items: MenuItem[];
  onClose: () => void;
  minWidth?: number;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState({ left: x, top: y });

  useEffect(() => {
    const width = ref.current?.offsetWidth ?? minWidth;
    const height = ref.current?.offsetHeight ?? 200;
    setPosition({
      left: Math.max(8, Math.min(x, window.innerWidth - width - 8)),
      top: Math.max(8, Math.min(y, window.innerHeight - height - 8)),
    });
  }, [x, y, minWidth]);

  useEffect(() => {
    const onPointerDown = (event: PointerEvent) => {
      if (ref.current && !ref.current.contains(event.target as Node)) onClose();
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') onClose();
    };
    document.addEventListener('pointerdown', onPointerDown, true);
    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true);
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [onClose]);

  if (typeof document === 'undefined') return null;

  return createPortal(
    <div
      ref={ref}
      className="fixed z-[90]"
      style={{ left: position.left, top: position.top, maxHeight: 'calc(100vh - 16px)', maxWidth: 'calc(100vw - 16px)', overflow: 'auto' }}
      onContextMenu={(event) => event.preventDefault()}
    >
      <MenuList items={items} onClose={onClose} minWidth={minWidth} />
    </div>,
    document.body,
  );
}
