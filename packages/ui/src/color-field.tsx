import {
  type ReactNode,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import { createPortal } from 'react-dom';
import { Ban, Check, ChevronDown } from 'lucide-react';
import { cn } from './primitives';

const PRESET_COLORS = [
  '#000000',
  '#1D2939',
  '#475467',
  '#98A2B3',
  '#FFFFFF',
  '#D92D20',
  '#F04438',
  '#F79009',
  '#12B76A',
  '#0BA5EC',
  '#3478F6',
  '#7A5AF8',
  '#EE46BC',
  '#B42318',
  '#027A48',
  '#085D9E',
];

const PANEL_WIDTH = 206;
const PANEL_MARGIN = 8;

/**
 * 功能：记住每个颜色入口上次使用的颜色，点击主按钮即可直接套用，无需每次重新选色；
 * 时间：2026-11-01；作者：lq。
 */
const lastColorByField = new Map<string, string>();

export function ColorField({
  value,
  onChange,
  label,
  allowClear = true,
  className,
  icon,
  indicatorColor,
  disabled = false,
  applyOnClick = true,
}: {
  value?: string;
  onChange: (color: string | undefined) => void;
  label?: string;
  allowClear?: boolean;
  className?: string;
  /** 功能：工具栏图标及当前颜色标识，兼容原有色值展示；时间：2026-10-19；作者：lq。 */
  icon?: ReactNode;
  indicatorColor?: string;
  disabled?: boolean;
  /** 功能：主按钮直接套用上次颜色（Excel 行为），false 时仍打开调色板；时间：2026-11-01；作者：lq。 */
  applyOnClick?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const buttonRef = useRef<HTMLButtonElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState({ left: 0, top: 0 });
  const fieldKey = label ?? 'default';
  const remembered = lastColorByField.get(fieldKey);
  const activeColor = remembered ?? value;

  const updatePosition = () => {
    const button = buttonRef.current;
    if (!button) return;
    const rect = button.getBoundingClientRect();
    const panelHeight = panelRef.current?.offsetHeight ?? 142;
    let left = rect.right - PANEL_WIDTH;
    left = Math.min(Math.max(PANEL_MARGIN, left), window.innerWidth - PANEL_WIDTH - PANEL_MARGIN);
    let top = rect.bottom + 6;
    if (top + panelHeight > window.innerHeight - PANEL_MARGIN) {
      top = Math.max(PANEL_MARGIN, rect.top - panelHeight - 6);
    }
    setPosition({ left, top });
  };

  useLayoutEffect(() => {
    if (!open) return;
    updatePosition();
    const frame = requestAnimationFrame(updatePosition);
    return () => cancelAnimationFrame(frame);
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const onPointerDown = (event: PointerEvent) => {
      const target = event.target as Node;
      if (panelRef.current?.contains(target) || buttonRef.current?.contains(target)) return;
      setOpen(false);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setOpen(false);
    };
    const close = () => setOpen(false);
    document.addEventListener('pointerdown', onPointerDown, true);
    document.addEventListener('keydown', onKeyDown);
    window.addEventListener('resize', close);
    window.addEventListener('scroll', close, true);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true);
      document.removeEventListener('keydown', onKeyDown);
      window.removeEventListener('resize', close);
      window.removeEventListener('scroll', close, true);
    };
  }, [open]);

  const choose = (color: string | undefined) => {
    if (color) lastColorByField.set(fieldKey, color);
    onChange(color);
    setOpen(false);
  };

  // 功能：主按钮直接套用上次颜色，没有历史颜色时才展开调色板；时间：2026-11-01；作者：lq。
  const applyLastColor = () => {
    if (!applyOnClick) { setOpen((current) => !current); return; }
    if (activeColor) onChange(activeColor);
    else setOpen(true);
  };

  return (
    <div className={cn('relative inline-flex shrink-0 items-stretch', className)}>
      <div
        className={cn(
          'flex h-7 shrink-0 items-stretch overflow-hidden rounded-md border bg-[var(--mo-surface)] text-[11px] text-[var(--mo-text-primary)] transition-colors duration-fast',
          open
            ? 'border-[var(--mo-accent)] ring-2 ring-[color-mix(in_srgb,var(--mo-accent)_16%,transparent)]'
            : 'border-[var(--mo-line-strong)]',
          disabled ? 'pointer-events-none opacity-40' : 'hover:bg-[var(--mo-surface-2)]',
        )}
      >
        <button
          ref={buttonRef}
          type="button"
          aria-label={label ?? '选择颜色'}
          disabled={disabled}
          data-tooltip={label ?? '选择颜色'}
          onClick={applyLastColor}
          className="flex items-center gap-1.5 px-1.5"
        >
          {icon ? (
            <span aria-hidden="true" className="flex w-4 flex-col items-center gap-0.5">
              {icon}
              <span className="h-[3px] w-4 rounded-sm border border-[var(--mo-line)]"
                style={{ background: indicatorColor ?? value ?? 'transparent' }} />
            </span>
          ) : <>
            <span
              className="h-4 w-4 rounded-sm border border-[var(--mo-line)]"
              style={{ background: value ?? 'transparent' }}
            />
            <span className="w-[48px] truncate font-mono text-[10px]">{value ?? '自动'}</span>
          </>}
        </button>
        <button
          type="button"
          aria-label={`${label ?? '颜色'}调色板`}
          data-tooltip="更多颜色"
          aria-haspopup="dialog"
          aria-expanded={open}
          disabled={disabled}
          onClick={() => setOpen((current) => !current)}
          className="flex w-4 items-center justify-center border-l border-[var(--mo-line)] text-[var(--mo-text-muted)]"
        >
          <ChevronDown size={10} className="opacity-70" />
        </button>
      </div>
      {open && !disabled
        ? createPortal(
            <div
              ref={panelRef}
              role="dialog"
              aria-label={label ?? '选择颜色'}
              className="mo-pop-in fixed z-[100] rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface-raised)] p-2 shadow-pop"
              style={{ left: position.left, top: position.top, width: PANEL_WIDTH }}
            >
              <div className="grid grid-cols-8 gap-1">
                {PRESET_COLORS.map((color) => (
                  <button
                    key={color}
                    type="button"
                    aria-label={color}
                    data-tooltip={color}
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={() => choose(color)}
                    className="flex h-5 w-5 items-center justify-center rounded-[4px] border border-[var(--mo-line)] transition-transform hover:scale-110"
                    style={{ background: color }}
                  >
                    {value?.toUpperCase() === color ? (
                      <Check
                        size={11}
                        className={color === '#FFFFFF' ? 'text-black' : 'text-white'}
                      />
                    ) : null}
                  </button>
                ))}
              </div>
              <div className="mt-2 flex items-center gap-1.5">
                <input
                  type="color"
                  aria-label="自定义颜色"
                  data-tooltip="自定义颜色"
                  value={value ?? '#000000'}
                  onChange={(event) => onChange(event.target.value.toUpperCase())}
                  className="h-7 w-8 cursor-pointer rounded border border-[var(--mo-line)] bg-transparent p-0.5"
                />
                <input
                  value={value ?? ''}
                  onChange={(event) => onChange(event.target.value || undefined)}
                  placeholder="#RRGGBB"
                  aria-label="颜色值"
                  className="mo-field h-7 min-w-0 flex-1 font-mono text-[10px]"
                />
                {allowClear ? (
                  <button
                    type="button"
                    aria-label="清除颜色"
                    data-tooltip="清除颜色"
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={() => choose(undefined)}
                    className="flex h-7 w-7 items-center justify-center rounded-md border border-[var(--mo-line)] text-[var(--mo-text-muted)] hover:bg-[var(--mo-surface-2)]"
                  >
                    <Ban size={12} />
                  </button>
                ) : null}
              </div>
            </div>,
            document.body,
          )
        : null}
    </div>
  );
}
