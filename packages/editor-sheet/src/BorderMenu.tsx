// 功能：提供 Excel 风格的边框分组菜单、重复应用按钮、线型和颜色设置；时间：2026-10-24；作者：lq。
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { ChevronDown } from 'lucide-react';
import type { BorderSide } from '@mini-office/document-model';
import { Select } from '@mini-office/ui';
import type { BorderPreset } from './sheet-utils';

const PRESETS: Array<[BorderPreset, string] | null> = [
  ['bottom', '下框线'], ['top', '上框线'], ['left', '左框线'], ['right', '右框线'], null,
  ['none', '无框线'], ['all', '所有框线'], ['inside', '内部框线'], ['outline', '外侧框线'], ['thickOutline', '粗外框线'], null,
  ['doubleBottom', '双底框线'], ['thickBottom', '粗底框线'], ['topBottom', '上下框线'], ['topThickBottom', '上框线和粗下框线'], ['topDoubleBottom', '上框线和双下框线'],
];

// 功能：用真实线段预览各边框预设，便于直接辨认位置和粗细；时间：2026-10-24；作者：lq。
function BorderIcon({ preset }: { preset: BorderPreset }) {
  const top = ['top', 'outline', 'thickOutline', 'all', 'topBottom', 'topThickBottom', 'topDoubleBottom'].includes(preset);
  const bottom = !['top', 'left', 'right', 'none', 'inside'].includes(preset);
  const left = ['left', 'outline', 'thickOutline', 'all'].includes(preset);
  const right = ['right', 'outline', 'thickOutline', 'all'].includes(preset);
  const heavy = preset === 'thickOutline';
  return <svg width="17" height="17" viewBox="0 0 20 20" aria-hidden="true"><rect x="2" y="2" width="16" height="16" fill="none" stroke="currentColor" opacity="0.25" strokeDasharray="1 2" />
    {top ? <path d="M2 2H18" stroke="currentColor" strokeWidth={heavy ? 2.5 : 1.2} /> : null}
    {bottom ? <path d="M2 18H18" stroke="currentColor" strokeWidth={heavy || preset.includes('Thick') || preset === 'thickBottom' ? 2.5 : 1.2} /> : null}
    {preset === 'doubleBottom' || preset === 'topDoubleBottom' ? <path d="M2 15H18" stroke="currentColor" strokeWidth="1.2" /> : null}
    {left ? <path d="M2 2V18" stroke="currentColor" strokeWidth={heavy ? 2.5 : 1.2} /> : null}
    {right ? <path d="M18 2V18" stroke="currentColor" strokeWidth={heavy ? 2.5 : 1.2} /> : null}
    {preset === 'all' || preset === 'inside' ? <path d="M10 2V18M2 10H18" stroke="currentColor" strokeWidth="1.2" /> : null}
  </svg>;
}

export function BorderMenu({ readonly, weight, color, onWeight, onColor, onApply, onMore }: {
  readonly: boolean; weight: BorderSide['style']; color: string;
  onWeight: (style: Exclude<BorderSide['style'], 'none'>) => void; onColor: (color: string) => void;
  onApply: (preset: BorderPreset) => void; onMore: () => void;
}) {
  const [open, setOpen] = useState(false);
  const [last, setLast] = useState<BorderPreset>('bottom');
  const [position, setPosition] = useState({ left: 0, top: 0, height: 600 });
  const anchor = useRef<HTMLDivElement>(null); const menu = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    if (!open) return;
    const rect = anchor.current!.getBoundingClientRect();
    setPosition({ left: Math.max(8, Math.min(rect.left, window.innerWidth - 244)), top: rect.bottom + 4, height: window.innerHeight - rect.bottom - 12 });
  }, [open]);
  useEffect(() => {
    if (!open) return;
    const outside = (event: PointerEvent) => { if (!menu.current?.contains(event.target as Node) && !anchor.current?.contains(event.target as Node)) setOpen(false); };
    const key = (event: KeyboardEvent) => { if (event.key === 'Escape') setOpen(false); };
    const close = () => setOpen(false);
    document.addEventListener('pointerdown', outside, true); document.addEventListener('keydown', key); window.addEventListener('resize', close);
    return () => { document.removeEventListener('pointerdown', outside, true); document.removeEventListener('keydown', key); window.removeEventListener('resize', close); };
  }, [open]);
  const apply = (preset: BorderPreset) => { setLast(preset); onApply(preset); setOpen(false); };
  return <div ref={anchor} className="inline-flex shrink-0 rounded hover:bg-[var(--mo-surface-2)]">
    <button type="button" aria-label="应用上次边框" title={PRESETS.find((item) => item?.[0] === last)?.[1]} disabled={readonly} onClick={() => apply(last)} className="flex h-7 w-7 items-center justify-center"><BorderIcon preset={last} /></button>
    <button type="button" aria-label="边框" aria-expanded={open} aria-haspopup="menu" disabled={readonly} onClick={() => setOpen(!open)} className="flex h-7 w-4 items-center justify-center"><ChevronDown size={11} /></button>
    {open ? createPortal(<div ref={menu} role="menu" aria-label="边框设置" className="mo-glass fixed z-[95] w-[236px] overflow-y-auto rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface)] p-1 text-[12px] text-[var(--mo-text-primary)] shadow-xl" style={{ left: position.left, top: position.top, maxHeight: position.height }}
      onMouseDown={(event) => { if ((event.target as HTMLElement).closest('[role="menuitem"]')) event.preventDefault(); }}>
      {PRESETS.map((item, i) => item ? <button key={item[0]} type="button" role="menuitem" className="flex w-full items-center gap-3 rounded px-2 py-1.5 text-left hover:bg-[var(--mo-accent-soft)]" onClick={() => apply(item[0])}><BorderIcon preset={item[0]} /><span>{item[1]}</span>{last === item[0] ? <span className="ml-auto text-[var(--mo-accent)]">✓</span> : null}</button> : <div key={i} role="separator" className="my-1 border-t border-[var(--mo-line)]" />)}
      <div className="mt-1 space-y-2 border-t border-[var(--mo-line)] px-2 py-2">
        <label className="flex items-center justify-between gap-2">线条颜色<input type="color" aria-label="边框线条颜色" value={color} onChange={(event) => onColor(event.target.value)} className="h-6 w-12 cursor-pointer bg-transparent" /></label>
        <label className="flex items-center justify-between gap-2">线型<Select aria-label="边框线型" className="h-7 w-[132px] text-[11px]" value={weight} onChange={(event) => onWeight(event.target.value as Exclude<BorderSide['style'], 'none'>)}>{[['thin', '细实线'], ['medium', '中实线'], ['thick', '粗实线'], ['dashed', '虚线'], ['dotted', '点线'], ['double', '双线']].map(([value, label]) => <option key={value} value={value}>{label}</option>)}</Select></label>
      </div>
      <button type="button" role="menuitem" className="flex w-full items-center gap-3 rounded px-2 py-2 text-left hover:bg-[var(--mo-accent-soft)]" onClick={() => { setOpen(false); onMore(); }}><BorderIcon preset="all" />更多边框…</button>
    </div>, document.body) : null}
  </div>;
}
