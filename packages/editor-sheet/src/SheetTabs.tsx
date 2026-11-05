// 功能：工作表标签支持右键操作、原位改名和颜色选择，使用固定 13px 字号及 32px 点击高度；时间：2026-11-03；作者：lq。
import { useEffect, useRef, useState } from 'react';
import type { SheetModel } from '@mini-office/document-model';
import { Button, ColorField, ContextMenu, Dialog } from '@mini-office/ui';

export function SheetTabs({ sheets, activeId, readonly, onSwitch, onRename, onColor, onAdd, onDelete }: {
  sheets: SheetModel[]; activeId: string; readonly: boolean;
  onSwitch: (id: string) => void; onRename: (id: string, name: string) => void;
  onColor: (id: string, color?: string) => void; onAdd: () => void; onDelete: (sheet: SheetModel) => void;
}) {
  const [menu, setMenu] = useState<{ x: number; y: number; sheet: SheetModel } | null>(null);
  const [rename, setRename] = useState<{ id: string; name: string } | null>(null);
  const [error, setError] = useState('');
  const [colorSheet, setColorSheet] = useState<SheetModel | null>(null);
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => { if (rename) { input.current?.focus({ preventScroll: true }); input.current?.select(); } }, [rename?.id]);
  const beginRename = (sheet: SheetModel) => { if (!readonly) { setError(''); setRename({ id: sheet.id, name: sheet.name }); } };
  // 功能：复用名称校验与公式引用更新，非法或重复名称保留输入并给出原因；时间：2026-10-24；作者：lq。
  const finishRename = () => {
    if (!rename) return;
    try { onRename(rename.id, rename.name); setRename(null); setError(''); }
    catch (reason) { setError(reason instanceof Error ? reason.message : String(reason)); requestAnimationFrame(() => input.current?.focus({ preventScroll: true })); }
  };
  return <>
    <div className="mo-scroll mo-sheet-tabs relative flex min-w-0 flex-1 items-center gap-1 overflow-x-auto" role="tablist" aria-label="工作表">
      {sheets.filter((item) => item.state === 'visible').map((item) => rename?.id === item.id ?
        <input key={item.id} ref={input} aria-label="工作表新名称" aria-invalid={!!error} title={error || 'Enter 确认，Esc 取消'} value={rename.name} maxLength={31} className="h-8 w-[150px] shrink-0 rounded border border-[var(--mo-accent)] bg-[var(--mo-surface)] px-2 text-[13px] outline-none"
          onChange={(event) => { setRename({ ...rename, name: event.target.value }); setError(''); }} onBlur={finishRename}
          onKeyDown={(event) => { event.stopPropagation(); if (event.key === 'Enter') { event.preventDefault(); finishRename(); } if (event.key === 'Escape') { event.preventDefault(); setRename(null); setError(''); } }} /> :
        <button key={item.id} type="button" role="tab" aria-selected={item.id === activeId} onClick={() => onSwitch(item.id)} onDoubleClick={() => beginRename(item)}
          onContextMenu={(event) => { event.preventDefault(); setMenu({ x: event.clientX, y: event.clientY, sheet: item }); }}
          className={`h-8 shrink-0 rounded-t-md border-b-[3px] px-4 text-[13px] transition-colors ${item.id === activeId ? 'bg-[var(--mo-surface)] font-semibold' : 'text-[var(--mo-text-muted)] hover:bg-[var(--mo-surface-2)]'}`}
          style={{ borderBottomColor: item.tabColor ?? (item.id === activeId ? 'var(--mo-accent)' : 'transparent'), background: item.tabColor && item.id === activeId ? `color-mix(in srgb, ${item.tabColor} 15%, var(--mo-surface))` : undefined }}>
          {item.name}
        </button>)}
      <button type="button" aria-label="添加工作表" disabled={readonly} onClick={onAdd} className="h-8 w-8 shrink-0 rounded text-[20px] hover:bg-[var(--mo-surface-2)] disabled:opacity-40">+</button>
    </div>
    {error ? <span role="alert" className="max-w-[240px] text-[13px] text-[var(--mo-danger)]">{error}</span> : null}
    {menu ? <ContextMenu x={menu.x} y={menu.y} onClose={() => setMenu(null)} items={[
      { id: 'add', label: '插入工作表', disabled: readonly, onSelect: onAdd },
      { id: 'rename', label: '重命名', disabled: readonly, onSelect: () => beginRename(menu.sheet) },
      { id: 'delete', label: '删除工作表', danger: true, disabled: readonly || sheets.filter((item) => item.state === 'visible').length <= 1, onSelect: () => onDelete(menu.sheet) },
      { type: 'separator' },
      { id: 'color', label: '工作表标签颜色…', disabled: readonly, onSelect: () => setColorSheet(menu.sheet) },
    ]} /> : null}
    <Dialog open={!!colorSheet} title="工作表标签颜色" width="sm" onClose={() => setColorSheet(null)} footer={<Button onClick={() => setColorSheet(null)}>完成</Button>}>
      {colorSheet ? <div className="space-y-4"><p className="text-[12px]">{colorSheet.name}</p><ColorField label="标签颜色" value={sheets.find((item) => item.id === colorSheet.id)?.tabColor} onChange={(color) => onColor(colorSheet.id, color)} /><div className="grid grid-cols-8 gap-2">{['#D92D20', '#F79009', '#FACA15', '#12B76A', '#0BA5EC', '#3478F6', '#7A5AF8', '#EE46BC', '#000000', '#475467', '#98A2B3', '#FFFFFF', '#B42318', '#027A48', '#085D9E', '#5925DC'].map((color) => <button key={color} type="button" aria-label={`标签颜色 ${color}`} className="h-7 w-7 rounded border border-[var(--mo-line)]" style={{ background: color }} onClick={() => { onColor(colorSheet.id, color); setColorSheet(null); }} />)}</div><Button onClick={() => { onColor(colorSheet.id, undefined); setColorSheet(null); }}>无颜色</Button></div> : null}
    </Dialog>
  </>;
}
