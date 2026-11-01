// 功能：回归公式点选/拖选、光标替换和 Excel 边框预设的选区语义；时间：2026-10-24；作者：lq。
import { act, createElement, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';
import { createBlankWorkbook } from '@mini-office/document-model';
import { FormulaInput } from '../packages/editor-sheet/src/FormulaInput';
import { SheetGrid } from '../packages/editor-sheet/src/SheetGrid';
import { SheetToolbar } from '../packages/editor-sheet/src/SheetToolbar';
import { completeFormulaParentheses, formulaReferenceSpan, pickFormulaReference } from '../packages/editor-sheet/src/formula/reference-picking';
import { applyBordersToRange, cellStyle } from '../packages/editor-sheet/src/sheet-utils';

// 功能：构造隔离的 React 容器并在验证后卸载；时间：2026-10-24；作者：lq。
function mount(node: React.ReactElement) {
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
  const host = document.createElement('div'); document.body.append(host); const root = createRoot(host);
  act(() => root.render(node));
  return { host, dispose: () => { act(() => root.unmount()); host.remove(); } };
}

describe('公式引用选择', () => {
  it('仅在操作数位置选择引用，已有引用可替换，字符串和完成的表达式不进入点选模式', () => {
    expect(formulaReferenceSpan('=SUM(', 5)).toEqual({ start: 5, end: 5 });
    expect(formulaReferenceSpan('=A1+', 4)).toEqual({ start: 4, end: 4 });
    expect(formulaReferenceSpan('=SUM($B$1:C4)', 8)).toEqual({ start: 5, end: 12 });
    expect(formulaReferenceSpan('=SUM(A1)', 8)).toBeNull();
    expect(formulaReferenceSpan('=LOG10(', 5)).toBeNull();
    expect(formulaReferenceSpan('=IF(A1="hello', 13)).toBeNull();
    expect(formulaReferenceSpan('普通文本', 4)).toBeNull();
    expect(formulaReferenceSpan("=SUM('A1", 8)).toBeNull();
  });

  it('回车提交可补齐嵌套函数括号，字符串内容和不完整运算符不受影响', () => {
    expect(completeFormulaParentheses('=SUM(B1:C1')).toBe('=SUM(B1:C1)');
    expect(completeFormulaParentheses('=IF(A1>0,SUM(B1:C1),0')).toBe('=IF(A1>0,SUM(B1:C1),0)');
    expect(completeFormulaParentheses('=SUM(A1+')).toBe('=SUM(A1+');
    expect(completeFormulaParentheses('="SUM(A1"')).toBe('="SUM(A1"');
    expect(completeFormulaParentheses('=SUM(\'测试(表)\'!A1')).toBe('=SUM(\'测试(表)\'!A1)');
  });

  it('拖选反复更新同一次引用，保留嵌套公式前后文和输入焦点', () => {
    const initial = '=IF(A1>0,SUM(B1:B3),0)';
    function Harness() { const [value, setValue] = useState(initial); return createElement(FormulaInput, { value, onValueChange: setValue, 'aria-label': '引用公式' }); }
    const view = mount(createElement(Harness));
    try {
      const input = view.host.querySelector('input')!;
      act(() => { input.focus(); input.setSelectionRange(14, 14); });
      act(() => { expect(pickFormulaReference(input, 'start', 'C2')).toBe(true); });
      act(() => { expect(pickFormulaReference(input, 'update', 'C2:D5')).toBe(true); });
      act(() => { pickFormulaReference(input, 'end', ''); });
      expect(input.value).toBe('=IF(A1>0,SUM(C2:D5),0)');
      expect(input.selectionStart).toBe(18);
      expect(document.activeElement).toBe(input);
    } finally { view.dispose(); }
  });

  it('公式栏拖选网格插入区域，不改变当前单元格选区，也不提交草稿', () => {
    const book = createBlankWorkbook(); const scrollRef = { current: null as HTMLDivElement | null };
    const change = vi.fn(); const commit = vi.fn();
    function Harness() {
      const [value, setValue] = useState('=SUM(');
      return createElement('div', null,
        createElement(FormulaInput, { value, onValueChange: setValue, 'aria-label': '公式栏' }),
        createElement(SheetGrid, { workbook: book, sheet: book.sheets[0], selection: { anchor: { row: 0, col: 3 }, focus: { row: 0, col: 3 } }, editing: null,
          formulaDraft: value, scrollElementRef: scrollRef, onSelectionChange: change, onStartEdit: vi.fn(), onCommitEdit: commit, onCancelEdit: vi.fn(), onColumnResize: vi.fn(), onRowResize: vi.fn(), onCellContextMenu: vi.fn(), onHeaderContextMenu: vi.fn(), onAction: vi.fn(), onFill: vi.fn() }));
    }
    const view = mount(createElement(Harness));
    try {
      const scroll = scrollRef.current!;
      vi.spyOn(scroll, 'getBoundingClientRect').mockReturnValue({ left: 46, top: 26, right: 646, bottom: 426, width: 600, height: 400, x: 46, y: 26, toJSON: () => ({}) });
      Object.defineProperty(scroll, 'offsetWidth', { value: 600 });
      scroll.scrollLeft = 0; scroll.scrollTop = 0;
      const input = view.host.querySelector('input')!;
      act(() => { input.focus(); input.setSelectionRange(5, 5); });
      act(() => view.host.querySelector('[data-cell="0:1"]')!.dispatchEvent(new MouseEvent('mousedown', { button: 0, bubbles: true, cancelable: true })));
      expect(input.value).toBe('=SUM(B1');
      act(() => window.dispatchEvent(new MouseEvent('mousemove', { clientX: 46 + 88 * 2 + 44, clientY: 26 + 24 + 12 })));
      act(() => window.dispatchEvent(new MouseEvent('mouseup')));
      expect(input.value).toBe('=SUM(B1:C2');
      expect(view.host.querySelector('[data-formula-reference="B1:C2"]')).not.toBeNull();
      expect(document.activeElement).toBe(input);
      expect(change).not.toHaveBeenCalled(); expect(commit).not.toHaveBeenCalled();
    } finally { view.dispose(); }
  });
});

describe('公式栏快捷操作', () => {
  it('fx 展开函数清单并写入所选函数，自动求和有独立入口', () => {
    const book = createBlankWorkbook();
    const sum = vi.fn(); const change = vi.fn();
    const noOp = vi.fn();
    const view = mount(createElement(SheetToolbar, {
      workbook: book, sheet: book.sheets[0], range: { r1: 9, c1: 0, r2: 17, c2: 0 },
      active: { row: 9, col: 0 }, formulaText: '', readonly: false, canUndo: false, canRedo: false,
      propertiesOpen: false, crosshair: false,
      onFormulaChange: change, onFormulaCommit: noOp, onFormulaCancel: noOp,
      onFormulaFocus: noOp, onAutoSum: sum,
      onHideRow: noOp, onShowRows: noOp, onHideColumn: noOp, onShowColumns: noOp,
      onJump: noOp, onStyle: noOp, onNumberFormat: noOp, onBorders: noOp,
      borderWeight: 'thin', onBorderWeight: noOp, borderColor: '#000000', onBorderColor: noOp,
      onMoreBorders: noOp, onMergeToggle: noOp, onFreeze: noOp, onInsertRow: noOp,
      onDeleteRow: noOp, onInsertColumn: noOp, onDeleteColumn: noOp, onSort: noOp,
      onFilter: noOp, onInsertImage: noOp, onCustomSort: noOp, onClearFilter: noOp,
      onReapplyFilter: noOp, onFind: noOp, onExport: noOp, onPrint: noOp,
      onPageSetup: noOp, onCompatibility: noOp, onToggleProperties: noOp,
      onToggleCrosshair: noOp, onUndo: noOp, onRedo: noOp, onAddSheet: noOp,
    }));
    try {
      // 功能：fx 现在是函数清单入口，选中具体函数后才写入公式栏；时间：2026-11-01；作者：lq。
      act(() => (view.host.querySelector('button[aria-label="插入函数"]') as HTMLButtonElement).click());
      const sumItem = Array.from(document.querySelectorAll<HTMLButtonElement>('[role="menuitem"]'))
        .find((item) => item.textContent?.startsWith('SUM'));
      expect(sumItem).toBeTruthy();
      act(() => sumItem!.click());
      expect(change).toHaveBeenCalledWith('=SUM(');
      act(() => (view.host.querySelector('button[aria-label="自动求和"]') as HTMLButtonElement).click());
      expect(sum).toHaveBeenCalledOnce();
    } finally { view.dispose(); }
  });
});

describe('边框选区语义', () => {
  it('下框线只画在选区底边，保留自定义线型和颜色', () => {
    const book = createBlankWorkbook(); const id = book.sheets[0].id;
    const next = applyBordersToRange(book, id, { r1: 0, c1: 0, r2: 2, c2: 2 }, 'bottom', { style: 'dashed', color: '#D92D20' });
    expect(next.sheets[0].cells['0:0']).toBeUndefined();
    expect(cellStyle(next, next.sheets[0].cells['2:1']).borders?.bottom).toEqual({ style: 'dashed', color: '#D92D20' });
  });
  it('复合边框上下分别使用细线和双线，粗外框四周使用粗线', () => {
    const book = createBlankWorkbook(); const id = book.sheets[0].id; const range = { r1: 0, c1: 0, r2: 2, c2: 2 };
    const next = applyBordersToRange(book, id, range, 'topDoubleBottom', { style: 'thin', color: '#3478F6' });
    expect(cellStyle(next, next.sheets[0].cells['0:1']).borders?.top?.style).toBe('thin');
    expect(cellStyle(next, next.sheets[0].cells['2:1']).borders?.bottom?.style).toBe('double');
    expect(next.sheets[0].cells['1:1']).toBeUndefined();
    const thick = applyBordersToRange(book, id, range, 'thickOutline', { style: 'thin' });
    expect(cellStyle(thick, thick.sheets[0].cells['0:0']).borders?.left?.style).toBe('thick');
    expect(cellStyle(thick, thick.sheets[0].cells['2:2']).borders?.right?.style).toBe('thick');
  });
  it('内部框线只修改相邻单元格的公共边，不产生外框', () => {
    const book = createBlankWorkbook(); const id = book.sheets[0].id;
    const next = applyBordersToRange(book, id, { r1: 0, c1: 0, r2: 1, c2: 1 }, 'inside', { style: 'dotted' });
    const first = cellStyle(next, next.sheets[0].cells['0:0']).borders;
    expect(first?.top).toBeUndefined(); expect(first?.left).toBeUndefined();
    expect(first?.bottom?.style).toBe('dotted'); expect(first?.right?.style).toBe('dotted');
    const last = cellStyle(next, next.sheets[0].cells['1:1']).borders;
    expect(last?.bottom).toBeUndefined(); expect(last?.right).toBeUndefined();
  });
});
