import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type MouseEvent as ReactMouseEvent,
  type MutableRefObject,
} from 'react';
import type {
  CellBorders,
  CellRange,
  CellStyle,
  SheetModel,
  SheetImage,
  WorkbookModel,
} from '@mini-office/document-model';
import { columnLabel } from '@mini-office/shared';
import { fontStackFor, paperFill, paperInk, paperLine, usePaperTheme } from '@mini-office/design-system';
import { cellKey } from '@mini-office/document-model';
// 功能：在表头显示筛选入口与激活状态；时间：2026-10-24；作者：lq。
import { ChevronDown, Filter } from 'lucide-react';
// 功能：为单元格内的公式编辑启用函数智能提示；时间：2026-10-24；作者：lq。
// 功能：渲染可编辑浮动图片；时间：2026-10-24；作者：lq。
import { pickFormulaReference } from './formula/reference-picking';
import { SheetImageLayer } from './SheetImageLayer';
import { FormulaInput } from './FormulaInput';
import {
  cellDisplay,
  cellStyle,
  selectionRange,
  rangeRef,
  type Selection,
} from './sheet-utils';

export type GridAction =
  | 'copy'
  | 'cut'
  | 'paste'
  | 'undo'
  | 'redo'
  | 'selectAll'
  | 'clear'
  | 'find'
  | 'delete';

export interface EditingCell {
  row: number;
  col: number;
  initialText: string;
}

export interface SheetGridProps {
  gridId?: string;
  scrollElementRef?: MutableRefObject<HTMLDivElement | null>;
  workbook: WorkbookModel;
  sheet: SheetModel;
  selection: Selection;
  editing: EditingCell | null;
  onSelectionChange: (selection: Selection, options?: { extend?: boolean }) => void;
  onStartEdit: (row: number, col: number, initialText?: string) => void;
  onCommitEdit: (value: string, move: 'down' | 'right' | 'up' | 'left' | 'none') => void;
  onCancelEdit: () => void;
  // 功能：单元格编辑与顶部公式栏同步同一份草稿；时间：2026-10-24；作者：lq。
  formulaDraft?: string;
  onDraftChange?: (value: string) => void;
  onColumnResize: (col: number, width: number) => void;
  onRowResize: (row: number, height: number) => void;
  onCellContextMenu: (event: ReactMouseEvent, row: number, col: number) => void;
  onHeaderContextMenu: (event: ReactMouseEvent, kind: 'row' | 'col', index: number) => void;
  onAction: (action: GridAction) => void;
  // 功能：在筛选区域表头打开列筛选菜单；时间：2026-10-24；作者：lq。
  onFilterColumn?: (col: number, anchor: DOMRect) => void;
  // 功能：图片编辑回调及缩放坐标换算参数；时间：2026-10-24；作者：lq。
  readonly?: boolean;
  /** 功能：阅读模式——十字高亮当前单元格所在行列；时间：2026-10-30；作者：lq。 */
  crosshair?: boolean;
  zoom?: number;
  onImageChange?: (id: string, patch: Partial<SheetImage>) => void;
  onImageDelete?: (id: string) => void;
  onFill: (source: CellRange, target: CellRange) => void;
}

// 功能：单个引用显示 B1，矩形引用显示 B1:C3，与 Excel 输入习惯一致；时间：2026-10-24；作者：lq。
function referenceAddress(range: CellRange): string {
  return range.r1 === range.r2 && range.c1 === range.c2 ? `${columnLabel(range.c1)}${range.r1 + 1}` : rangeRef(range);
}

const ROW_HEADER_WIDTH = 46;
const COL_HEADER_HEIGHT = 26;
const OVERSCAN = 4;

function findIndex(offsets: number[], value: number): number {
  let low = 0;
  let high = offsets.length - 2;
  let result = 0;
  while (low <= high) {
    const mid = Math.floor((low + high) / 2);
    if (offsets[mid] <= value) {
      result = mid;
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }
  return Math.max(0, result);
}

function borderCss(side: CellBorders['top']): string | undefined {
  if (!side) return undefined;
  // 功能：让“无框线”覆盖默认网格线，避免边框清除后仍显示浅色网格；时间：2026-10-24；作者：lq。
  if (side.style === 'none') return 'none';
  const width = side.style === 'thick' || side.style === 'double' ? 3 : side.style === 'medium' ? 2 : 1;
  // 功能：将 Excel 的“细/中/粗”边框语义转换为浏览器可识别的实线；时间：2026-10-24；作者：lq。
  const lineStyle = side.style === 'double' ? 'double' : side.style === 'dashed' ? 'dashed' : side.style === 'dotted' ? 'dotted' : 'solid';
  // 功能：黑色框线走主题变量，深色纸张上自动提亮、打印与导出时仍是纯黑；时间：2026-10-31；作者：lq。
  return `${width}px ${lineStyle} ${paperLine(side.color, '#000000')}`;
}

/** 功能：边框粗细排序，用于相邻单元格共用一条线时保留更粗的一侧；时间：2026-11-01；作者：lq。 */
function borderWeightOf(side: CellBorders['top']): number {
  if (!side || side.style === 'none') return 0;
  if (side.style === 'thick' || side.style === 'double') return 3;
  if (side.style === 'medium') return 2;
  return 1;
}

/**
 * 功能：解析公式中光标处的单元格引用，供编辑时高亮求和范围等区域；
 * 时间：2026-11-01；作者：lq。
 */
function referenceAtCaret(text: string, caret: number): CellRange | null {
  if (!text.startsWith('=')) return null;
  const pattern = /\$?([A-Za-z]{1,3})\$?([1-9]\d*)(?::\$?([A-Za-z]{1,3})\$?([1-9]\d*))?/g;
  let match: RegExpExecArray | null;
  let best: { range: CellRange; distance: number } | null = null;
  while ((match = pattern.exec(text))) {
    const start = match.index;
    const end = start + match[0].length;
    // 功能：避免把函数名或字符串内容误判成引用；时间：2026-11-01；作者：lq。
    if (/[A-Za-z0-9_.]/.test(text[start - 1] ?? '')) continue;
    const distance = caret < start ? start - caret : caret > end ? caret - end : 0;
    const toIndex = (letters: string, digits: string) => ({
      row: Number(digits) - 1,
      col: letters.toUpperCase().split('').reduce((sum, char) => sum * 26 + (char.charCodeAt(0) - 64), 0) - 1,
    });
    const first = toIndex(match[1], match[2]);
    const second = match[3] && match[4] ? toIndex(match[3], match[4]) : first;
    const range: CellRange = {
      r1: Math.min(first.row, second.row),
      r2: Math.max(first.row, second.row),
      c1: Math.min(first.col, second.col),
      c2: Math.max(first.col, second.col),
    };
    if (!best || distance < best.distance) best = { range, distance };
  }
  return best?.range ?? null;
}

function buildOffsets(count: number, sizeAt: (index: number) => number): number[] {
  const offsets = new Array<number>(count + 1);
  let total = 0;
  offsets[0] = 0;
  for (let index = 0; index < count; index += 1) {
    total += sizeAt(index);
    offsets[index + 1] = total;
  }
  return offsets;
}

export function SheetGrid(props: SheetGridProps) {
  const {
    gridId,
    scrollElementRef,
    workbook,
    sheet,
    selection,
    editing,
    onSelectionChange,
    onStartEdit,
    onCommitEdit,
    onCancelEdit,
    onColumnResize,
    onRowResize,
    onCellContextMenu,
    onHeaderContextMenu,
    onAction,
    onFill,
    crosshair,
  } = props;

  const containerRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement | null>(null);
  // 功能：单元格底色/文字色需要按当前纸张主题换算，主题切换时重新渲染；时间：2026-10-31；作者：lq。
  const darkPaper = usePaperTheme() === 'dark';
  const [viewport, setViewport] = useState({ width: 800, height: 500 });
  const [scroll, setScroll] = useState({ left: 0, top: 0 });
  const [resizePreview, setResizePreview] = useState<{ kind: 'col' | 'row'; index: number; size: number } | null>(null);
  const [draft, setDraft] = useState('');
  const [dragging, setDragging] = useState(false);
  // 功能：引用选区独立于正在编辑的单元格，点选不移动编辑位置；时间：2026-10-24；作者：lq。
  const referenceDrag = useRef<{ input: HTMLInputElement; row: number; col: number } | null>(null);
  const referenceInput = useRef<HTMLInputElement | null>(null);
  const [referenceDragging, setReferenceDragging] = useState(false);
  const [referenceRange, setReferenceRange] = useState<CellRange | null>(null);
  const [fillPreview, setFillPreview] = useState<CellRange | null>(null);
  // 功能：拖拽结束需要读取最新预览范围，避免在 setState 回调里做副作用；时间：2026-11-01；作者：lq。
  const fillPreviewRef = useRef<CellRange | null>(null);
  const [fillSession, setFillSession] = useState(0);
  const fillRef = useRef<{ source: CellRange; startX: number; startY: number } | null>(null);
  const scrollRaf = useRef<number | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  // 功能：记录已定位光标的编辑会话，防止后续输入时反复重置光标；时间：2026-10-24；作者：lq。
  const focusedEditingRef = useRef<EditingCell | null>(null);
  const resizeRef = useRef<{ kind: 'col' | 'row'; index: number; start: number; initial: number } | null>(null);

  const range = selectionRange(selection);
  const active = selection.kind ? selection.anchor : selection.focus;

  const rowsCount = Math.max(sheet.dimension.rows, range.r2 + (selection.kind ? 1 : 2), 100);
  const colsCount = Math.max(sheet.dimension.cols, range.c2 + (selection.kind ? 1 : 2), 26);

  const colOffsets = useMemo(
    () =>
      buildOffsets(colsCount, (col) => {
        if (resizePreview?.kind === 'col' && resizePreview.index === col) return resizePreview.size;
        return sheet.hiddenCols[col] ? 0 : sheet.colWidths[col] ?? sheet.defaultColWidth;
      }),
    [colsCount, resizePreview, sheet.colWidths, sheet.defaultColWidth, sheet.hiddenCols],
  );
  const rowOffsets = useMemo(
    () =>
      buildOffsets(rowsCount, (row) => {
        if (resizePreview?.kind === 'row' && resizePreview.index === row) return resizePreview.size;
        return sheet.hiddenRows[row] ? 0 : sheet.rowHeights[row] ?? sheet.defaultRowHeight;
      }),
    [rowsCount, resizePreview, sheet.rowHeights, sheet.defaultRowHeight, sheet.hiddenRows],
  );

  const totalWidth = colOffsets[colOffsets.length - 1] ?? 0;
  const totalHeight = rowOffsets[rowOffsets.length - 1] ?? 0;

  const mergeLookup = useMemo(() => {
    const origin = new Map<string, CellRange>();
    const covered = new Set<string>();
    for (const merge of sheet.merges) {
      origin.set(cellKey(merge.r1, merge.c1), merge);
      for (let row = merge.r1; row <= merge.r2; row += 1) {
        for (let col = merge.c1; col <= merge.c2; col += 1) {
          if (row !== merge.r1 || col !== merge.c1) covered.add(cellKey(row, col));
        }
      }
    }
    return { origin, covered };
  }, [sheet.merges]);

  // 功能：预先收集显式边框，渲染时判断相邻单元格是否重复画同一条线；时间：2026-11-01；作者：lq。
  const borderLookup = useMemo(() => {
    const map = new Map<string, CellBorders>();
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const borders = cellStyle(workbook, cell).borders;
      if (borders) map.set(key, borders);
    }
    return map;
  }, [sheet.cells, workbook]);

  useLayoutEffect(() => {
    // Filtering can shrink the scrollable area while the browser clamps scrollTop.
    const element = scrollRef.current;
    if (!element) return;
    setScroll((previous) => previous.left === element.scrollLeft && previous.top === element.scrollTop
      ? previous : { left: element.scrollLeft, top: element.scrollTop });
  }, [totalHeight, totalWidth]);

  const startRow = findIndex(rowOffsets, scroll.top);
  const endRow = Math.min(rowsCount - 1, findIndex(rowOffsets, scroll.top + viewport.height) + OVERSCAN);
  const startCol = findIndex(colOffsets, scroll.left);
  const endCol = Math.min(colsCount - 1, findIndex(colOffsets, scroll.left + viewport.width) + OVERSCAN);
  const visibleRows = useMemo(() => {
    const result: number[] = [];
    for (let row = Math.max(0, startRow - 1); row <= endRow; row += 1) result.push(row);
    return result;
  }, [endRow, startRow]);
  const visibleCols = useMemo(() => {
    const result: number[] = [];
    for (let col = Math.max(0, startCol - 1); col <= endCol; col += 1) result.push(col);
    return result;
  }, [endCol, startCol]);

  useLayoutEffect(() => {
    const element = containerRef.current;
    if (!element) return;
    const update = () => {
      const scrollElement = scrollRef.current;
      setViewport({
        width: scrollElement?.clientWidth ?? element.clientWidth - ROW_HEADER_WIDTH,
        height: scrollElement?.clientHeight ?? element.clientHeight - COL_HEADER_HEIGHT,
      });
    };
    update();
    if (typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(update);
    observer.observe(element);
    return () => observer.disconnect();
  }, []);

  // 功能：记录编辑光标位置，用于高亮公式中当前引用的区域（例如求和范围）；时间：2026-11-01；作者：lq。
  const [editCaret, setEditCaret] = useState(0);

  // 功能：在绘制前初始化编辑文本，避免快速输入时光标落在等号前面；时间：2026-10-24；作者：lq。
  useLayoutEffect(() => {
    if (!editing) {
      focusedEditingRef.current = null;
      return;
    }
    setDraft(editing.initialText);
    setEditCaret(editing.initialText.length);
  }, [editing]);

  // 功能：从公式栏修改草稿时同步网格输入框，避免两个入口互相覆盖；时间：2026-10-24；作者：lq。
  useLayoutEffect(() => {
    if (editing && props.formulaDraft !== undefined) setDraft(props.formulaDraft);
  }, [editing, props.formulaDraft]);

  useEffect(() => {
    const clear = (event: FocusEvent) => { if (event.target === referenceInput.current) setReferenceRange(null); };
    document.addEventListener('focusout', clear);
    return () => document.removeEventListener('focusout', clear);
  }, []);

  // 功能：等待受控草稿实际写入 DOM 后同步聚焦，替代有竞态的延时聚焦；时间：2026-10-24；作者：lq。
  useLayoutEffect(() => {
    if (!editing || draft !== editing.initialText || focusedEditingRef.current === editing) return;
    focusedEditingRef.current = editing;
    inputRef.current?.focus();
    inputRef.current?.setSelectionRange(draft.length, draft.length);
  }, [draft, editing]);

  useEffect(() => {
    if (!dragging) return;
    const stop = () => setDragging(false);
    window.addEventListener('mouseup', stop);
    return () => window.removeEventListener('mouseup', stop);
  }, [dragging]);

  useEffect(() => {
    if (!resizeRef.current) return;
    const onMove = (event: MouseEvent) => {
      const resize = resizeRef.current;
      if (!resize) return;
      const delta = resize.kind === 'col' ? event.clientX - resize.start : event.clientY - resize.start;
      setResizePreview({
        kind: resize.kind,
        index: resize.index,
        size: Math.max(resize.kind === 'col' ? 24 : 14, resize.initial + delta),
      });
    };
    const onUp = () => {
      const resize = resizeRef.current;
      if (resize && resizePreview) {
        if (resize.kind === 'col') onColumnResize(resize.index, resizePreview.size);
        else onRowResize(resize.index, resizePreview.size);
      }
      resizeRef.current = null;
      setResizePreview(null);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
  }, [onColumnResize, onRowResize, resizePreview]);

  useEffect(() => {
    const scrollElement = scrollRef.current;
    // 功能：表头选择仅改变选区，保持当前水平和垂直滚动位置；时间：2026-10-24；作者：lq。
    if (!scrollElement || selection.kind || sheet.hiddenRows[active.row] || sheet.hiddenCols[active.col]) return;
    const left = colOffsets[active.col] ?? 0;
    const right = colOffsets[active.col + 1] ?? left;
    const top = rowOffsets[active.row] ?? 0;
    const bottom = rowOffsets[active.row + 1] ?? top;
    if (left < scrollElement.scrollLeft) scrollElement.scrollLeft = left;
    else if (right > scrollElement.scrollLeft + scrollElement.clientWidth) {
      scrollElement.scrollLeft = right - scrollElement.clientWidth;
    }
    if (top < scrollElement.scrollTop) scrollElement.scrollTop = top;
    else if (bottom > scrollElement.scrollTop + scrollElement.clientHeight) {
      scrollElement.scrollTop = bottom - scrollElement.clientHeight;
    }
  }, [active.col, active.row, colOffsets, rowOffsets, selection.kind, sheet.hiddenRows, sheet.hiddenCols]);

  const handleScroll = useCallback(() => {
    if (scrollRaf.current !== null) return;
    scrollRaf.current = requestAnimationFrame(() => {
      scrollRaf.current = null;
      const element = scrollRef.current;
      if (!element) return;
      setScroll({ left: element.scrollLeft, top: element.scrollTop });
    });
  }, []);

  // 功能：在公式光标处插入单元格地址并保持焦点，拖动时扩展成区域引用；时间：2026-10-24；作者：lq。
  const beginReference = (row: number, col: number): boolean => {
    const input = document.activeElement;
    if (!(input instanceof HTMLInputElement) || !input.hasAttribute('data-sheet-formula-input')) return false;
    const next = { r1: row, c1: col, r2: row, c2: col };
    if (!pickFormulaReference(input, 'start', referenceAddress(next))) return false;
    referenceInput.current = input; referenceDrag.current = { input, row, col };
    setReferenceRange(next); setReferenceDragging(true);
    return true;
  };

  const beginSelection = useCallback(
    (row: number, col: number, extend: boolean) => {
      const next: Selection = extend
        ? { anchor: selection.anchor, focus: { row, col } }
        : { anchor: { row, col }, focus: { row, col } };
      onSelectionChange(next, { extend });
      setDragging(true);
    },
    [onSelectionChange, selection.anchor],
  );

  const nextVisibleRow = useCallback(
    (row: number, direction: 1 | -1, jump = false): number => {
      let next = row;
      let remaining = jump ? 50 : 1;
      while (remaining > 0) {
        next += direction;
        if (next < 0) return 0;
        if (next >= rowsCount) return rowsCount - 1;
        if (!sheet.hiddenRows[next] || jump) remaining -= 1;
      }
      return next;
    },
    [rowsCount, sheet.hiddenRows],
  );

  const nextVisibleCol = useCallback(
    (col: number, direction: 1 | -1, jump = false): number => {
      let next = col;
      let remaining = jump ? 20 : 1;
      while (remaining > 0) {
        next += direction;
        if (next < 0) return 0;
        if (next >= colsCount) return colsCount - 1;
        if (!sheet.hiddenCols[next] || jump) remaining -= 1;
      }
      return next;
    },
    [colsCount, sheet.hiddenCols],
  );

  const moveSelection = useCallback(
    (row: number, col: number, extend: boolean) => {
      onSelectionChange(
        extend ? { anchor: selection.anchor, focus: { row, col } } : { anchor: { row, col }, focus: { row, col } },
        { extend },
      );
    },
    [onSelectionChange, selection.anchor],
  );

  const handleKeyDown = useCallback(
    (event: React.KeyboardEvent<HTMLDivElement>) => {
      if (editing) return;
      const mod = event.metaKey || event.ctrlKey;
      const extend = event.shiftKey;
      const key = event.key;
      if (mod && key.toLowerCase() === 'c') {
        event.preventDefault();
        onAction('copy');
        return;
      }
      if (mod && key.toLowerCase() === 'x') {
        event.preventDefault();
        onAction('cut');
        return;
      }
      if (mod && key.toLowerCase() === 'v') {
        event.preventDefault();
        onAction('paste');
        return;
      }
      if (mod && key.toLowerCase() === 'z') {
        event.preventDefault();
        onAction(event.shiftKey ? 'redo' : 'undo');
        return;
      }
      if (mod && key.toLowerCase() === 'a') {
        event.preventDefault();
        onAction('selectAll');
        return;
      }
      if (mod && key.toLowerCase() === 'f') {
        event.preventDefault();
        onAction('find');
        return;
      }
      if (key === 'Delete' || key === 'Backspace') {
        event.preventDefault();
        onAction('clear');
        return;
      }
      if (key === 'F2' || key === 'Enter') {
        event.preventDefault();
        if (key === 'Enter') moveSelection(nextVisibleRow(active.row, 1), active.col, false);
        else onStartEdit(active.row, active.col);
        return;
      }
      if (key === 'Tab') {
        event.preventDefault();
        moveSelection(active.row, nextVisibleCol(active.col, event.shiftKey ? -1 : 1), false);
        return;
      }
      if (key === 'ArrowUp' || key === 'ArrowDown' || key === 'ArrowLeft' || key === 'ArrowRight') {
        event.preventDefault();
        const jump = mod;
        if (key === 'ArrowUp') moveSelection(nextVisibleRow(active.row, -1, jump), active.col, extend);
        else if (key === 'ArrowDown') moveSelection(nextVisibleRow(active.row, 1, jump), active.col, extend);
        else if (key === 'ArrowLeft') moveSelection(active.row, nextVisibleCol(active.col, -1, jump), extend);
        else moveSelection(active.row, nextVisibleCol(active.col, 1, jump), extend);
        return;
      }
      if (key === 'Home') {
        event.preventDefault();
        moveSelection(mod ? 0 : active.row, 0, extend);
        return;
      }
      if (key === 'End') {
        event.preventDefault();
        moveSelection(active.row, Math.max(0, colsCount - 1), extend);
        return;
      }
      if (key === 'PageDown') {
        event.preventDefault();
        moveSelection(Math.min(rowsCount - 1, active.row + 20), active.col, extend);
        return;
      }
      if (key === 'PageUp') {
        event.preventDefault();
        moveSelection(Math.max(0, active.row - 20), active.col, extend);
        return;
      }
      if (!mod && !event.altKey && key.length === 1) {
        event.preventDefault();
        onStartEdit(active.row, active.col, key);
      }
    },
    [active.col, active.row, colsCount, editing, moveSelection, nextVisibleCol, nextVisibleRow, onAction, onStartEdit, rowsCount],
  );

  const commitEditing = useCallback(
    (move: 'down' | 'right' | 'up' | 'left' | 'none') => {
      onCommitEdit(draft, move);
    },
    [draft, onCommitEdit],
  );

  const autoFit = (kind: 'col' | 'row', index: number) => {
    if (props.readonly) return;
    resizeRef.current = null;
    setResizePreview(null);
    // Measure with the same font and text as the grid, including imported styles.
    const measure = document.createElement('span');
    measure.style.cssText = 'position:absolute;visibility:hidden;white-space:pre;pointer-events:none';
    document.body.append(measure);
    try {
      if (kind === 'col') {
        let width = 24;
        for (const key of Object.keys(sheet.cells)) {
          const [row, col] = key.split(':').map(Number);
          if (col !== index || sheet.hiddenRows[row]) continue;
          const style = cellStyle(workbook, sheet.cells[key]);
          measure.style.font = `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${style.fontSizePt ?? workbook.defaultFontSizePt}pt ${fontStackFor(style.fontFamily ?? workbook.defaultFont)}`;
          const text = cellDisplay(workbook, sheet, row, col);
          measure.textContent = text;
          width = Math.max(width, Math.ceil(measure.getBoundingClientRect().width || text.length * (style.fontSizePt ?? workbook.defaultFontSizePt)) + 12 + (style.indent ?? 0) * 8 + (sheet.autoFilter?.r1 === row ? 18 : 0));
        }
        onColumnResize(index, Math.min(800, width));
      } else {
        let height = sheet.defaultRowHeight;
        for (const key of Object.keys(sheet.cells)) {
          const [row, col] = key.split(':').map(Number);
          if (row !== index || sheet.hiddenCols[col]) continue;
          const style = cellStyle(workbook, sheet.cells[key]);
          const lines = cellDisplay(workbook, sheet, row, col).split('\n').length;
          height = Math.max(height, Math.ceil(lines * (style.fontSizePt ?? workbook.defaultFontSizePt) * 1.6 + 6));
        }
        onRowResize(index, Math.min(400, height));
      }
    } finally { measure.remove(); }
  };

  const renderCell = (
    row: number,
    col: number,
    position: { left: number; top: number; width: number; height: number },
    pane = false,
  ) => {
    const key = cellKey(row, col);
    const origin = mergeLookup.origin.get(key);
    const covered = mergeLookup.covered.has(key);
    if (covered && !pane) return null;
    const width = origin ? (colOffsets[origin.c2 + 1] ?? position.left + position.width) - (colOffsets[origin.c1] ?? position.left) : position.width;
    const height = origin ? (rowOffsets[origin.r2 + 1] ?? position.top + position.height) - (rowOffsets[origin.r1] ?? position.top) : position.height;
    const cell = sheet.cells[key];
    const style = cellStyle(workbook, cell);
    const value = workbook.computed?.[sheet.id]?.[key]?.value ?? cell?.value ?? null;
    const text = workbook.computed?.[sheet.id]?.[key]?.display ?? cellDisplay(workbook, sheet, row, col);
    const numeric = typeof value === 'number';
    const horizontal =
      style.horizontalAlign && style.horizontalAlign !== 'general'
        ? style.horizontalAlign
        : numeric
          ? 'right'
          : typeof value === 'boolean'
            ? 'center'
            : 'left';
    const vertical = style.verticalAlign ?? 'bottom';
    const alignSelf = vertical === 'middle' ? 'center' : vertical === 'bottom' ? 'flex-end' : 'flex-start';
    const isActive = row === active.row && col === active.col;
    const inRange =
      row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2;
    // 功能：相邻单元格共用一条边框线时只画一次，避免“所有框线”后内部线比外框粗一倍；时间：2026-11-01；作者：lq。
    const explicit = borderLookup.get(key);
    const neighborBorders = (rowOffset: number, colOffset: number) =>
      borderLookup.get(cellKey(row + rowOffset, col + colOffset));
    const skipRight = !!explicit?.right && borderWeightOf(neighborBorders(0, 1)?.left) >= borderWeightOf(explicit.right);
    const skipBottom = !!explicit?.bottom && borderWeightOf(neighborBorders(1, 0)?.top) >= borderWeightOf(explicit.bottom);
    const rightCss = skipRight ? undefined : borderCss(explicit?.right) ?? (sheet.showGridLines ? '1px solid var(--mo-grid-line)' : undefined);
    const bottomCss = skipBottom ? undefined : borderCss(explicit?.bottom) ?? (sheet.showGridLines ? '1px solid var(--mo-grid-line)' : undefined);
    const cellCss: CSSProperties = {
      position: 'absolute',
      left: position.left,
      top: position.top,
      width: Math.max(1, width),
      height: Math.max(1, height),
      display: 'flex',
      alignItems: alignSelf,
      justifyContent:
        horizontal === 'center' ? 'center' : horizontal === 'right' ? 'flex-end' : 'flex-start',
      paddingLeft: 4 + (style.indent ?? 0) * 8,
      paddingRight: sheet.autoFilter && row === sheet.autoFilter.r1 ? 21 : 4,
      boxSizing: 'border-box',
      overflow: 'hidden',
      // 功能：底色与文字色按纸张主题换算——深色纸下白底压暗、深字提亮；打印与导出由 --mo-cell-*-plain 还原；时间：2026-10-31；作者：lq。
      background: 'var(--mo-cell-fill)',
      color: 'var(--mo-cell-ink)',
      fontSize: `${style.fontSizePt ?? workbook.defaultFontSizePt}pt`,
      fontFamily: fontStackFor(style.fontFamily ?? workbook.defaultFont),
      fontWeight: style.bold ? 700 : 400,
      fontStyle: style.italic ? 'italic' : 'normal',
      textDecoration: [style.underline ? 'underline' : '', style.strike ? 'line-through' : '']
        .filter(Boolean)
        .join(' ') || undefined,
      whiteSpace: style.wrapText ? 'pre-wrap' : 'nowrap',
      wordBreak: style.wrapText ? 'break-word' : 'normal',
      borderRight: rightCss,
      borderBottom: bottomCss,
      borderTop: borderCss(explicit?.top),
      borderLeft: borderCss(explicit?.left),
      outline: isActive && !editing ? '2px solid var(--mo-accent)' : undefined,
      outlineOffset: isActive && !editing ? '-2px' : undefined,
      zIndex: isActive ? 7 : origin ? 5 : explicit ? 4 : 1,
      cursor: 'cell',
    };
    const displayText = style.shrinkToFit ? text : text;
    // 功能：同一单元格同时挂「屏幕色」与「原色」两套变量，导出/打印时由 CSS 切回原色；时间：2026-10-31；作者：lq。
    const themeVars = {
      '--mo-cell-fill': paperFill(style.fill, darkPaper),
      '--mo-cell-fill-plain': style.fill ?? 'var(--mo-paper)',
      '--mo-cell-ink': paperInk(style.color, darkPaper),
      '--mo-cell-ink-plain': style.color ?? 'var(--mo-paper-ink)',
    } as CSSProperties;
    return (
      <div
        key={key}
        data-cell={key}
        data-row={row}
        data-col={col}
        style={{ ...cellCss, ...themeVars }}
        title={text.length > 30 ? text : undefined}
        onMouseDown={(event) => {
          if (event.button !== 0) return;
          event.preventDefault();
          if (beginReference(row, col)) return;
          containerRef.current?.focus({ preventScroll: true });
          beginSelection(row, col, event.shiftKey);
        }}
        onDoubleClick={() => { if (!referenceDrag.current && document.activeElement !== referenceInput.current) onStartEdit(row, col); }}
        onContextMenu={(event) => {
          event.preventDefault();
          if (!inRange) onSelectionChange({ anchor: { row, col }, focus: { row, col } });
          onCellContextMenu(event, row, col);
        }}
      >
        {/* 功能：在实际字段表头显示下拉按钮，筛选状态以漏斗标识；时间：2026-10-24；作者：lq。 */}
        {sheet.autoFilter && row === sheet.autoFilter.r1 && col >= sheet.autoFilter.c1 && col <= sheet.autoFilter.c2 ? (
          <button type="button" aria-label={`筛选 ${columnLabel(col)} 列`} className="absolute right-0.5 top-0.5 z-10 flex h-4 w-4 items-center justify-center rounded border border-[var(--mo-line)] bg-[var(--mo-surface)] text-[var(--mo-text-primary)]"
            onMouseDown={(event) => { event.preventDefault(); event.stopPropagation(); }}
            onClick={(event) => { event.stopPropagation(); props.onFilterColumn?.(col, event.currentTarget.getBoundingClientRect()); }}>
            {sheet.filterCriteria?.[col] ? <Filter size={11} /> : <ChevronDown size={11} />}
          </button>
        ) : null}
        {editing?.row === row && editing.col === col ? null : <SheetCellText text={displayText} shrink={!!style.shrinkToFit && !style.wrapText} width={width} font={`${style.fontFamily}-${style.fontSizePt}-${style.bold}`} />}
      </div>
    );
  };

  const headerBase: CSSProperties = {
    position: 'absolute',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    background: 'var(--mo-surface-2)',
    color: 'var(--mo-text-muted)',
    fontSize: 10,
    borderRight: '1px solid var(--mo-line)',
    borderBottom: '1px solid var(--mo-line)',
    boxSizing: 'border-box',
    userSelect: 'none',
  };

  const frozenRows = Math.min(sheet.freeze?.rows ?? 0, rowsCount);
  const frozenCols = Math.min(sheet.freeze?.cols ?? 0, colsCount);
  const frozenRowsHeight = rowOffsets[frozenRows] ?? 0;
  const frozenColsWidth = colOffsets[frozenCols] ?? 0;

  /**
   * 功能：把屏幕坐标换算成行列号，兼容 CSS 缩放、滚动位置与冻结窗格；
   * 时间：2026-11-01；作者：lq。
   */
  const locateCell = useCallback(
    (clientX: number, clientY: number) => {
      const element = scrollRef.current;
      if (!element) return null;
      const rect = element.getBoundingClientRect();
      // 功能：CSS zoom 下 rect 与 offsetWidth 不一致，先换算回网格逻辑像素；时间：2026-10-24；作者：lq。
      const scale = rect.width / (element.offsetWidth || rect.width) || 1;
      const x = (clientX - rect.left) / scale;
      const y = (clientY - rect.top) / scale;
      const col = Math.min(colsCount - 1, findIndex(colOffsets, Math.max(0, x + (x < frozenColsWidth ? 0 : element.scrollLeft))));
      const row = Math.min(rowsCount - 1, findIndex(rowOffsets, Math.max(0, y + (y < frozenRowsHeight ? 0 : element.scrollTop))));
      return { row, col };
    },
    [colOffsets, colsCount, frozenColsWidth, frozenRowsHeight, rowOffsets, rowsCount],
  );

  /**
   * 功能：拖选改用全局指针坐标，替代逐单元格 mouseenter——快速拖动、跨冻结窗格与缩放时都能连续选中；
   * 时间：2026-11-01；作者：lq。
   */
  useEffect(() => {
    if (!dragging) return;
    const anchor = selection.anchor;
    let last = { row: anchor.row, col: anchor.col };
    let pointer: { x: number; y: number } | null = null;
    let frame = 0;
    const apply = () => {
      if (!pointer) return;
      const cell = locateCell(pointer.x, pointer.y);
      if (!cell || (cell.row === last.row && cell.col === last.col)) return;
      last = cell;
      onSelectionChange({ anchor, focus: cell }, { extend: true });
    };
    const move = (event: MouseEvent) => {
      pointer = { x: event.clientX, y: event.clientY };
      apply();
    };
    // 功能：拖到边缘时自动滚动，长距离选区不必反复滚动再拖；时间：2026-11-01；作者：lq。
    const tick = () => {
      const element = scrollRef.current;
      if (pointer && element) {
        const rect = element.getBoundingClientRect();
        const dx = pointer.x < rect.left ? -18 : pointer.x > rect.right ? 18 : 0;
        const dy = pointer.y < rect.top ? -18 : pointer.y > rect.bottom ? 18 : 0;
        if (dx || dy) {
          element.scrollLeft += dx;
          element.scrollTop += dy;
          apply();
        }
      }
      frame = requestAnimationFrame(tick);
    };
    frame = requestAnimationFrame(tick);
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', () => setDragging(false));
    return () => {
      cancelAnimationFrame(frame);
      window.removeEventListener('mousemove', move);
    };
  }, [dragging, locateCell, onSelectionChange, selection.anchor]);

  /**
   * 功能：填充柄拖动使用同一套坐标换算，支持缩放、冻结窗格与拖到边缘自动滚动；
   * 时间：2026-11-01；作者：lq。
   */
  useEffect(() => {
    if (!fillRef.current) return;
    let frame = 0;
    let pointer: { x: number; y: number } | null = null;
    const update = () => {
      const fill = fillRef.current;
      if (!fill || !pointer) return;
      const cell = locateCell(pointer.x, pointer.y);
      if (!cell) return;
      const source = fill.source;
      const row = cell.row;
      const col = cell.col;
      const next: CellRange =
        Math.abs(row - source.r2) >= Math.abs(col - source.c2)
          ? row >= source.r2
            ? { r1: source.r1, c1: source.c1, r2: row, c2: source.c2 }
            : { r1: row, c1: source.c1, r2: source.r2, c2: source.c2 }
          : col >= source.c2
            ? { r1: source.r1, c1: source.c1, r2: source.r2, c2: col }
            : { r1: source.r1, c1: col, r2: source.r2, c2: source.c2 };
      setFillPreview((current) =>
        current && current.r1 === next.r1 && current.r2 === next.r2 && current.c1 === next.c1 && current.c2 === next.c2
          ? current
          : next,
      );
      fillPreviewRef.current = next;
    };
    const move = (event: MouseEvent) => {
      pointer = { x: event.clientX, y: event.clientY };
      update();
    };
    const tick = () => {
      const element = scrollRef.current;
      if (pointer && element) {
        const rect = element.getBoundingClientRect();
        const dx = pointer.x < rect.left ? -18 : pointer.x > rect.right ? 18 : 0;
        const dy = pointer.y < rect.top ? -18 : pointer.y > rect.bottom ? 18 : 0;
        if (dx || dy) {
          element.scrollLeft += dx;
          element.scrollTop += dy;
          update();
        }
      }
      frame = requestAnimationFrame(tick);
    };
    const finish = () => {
      const fill = fillRef.current;
      const preview = fillPreviewRef.current;
      fillRef.current = null;
      fillPreviewRef.current = null;
      setFillPreview(null);
      if (fill && preview && (preview.r1 !== fill.source.r1 || preview.r2 !== fill.source.r2 || preview.c1 !== fill.source.c1 || preview.c2 !== fill.source.c2)) {
        onFill(fill.source, preview);
      }
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', finish);
      cancelAnimationFrame(frame);
    };
    frame = requestAnimationFrame(tick);
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', finish);
    return () => {
      cancelAnimationFrame(frame);
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', finish);
    };
    // 功能：只依赖一次拖拽的起点，拖动过程中不重建监听；时间：2026-11-01；作者：lq。
  }, [fillSession, locateCell, onFill]);

  // 功能：通过全局鼠标坐标拖选引用，跨虚拟滚动和冻结窗格仍正确定位；时间：2026-10-24；作者：lq。
  useEffect(() => {
    if (!referenceDragging) return;
    let pointer: { x: number; y: number } | null = null;
    let frame = 0;
    let lastReference = '';
    const update = () => {
      const drag = referenceDrag.current; const element = scrollRef.current;
      if (!drag || !element || !pointer) return;
      const rect = element.getBoundingClientRect();
      // 功能：CSS 缩放下将屏幕坐标换回网格逻辑像素；时间：2026-10-24；作者：lq。
      const scale = rect.width / (element.offsetWidth || rect.width) || 1;
      const x = (pointer.x - rect.left) / scale; const y = (pointer.y - rect.top) / scale;
      const col = Math.min(colsCount - 1, findIndex(colOffsets, Math.max(0, x + (x < frozenColsWidth ? 0 : element.scrollLeft))));
      const row = Math.min(rowsCount - 1, findIndex(rowOffsets, Math.max(0, y + (y < frozenRowsHeight ? 0 : element.scrollTop))));
      const next = { r1: Math.min(drag.row, row), c1: Math.min(drag.col, col), r2: Math.max(drag.row, row), c2: Math.max(drag.col, col) };
      const reference = referenceAddress(next);
      if (reference !== lastReference) { lastReference = reference; setReferenceRange(next); pickFormulaReference(drag.input, 'update', reference); }
    };
    const move = (event: MouseEvent) => { pointer = { x: event.clientX, y: event.clientY }; update(); };
    const stop = () => {
      const drag = referenceDrag.current;
      if (drag) pickFormulaReference(drag.input, 'end', '');
      referenceDrag.current = null; setReferenceDragging(false);
    };
    const tick = () => {
      const element = scrollRef.current;
      if (pointer && element) {
        const rect = element.getBoundingClientRect();
        const dx = pointer.x < rect.left ? -20 : pointer.x > rect.right ? 20 : 0;
        const dy = pointer.y < rect.top ? -20 : pointer.y > rect.bottom ? 20 : 0;
        if (dx || dy) { element.scrollLeft += dx; element.scrollTop += dy; update(); }
      }
      frame = requestAnimationFrame(tick);
    };
    frame = requestAnimationFrame(tick);
    window.addEventListener('mousemove', move); window.addEventListener('mouseup', stop);
    return () => { cancelAnimationFrame(frame); window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', stop); };
  }, [referenceDragging, colOffsets, rowOffsets, colsCount, rowsCount, frozenColsWidth, frozenRowsHeight]);

  const renderFrozenPane = () => {
    if (frozenRows === 0 && frozenCols === 0) return null;
    const panes: React.ReactNode[] = [];
    if (frozenRows > 0) {
      const cells: React.ReactNode[] = [];
      for (let row = 0; row < frozenRows; row += 1) {
        if (sheet.hiddenRows[row]) continue;
        for (let col = Math.max(frozenCols, startCol - 1); col <= endCol; col += 1) {
          if (sheet.hiddenCols[col]) continue;
          const left = (colOffsets[col] ?? 0) - scroll.left - frozenColsWidth;
          if (left + (colOffsets[col + 1] ?? 0) - (colOffsets[col] ?? 0) < 0) continue;
          cells.push(
            renderCell(row, col, { left, top: rowOffsets[row] ?? 0, width: (colOffsets[col + 1] ?? 0) - (colOffsets[col] ?? 0), height: (rowOffsets[row + 1] ?? 0) - (rowOffsets[row] ?? 0) }, true),
          );
        }
      }
      panes.push(
        <div
          key="top"
          style={{
            position: 'absolute',
            left: frozenColsWidth,
            top: 0,
            width: Math.max(0, viewport.width - frozenColsWidth),
            height: frozenRowsHeight,
            overflow: 'hidden',
            zIndex: 12,
            background: 'var(--mo-paper)',
            borderBottom: '1px solid var(--mo-line-strong)',
          }}
        >
          <div style={{ position: 'absolute', inset: 0 }}>
            {cells}
            {(() => {
              const box = clipRange(range, 0, frozenRows - 1, frozenCols, colsCount - 1);
              return box ? rangeBox(box, scroll.left + frozenColsWidth, 0, 'selection') : null;
            })()}
            {editingReference
              ? (() => {
                  const box = clipRange(editingReference, 0, frozenRows - 1, frozenCols, colsCount - 1);
                  return box ? rangeBox(box, scroll.left + frozenColsWidth, 0, 'reference') : null;
                })()
              : null}
          </div>
        </div>,
      );
    }
    if (frozenCols > 0) {
      const cells: React.ReactNode[] = [];
      for (let col = 0; col < frozenCols; col += 1) {
        if (sheet.hiddenCols[col]) continue;
        for (let row = Math.max(frozenRows, startRow - 1); row <= endRow; row += 1) {
          if (sheet.hiddenRows[row]) continue;
          const top = (rowOffsets[row] ?? 0) - scroll.top - frozenRowsHeight;
          if (top + (rowOffsets[row + 1] ?? 0) - (rowOffsets[row] ?? 0) < 0) continue;
          cells.push(
            renderCell(row, col, { left: colOffsets[col] ?? 0, top, width: (colOffsets[col + 1] ?? 0) - (colOffsets[col] ?? 0), height: (rowOffsets[row + 1] ?? 0) - (rowOffsets[row] ?? 0) }, true),
          );
        }
      }
      panes.push(
        <div
          key="left"
          style={{
            position: 'absolute',
            left: 0,
            top: frozenRowsHeight,
            width: frozenColsWidth,
            height: Math.max(0, viewport.height - frozenRowsHeight),
            overflow: 'hidden',
            zIndex: 12,
            background: 'var(--mo-paper)',
            borderRight: '1px solid var(--mo-line-strong)',
          }}
        >
          <div style={{ position: 'absolute', inset: 0 }}>
            {cells}
            {(() => {
              const box = clipRange(range, frozenRows, rowsCount - 1, 0, frozenCols - 1);
              return box ? rangeBox(box, 0, scroll.top + frozenRowsHeight, 'selection') : null;
            })()}
            {editingReference
              ? (() => {
                  const box = clipRange(editingReference, frozenRows, rowsCount - 1, 0, frozenCols - 1);
                  return box ? rangeBox(box, 0, scroll.top + frozenRowsHeight, 'reference') : null;
                })()
              : null}
          </div>
        </div>,
      );
    }
    if (frozenRows > 0 && frozenCols > 0) {
      const cells: React.ReactNode[] = [];
      for (let row = 0; row < frozenRows; row += 1) {
        if (sheet.hiddenRows[row]) continue;
        for (let col = 0; col < frozenCols; col += 1) {
          if (sheet.hiddenCols[col]) continue;
          cells.push(
            renderCell(row, col, { left: colOffsets[col] ?? 0, top: rowOffsets[row] ?? 0, width: (colOffsets[col + 1] ?? 0) - (colOffsets[col] ?? 0), height: (rowOffsets[row + 1] ?? 0) - (rowOffsets[row] ?? 0) }, true),
          );
        }
      }
      panes.push(
        <div
          key="corner"
          style={{
            position: 'absolute',
            left: 0,
            top: 0,
            width: frozenColsWidth,
            height: frozenRowsHeight,
            overflow: 'hidden',
            zIndex: 13,
            background: 'var(--mo-paper)',
          }}
        >
          {cells}
          {(() => {
            const box = clipRange(range, 0, frozenRows - 1, 0, frozenCols - 1);
            return box ? rangeBox(box, 0, 0, 'selection') : null;
          })()}
          {editingReference
            ? (() => {
                const box = clipRange(editingReference, 0, frozenRows - 1, 0, frozenCols - 1);
                return box ? rangeBox(box, 0, 0, 'reference') : null;
              })()
            : null}
        </div>,
      );
    }
    return panes;
  };

  useEffect(() => {
    const handleGlobalMouseUp = () => setDragging(false);
    window.addEventListener('mouseup', handleGlobalMouseUp);
    return () => window.removeEventListener('mouseup', handleGlobalMouseUp);
  }, []);

  const selectionLeft = colOffsets[range.c1] ?? 0;
  const selectionTop = rowOffsets[range.r1] ?? 0;
  const selectionWidth = (colOffsets[range.c2 + 1] ?? selectionLeft) - selectionLeft;
  const selectionHeight = (rowOffsets[range.r2 + 1] ?? selectionTop) - selectionTop;
  // 功能：编辑框沿用单元格自身的底色与文字色，并同样做深色纸张换算；时间：2026-10-31；作者：lq。
  const editingStyle = editing ? cellStyle(workbook, sheet.cells[cellKey(editing.row, editing.col)]) : null;
  // 功能：编辑公式时高亮光标所在的引用区域，求和结果能一眼看出统计了哪些单元格；时间：2026-11-01；作者：lq。
  const editingReference = editing ? referenceAtCaret(draft, editCaret) : null;

  /** 功能：在指定坐标系里画一个选区/引用框；时间：2026-11-01；作者：lq。 */
  const rangeBox = (target: CellRange, leftBase: number, topBase: number, tone: 'selection' | 'reference') => {
    const left = (colOffsets[target.c1] ?? 0) - leftBase;
    const top = (rowOffsets[target.r1] ?? 0) - topBase;
    const width = (colOffsets[target.c2 + 1] ?? 0) - (colOffsets[target.c1] ?? 0);
    const height = (rowOffsets[target.r2 + 1] ?? 0) - (rowOffsets[target.r1] ?? 0);
    if (width <= 0 || height <= 0) return null;
    return (
      <div
        // 功能：给选区框打标记，便于诊断脚本核对冻结窗格里是否都画出了选中框；时间：2026-11-01；作者：lq。
        data-range-box={tone}
        style={{
          position: 'absolute',
          left,
          top,
          width,
          height,
          boxSizing: 'border-box',
          border: tone === 'selection' ? '2px solid var(--mo-accent)' : '2px dashed #3478F6',
          background:
            tone === 'selection'
              ? 'color-mix(in srgb, var(--mo-accent) 8%, transparent)'
              : 'rgba(52,120,246,0.08)',
          pointerEvents: 'none',
          zIndex: 5,
        }}
      />
    );
  };

  /** 功能：把选区裁剪到某个窗格范围内，冻结区域同样显示选中框；时间：2026-11-01；作者：lq。 */
  const clipRange = (target: CellRange, rowFrom: number, rowTo: number, colFrom: number, colTo: number): CellRange | null => {
    const clipped = {
      r1: Math.max(target.r1, rowFrom),
      r2: Math.min(target.r2, rowTo),
      c1: Math.max(target.c1, colFrom),
      c2: Math.min(target.c2, colTo),
    };
    return clipped.r1 > clipped.r2 || clipped.c1 > clipped.c2 ? null : clipped;
  };

  return (
    <div
      ref={containerRef}
      data-sheet-grid={gridId}
      tabIndex={0}
      className="mo-sheet relative flex h-full min-h-0 flex-col outline-none"
      onKeyDown={handleKeyDown}
    >
      <div className="flex shrink-0" style={{ height: COL_HEADER_HEIGHT }}>
        <div
          style={{
            width: ROW_HEADER_WIDTH,
            height: COL_HEADER_HEIGHT,
            background: 'var(--mo-surface-2)',
            borderRight: '1px solid var(--mo-line)',
            borderBottom: '1px solid var(--mo-line)',
            boxSizing: 'border-box',
          }}
          aria-label="全选工作表"
          onMouseDown={(event) => {
            event.preventDefault();
            containerRef.current?.focus({ preventScroll: true });
            onSelectionChange({ kind: 'all', anchor: { row: 0, col: 0 }, focus: { row: rowsCount - 1, col: colsCount - 1 } });
          }}
        />
        <div className="relative flex-1 overflow-hidden">
          {/* 功能：冻结列的表头固定在左侧，不再随水平滚动一起移出；时间：2026-11-01；作者：lq。 */}
          <div style={{ position: 'absolute', left: 0, top: 0, width: totalWidth, height: COL_HEADER_HEIGHT }}>
            {[...visibleCols.filter((col) => col >= frozenCols), ...Array.from({ length: frozenCols }, (_, index) => index)].map((col) => {
              const left = (colOffsets[col] ?? 0) - (col < frozenCols ? 0 : scroll.left);
              const width = (colOffsets[col + 1] ?? left) - left;
              if (width <= 0) return null;
              const selected = col >= range.c1 && col <= range.c2;
              // 功能：阅读模式下把当前单元格所在列的表头也点亮，配合行表头形成十字参考；时间：2026-10-30；作者：lq。
              const crossActive = crosshair && col === active.col;
              return (
                <div
                  key={col}
                  data-cross-active={crossActive ? 'col' : undefined}
                  style={{
                    ...headerBase,
                    left,
                    top: 0,
                    width,
                    height: COL_HEADER_HEIGHT,
                    color: selected || crossActive ? 'var(--mo-accent)' : undefined,
                    background: crossActive
                      ? 'color-mix(in srgb, var(--mo-accent) 16%, var(--mo-surface-2))'
                      : selected
                        ? 'var(--mo-accent-soft)'
                        : undefined,
                    boxShadow: crossActive ? 'inset 0 -2px 0 var(--mo-accent)' : undefined,
                  }}
                  onMouseDown={(event) => {
                    event.preventDefault();
                    containerRef.current?.focus({ preventScroll: true });
                    onSelectionChange({ kind: 'col', anchor: { row: 0, col }, focus: { row: rowsCount - 1, col } });
                  }}
                  onContextMenu={(event) => {
                    event.preventDefault();
                    onHeaderContextMenu(event, 'col', col);
                  }}
                >
                  {columnLabel(col)}
                  <div
                    style={{
                      position: 'absolute',
                      right: -2,
                      top: 0,
                      width: 5,
                      height: '100%',
                      cursor: 'col-resize',
                    }}
                    onDoubleClick={(event) => { event.preventDefault(); event.stopPropagation(); autoFit('col', col); }}
                    onMouseDown={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      resizeRef.current = {
                        kind: 'col',
                        index: col,
                        start: event.clientX,
                        initial: width,
                      };
                      setResizePreview({ kind: 'col', index: col, size: width });
                    }}
                  />
                </div>
              );
            })}
          </div>
        </div>
      </div>

      <div className="flex min-h-0 flex-1">
        <div className="relative shrink-0 overflow-hidden" style={{ width: ROW_HEADER_WIDTH }}>
          {/* 功能：冻结行的行号固定在顶部，避免滚动后冻结区左侧出现空白条；时间：2026-11-01；作者：lq。 */}
          <div style={{ position: 'absolute', left: 0, top: 0, width: ROW_HEADER_WIDTH, height: totalHeight }}>
            {[...visibleRows.filter((row) => row >= frozenRows), ...Array.from({ length: frozenRows }, (_, index) => index)].map((row) => {
              const top = (rowOffsets[row] ?? 0) - (row < frozenRows ? 0 : scroll.top);
              const height = (rowOffsets[row + 1] ?? top) - top;
              if (height <= 0) return null;
              const selected = row >= range.r1 && row <= range.r2;
              const crossActive = crosshair && row === active.row;
              return (
                <div
                  key={row}
                  data-cross-active={crossActive ? 'row' : undefined}
                  style={{
                    ...headerBase,
                    left: 0,
                    top,
                    width: ROW_HEADER_WIDTH,
                    height,
                    color: selected || crossActive ? 'var(--mo-accent)' : undefined,
                    background: crossActive
                      ? 'color-mix(in srgb, var(--mo-accent) 16%, var(--mo-surface-2))'
                      : selected
                        ? 'var(--mo-accent-soft)'
                        : undefined,
                    boxShadow: crossActive ? 'inset -2px 0 0 var(--mo-accent)' : undefined,
                  }}
                  onMouseDown={(event) => {
                    event.preventDefault();
                    containerRef.current?.focus({ preventScroll: true });
                    onSelectionChange({ kind: 'row', anchor: { row, col: 0 }, focus: { row, col: colsCount - 1 } });
                  }}
                  onContextMenu={(event) => {
                    event.preventDefault();
                    onHeaderContextMenu(event, 'row', row);
                  }}
                >
                  {row + 1}
                  <div
                    style={{
                      position: 'absolute',
                      bottom: -2,
                      left: 0,
                      width: '100%',
                      height: 5,
                      cursor: 'row-resize',
                    }}
                    onDoubleClick={(event) => { event.preventDefault(); event.stopPropagation(); autoFit('row', row); }}
                    onMouseDown={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      resizeRef.current = {
                        kind: 'row',
                        index: row,
                        start: event.clientY,
                        initial: height,
                      };
                      setResizePreview({ kind: 'row', index: row, size: height });
                    }}
                  />
                </div>
              );
            })}
          </div>
        </div>

        <div className="relative min-w-0 flex-1">
      <div
        ref={(element) => {
          scrollRef.current = element;
          if (scrollElementRef) scrollElementRef.current = element;
        }}
            className="absolute inset-0 overflow-auto"
            onScroll={handleScroll}
          >
            <div
              style={{
                position: 'relative',
                width: totalWidth,
                height: totalHeight,
                minWidth: '100%',
                minHeight: '100%',
              }}
            >
              {visibleRows.map((row) =>
                visibleCols.map((col) => {
                  const left = colOffsets[col] ?? 0;
                  const top = rowOffsets[row] ?? 0;
                  const width = (colOffsets[col + 1] ?? left) - left;
                  const height = (rowOffsets[row + 1] ?? top) - top;
                  if (width <= 0 || height <= 0) return null;
                  return renderCell(row, col, { left, top, width, height });
                }),
              )}

              {/* 功能：以蓝色边框显示正在点选或拖选的公式引用区域；时间：2026-10-24；作者：lq。 */}
              {referenceRange ? <div data-formula-reference={referenceAddress(referenceRange)} style={{ position: 'absolute', left: colOffsets[referenceRange.c1] ?? 0, top: rowOffsets[referenceRange.r1] ?? 0,
                width: (colOffsets[referenceRange.c2 + 1] ?? 0) - (colOffsets[referenceRange.c1] ?? 0), height: (rowOffsets[referenceRange.r2 + 1] ?? 0) - (rowOffsets[referenceRange.r1] ?? 0),
                border: '2px dashed #3478F6', background: 'rgba(52,120,246,0.08)', pointerEvents: 'none', zIndex: 14 }} /> : null}
              {/* 功能：编辑公式时高亮当前引用，例如 =SUM(B4:B11) 直接框出求和范围；时间：2026-11-01；作者：lq。 */}
              {editingReference ? (
                <div
                  data-editing-reference={referenceAddress(editingReference)}
                  style={{
                    position: 'absolute',
                    left: colOffsets[editingReference.c1] ?? 0,
                    top: rowOffsets[editingReference.r1] ?? 0,
                    width: (colOffsets[editingReference.c2 + 1] ?? 0) - (colOffsets[editingReference.c1] ?? 0),
                    height: (rowOffsets[editingReference.r2 + 1] ?? 0) - (rowOffsets[editingReference.r1] ?? 0),
                    border: '2px dashed #3478F6',
                    background: 'rgba(52,120,246,0.08)',
                    pointerEvents: 'none',
                    zIndex: 14,
                  }}
                />
              ) : null}
              {/* 功能：图片与网格共用滚动坐标，拖动和缩放结束后保存；时间：2026-10-24；作者：lq。 */}
              <SheetImageLayer images={sheet.images ?? []} colOffsets={colOffsets} rowOffsets={rowOffsets} readonly={!!props.readonly} zoom={props.zoom ?? 1}
                onChange={(id, patch) => props.onImageChange?.(id, patch)} onDelete={(id) => props.onImageDelete?.(id)} />
              {/* 功能：阅读模式的行列参考带，浅色叠加不遮挡内容，打印与导出不包含；时间：2026-10-30；作者：lq。 */}
              {crosshair ? (
                <>
                  <div
                    data-crosshair="row"
                    style={{
                      position: 'absolute',
                      left: 0,
                      top: rowOffsets[active.row] ?? 0,
                      width: totalWidth,
                      height: Math.max(1, (rowOffsets[active.row + 1] ?? 0) - (rowOffsets[active.row] ?? 0)),
                      background: 'color-mix(in srgb, var(--mo-accent) 7%, transparent)',
                      pointerEvents: 'none',
                      zIndex: 3,
                    }}
                  />
                  <div
                    data-crosshair="col"
                    style={{
                      position: 'absolute',
                      left: colOffsets[active.col] ?? 0,
                      top: 0,
                      width: Math.max(1, (colOffsets[active.col + 1] ?? 0) - (colOffsets[active.col] ?? 0)),
                      height: totalHeight,
                      background: 'color-mix(in srgb, var(--mo-accent) 7%, transparent)',
                      pointerEvents: 'none',
                      zIndex: 3,
                    }}
                  />
                </>
              ) : null}
              <div
                style={{
                  position: 'absolute',
                  left: selectionLeft,
                  top: selectionTop,
                  width: Math.max(1, selectionWidth),
                  height: Math.max(1, selectionHeight),
                  boxSizing: 'border-box',
                  border: '2px solid var(--mo-accent)',
                  background: 'color-mix(in srgb, var(--mo-accent) 8%, transparent)',
                  pointerEvents: 'none',
                  zIndex: 5,
                }}
              />

              {fillPreview ? (
                <div
                  style={{
                    position: 'absolute',
                    left: colOffsets[fillPreview.c1] ?? 0,
                    top: rowOffsets[fillPreview.r1] ?? 0,
                    width: (colOffsets[fillPreview.c2 + 1] ?? 0) - (colOffsets[fillPreview.c1] ?? 0),
                    height: (rowOffsets[fillPreview.r2 + 1] ?? 0) - (rowOffsets[fillPreview.r1] ?? 0),
                    border: '1px dashed var(--mo-accent)',
                    background: 'color-mix(in srgb, var(--mo-accent) 6%, transparent)',
                    pointerEvents: 'none',
                    zIndex: 8,
                  }}
                />
              ) : null}

              <div
                title="拖动填充"
                style={{
                  position: 'absolute',
                  left: selectionLeft + Math.max(1, selectionWidth) - 4,
                  top: selectionTop + Math.max(1, selectionHeight) - 4,
                  width: 7,
                  height: 7,
                  background: 'var(--mo-accent)',
                  // 功能：填充柄外圈用纸张色描边，深色纸下不会出现刺眼的白框；时间：2026-10-31；作者：lq。
                  border: '1px solid var(--mo-paper)',
                  cursor: 'crosshair',
                  zIndex: 9,
                }}
                onMouseDown={(event) => {
                  event.preventDefault();
                  event.stopPropagation();
                  const source = selectionRange(selection);
                  fillRef.current = { source, startX: event.clientX, startY: event.clientY };
                  fillPreviewRef.current = source;
                  setFillPreview(source);
                  // 功能：递增会话号触发一次监听注册，之后由全局鼠标事件驱动预览；时间：2026-11-01；作者：lq。
                  setFillSession((value) => value + 1);
                }}
              />

              {editing ? (
                <FormulaInput
                  plain
                  aria-label="单元格编辑"
                  ref={inputRef}
                  value={draft}
                  onValueChange={(value) => { setDraft(value); props.onDraftChange?.(value); }}
                  onSelect={(event) => setEditCaret(event.currentTarget.selectionStart ?? 0)}
                  onKeyUp={(event) => setEditCaret(event.currentTarget.selectionStart ?? 0)}
                  onClick={(event) => setEditCaret(event.currentTarget.selectionStart ?? 0)}
                  onKeyDown={(event) => {
                    if (event.key === 'Enter') {
                      event.preventDefault();
                      commitEditing(event.shiftKey ? 'up' : 'down');
                    } else if (event.key === 'Tab') {
                      event.preventDefault();
                      commitEditing(event.shiftKey ? 'left' : 'right');
                    } else if (event.key === 'Escape') {
                      event.preventDefault();
                      onCancelEdit();
                    }
                  }}
                  onBlur={(event) => {
                    if ((event.relatedTarget as HTMLElement | null)?.hasAttribute('data-sheet-formula-input')) return;
                    commitEditing('none');
                  }}
                  style={{
                    position: 'absolute',
                    left: colOffsets[editing.col] ?? 0,
                    top: rowOffsets[editing.row] ?? 0,
                    width: Math.max(24, (colOffsets[editing.col + 1] ?? 0) - (colOffsets[editing.col] ?? 0)),
                    height: Math.max(22, (rowOffsets[editing.row + 1] ?? 0) - (rowOffsets[editing.row] ?? 0)),
                    padding: '0 4px',
                    border: '2px solid var(--mo-accent)',
                    outline: 'none',
                    background: 'var(--mo-cell-fill)',
                    color: 'var(--mo-cell-ink)',
                    '--mo-cell-fill': paperFill(editingStyle?.fill, darkPaper),
                    '--mo-cell-fill-plain': editingStyle?.fill ?? 'var(--mo-paper)',
                    '--mo-cell-ink': paperInk(editingStyle?.color, darkPaper),
                    '--mo-cell-ink-plain': editingStyle?.color ?? 'var(--mo-paper-ink)',
                    fontSize: `${editingStyle?.fontSizePt ?? workbook.defaultFontSizePt}pt`,
                    fontFamily: fontStackFor(editingStyle?.fontFamily ?? workbook.defaultFont),
                    fontWeight: editingStyle?.bold ? 700 : 400,
                    textAlign: editingStyle?.horizontalAlign === 'right' ? 'right' : editingStyle?.horizontalAlign === 'center' ? 'center' : 'left',
                    boxSizing: 'border-box',
                    zIndex: 20,
                  } as CSSProperties}
                />
              ) : null}
            </div>
          </div>

          {renderFrozenPane()}
        </div>
      </div>
    </div>
  );
}

// 功能：按单元格实际宽度缩小文字显示，保留原始字号与内容；时间：2026-10-24；作者：lq。
function SheetCellText({ text, shrink, width, font }: { text: string; shrink: boolean; width: number; font: string }) {
  const ref = useRef<HTMLSpanElement>(null);
  useLayoutEffect(() => {
    const span = ref.current;
    if (!span) return;
    span.style.fontSize = '';
    const available = (span.parentElement?.clientWidth ?? width) - 8;
    if (shrink && span.scrollWidth > available && available > 0) span.style.fontSize = `${Math.max(1, Math.floor(100 * available / span.scrollWidth))}%`;
  }, [text, shrink, width, font]);
  return <span ref={ref} style={{ overflow: 'hidden', textOverflow: 'clip', whiteSpace: 'inherit', maxWidth: '100%' }}>{text}</span>;
}
