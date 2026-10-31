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
  zoom?: number;
  onImageChange?: (id: string, patch: Partial<SheetImage>) => void;
  onImageDelete?: (id: string) => void;
  onFill: (source: CellRange, target: CellRange) => void;
  /** 高亮当前行/列（对应活动单元格）。 */
  highlightRowCol?: boolean;
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
  const width = side.style === 'thick' ? 3 : side.style === 'medium' || side.style === 'double' ? 2 : 1;
  // 功能：将 Excel 的“细/中/粗”边框语义转换为浏览器可识别的实线；时间：2026-10-24；作者：lq。
  const lineStyle = side.style === 'double' ? 'double' : side.style === 'dashed' ? 'dashed' : side.style === 'dotted' ? 'dotted' : 'solid';
  return `${width}px ${lineStyle} ${side.color ?? '#000000'}`;
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
    highlightRowCol,
  } = props;

  const containerRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement | null>(null);
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

  // 功能：在绘制前初始化编辑文本，避免快速输入时光标落在等号前面；时间：2026-10-24；作者：lq。
  useLayoutEffect(() => {
    if (!editing) {
      focusedEditingRef.current = null;
      return;
    }
    setDraft(editing.initialText);
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
    if (!resizeRef.current && !fillRef.current) return;
    const onMove = (event: MouseEvent) => {
      const resize = resizeRef.current;
      if (resize) {
        const delta = resize.kind === 'col' ? event.clientX - resize.start : event.clientY - resize.start;
        setResizePreview({
          kind: resize.kind,
          index: resize.index,
          size: Math.max(resize.kind === 'col' ? 24 : 14, resize.initial + delta),
        });
        return;
      }
      const fill = fillRef.current;
      if (fill && scrollRef.current) {
        const rect = scrollRef.current.getBoundingClientRect();
        const x = event.clientX - rect.left + scrollRef.current.scrollLeft;
        const y = event.clientY - rect.top + scrollRef.current.scrollTop;
        const col = findIndex(colOffsets, Math.max(0, x));
        const row = findIndex(rowOffsets, Math.max(0, y));
        const source = fill.source;
        let next: CellRange;
        if (Math.abs(row - source.r2) >= Math.abs(col - source.c2)) {
          next = row >= source.r2
            ? { r1: source.r1, c1: source.c1, r2: row, c2: source.c2 }
            : { r1: row, c1: source.c1, r2: source.r2, c2: source.c2 };
        } else {
          next = col >= source.c2
            ? { r1: source.r1, c1: source.c1, r2: source.r2, c2: col }
            : { r1: source.r1, c1: col, r2: source.r2, c2: source.c2 };
        }
        setFillPreview(next);
      }
    };
    const onUp = () => {
      const resize = resizeRef.current;
      if (resize && resizePreview) {
        if (resize.kind === 'col') onColumnResize(resize.index, resizePreview.size);
        else onRowResize(resize.index, resizePreview.size);
      }
      resizeRef.current = null;
      setResizePreview(null);
      const fill = fillRef.current;
      if (fill && fillPreview) {
        onFill(fill.source, fillPreview);
      }
      fillRef.current = null;
      setFillPreview(null);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
  }, [fillPreview, onColumnResize, onFill, onRowResize, resizePreview]);

  useEffect(() => {
    const scrollElement = scrollRef.current;
    // 功能：表头选择仅改变选区，保持当前水平和垂直滚动位置；时间：2026-10-24；作者：lq。
    if (!scrollElement || selection.kind) return;
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
  }, [active.col, active.row, colOffsets, rowOffsets, selection.kind]);

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
      background: style.fill ?? '#FFFFFF',
      color: style.color ?? '#1D2939',
      fontSize: `${style.fontSizePt ?? workbook.defaultFontSizePt}pt`,
      fontFamily: style.fontFamily ?? workbook.defaultFont,
      fontWeight: style.bold ? 700 : 400,
      fontStyle: style.italic ? 'italic' : 'normal',
      textDecoration: [style.underline ? 'underline' : '', style.strike ? 'line-through' : '']
        .filter(Boolean)
        .join(' ') || undefined,
      whiteSpace: style.wrapText ? 'pre-wrap' : 'nowrap',
      wordBreak: style.wrapText ? 'break-word' : 'normal',
      borderRight: borderCss(style.borders?.right) ?? (sheet.showGridLines ? '1px solid rgba(228,233,242,0.9)' : undefined),
      borderBottom: borderCss(style.borders?.bottom) ?? (sheet.showGridLines ? '1px solid rgba(228,233,242,0.9)' : undefined),
      borderTop: borderCss(style.borders?.top),
      borderLeft: borderCss(style.borders?.left),
      outline: isActive ? '2px solid var(--mo-accent)' : undefined,
      outlineOffset: isActive ? '-2px' : undefined,
      zIndex: isActive ? 7 : origin ? 3 : 1,
      cursor: 'cell',
    };
    const displayText = style.shrinkToFit ? text : text;
    return (
      <div
        key={key}
        data-cell={key}
        data-row={row}
        data-col={col}
        style={cellCss}
        title={text.length > 30 ? text : undefined}
        onMouseDown={(event) => {
          if (event.button !== 0) return;
          event.preventDefault();
          if (beginReference(row, col)) return;
          containerRef.current?.focus({ preventScroll: true });
          beginSelection(row, col, event.shiftKey);
        }}
        onMouseEnter={() => {
          if (dragging) onSelectionChange({ anchor: selection.anchor, focus: { row, col } }, { extend: true });
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
        <SheetCellText text={displayText} shrink={!!style.shrinkToFit && !style.wrapText} width={width} font={`${style.fontFamily}-${style.fontSizePt}-${style.bold}`} />
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
        for (let col = frozenCols; col <= endCol; col += 1) {
          const left = (colOffsets[col] ?? 0) - scroll.left;
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
            background: '#fff',
            borderBottom: '1px solid var(--mo-line-strong)',
          }}
        >
          <div style={{ position: 'absolute', inset: 0 }}>{cells}</div>
        </div>,
      );
    }
    if (frozenCols > 0) {
      const cells: React.ReactNode[] = [];
      for (let col = 0; col < frozenCols; col += 1) {
        for (let row = frozenRows; row <= endRow; row += 1) {
          const top = (rowOffsets[row] ?? 0) - scroll.top;
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
            background: '#fff',
            borderRight: '1px solid var(--mo-line-strong)',
          }}
        >
          <div style={{ position: 'absolute', inset: 0 }}>{cells}</div>
        </div>,
      );
    }
    if (frozenRows > 0 && frozenCols > 0) {
      const cells: React.ReactNode[] = [];
      for (let row = 0; row < frozenRows; row += 1) {
        for (let col = 0; col < frozenCols; col += 1) {
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
            background: '#fff',
          }}
        >
          {cells}
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

  return (
    <div
      ref={containerRef}
      data-sheet-grid={gridId}
      tabIndex={0}
      className="relative flex h-full min-h-0 flex-col bg-[var(--mo-surface)] outline-none"
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
          <div style={{ position: 'absolute', left: 0, top: 0, width: totalWidth, height: COL_HEADER_HEIGHT, transform: `translateX(${-scroll.left}px)` }}>
            {visibleCols.map((col) => {
              const left = colOffsets[col] ?? 0;
              const width = (colOffsets[col + 1] ?? left) - left;
              if (width <= 0) return null;
              const selected = col >= range.c1 && col <= range.c2;
              return (
                <div
                  key={col}
                  style={{
                    ...headerBase,
                    left,
                    top: 0,
                    width,
                    height: COL_HEADER_HEIGHT,
                    color: selected ? 'var(--mo-accent)' : undefined,
                    background: selected ? 'var(--mo-accent-soft)' : undefined,
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
          <div style={{ position: 'absolute', left: 0, top: 0, width: ROW_HEADER_WIDTH, height: totalHeight, transform: `translateY(${-scroll.top}px)` }}>
            {visibleRows.map((row) => {
              const top = rowOffsets[row] ?? 0;
              const height = (rowOffsets[row + 1] ?? top) - top;
              if (height <= 0) return null;
              const selected = row >= range.r1 && row <= range.r2;
              return (
                <div
                  key={row}
                  style={{
                    ...headerBase,
                    left: 0,
                    top,
                    width: ROW_HEADER_WIDTH,
                    height,
                    color: selected ? 'var(--mo-accent)' : undefined,
                    background: selected ? 'var(--mo-accent-soft)' : undefined,
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

              {/* 功能：高亮当前行/列（十字架样式，行带与列带加边缘线）；时间：2026-10-30；作者：hou0。 */}
              {highlightRowCol && active.row >= 0 && active.col >= 0 ? (
                <>
                  <div style={{ position: 'absolute', left: 0, top: rowOffsets[active.row] ?? 0, width: totalWidth,
                    height: (rowOffsets[active.row + 1] ?? 0) - (rowOffsets[active.row] ?? 0),
                    background: 'color-mix(in srgb, var(--mo-accent) 12%, transparent)',
                    borderTop: '1px solid color-mix(in srgb, var(--mo-accent) 45%, transparent)',
                    borderBottom: '1px solid color-mix(in srgb, var(--mo-accent) 45%, transparent)',
                    boxSizing: 'border-box', pointerEvents: 'none', zIndex: 4 }} />
                  <div style={{ position: 'absolute', left: colOffsets[active.col] ?? 0, top: 0,
                    width: (colOffsets[active.col + 1] ?? 0) - (colOffsets[active.col] ?? 0), height: totalHeight,
                    background: 'color-mix(in srgb, var(--mo-accent) 12%, transparent)',
                    borderLeft: '1px solid color-mix(in srgb, var(--mo-accent) 45%, transparent)',
                    borderRight: '1px solid color-mix(in srgb, var(--mo-accent) 45%, transparent)',
                    boxSizing: 'border-box', pointerEvents: 'none', zIndex: 4 }} />
                </>
              ) : null}

              {/* 功能：以蓝色边框显示正在点选或拖选的公式引用区域；时间：2026-10-24；作者：lq。 */}
              {referenceRange ? <div data-formula-reference={referenceAddress(referenceRange)} style={{ position: 'absolute', left: colOffsets[referenceRange.c1] ?? 0, top: rowOffsets[referenceRange.r1] ?? 0,
                width: (colOffsets[referenceRange.c2 + 1] ?? 0) - (colOffsets[referenceRange.c1] ?? 0), height: (rowOffsets[referenceRange.r2 + 1] ?? 0) - (rowOffsets[referenceRange.r1] ?? 0),
                border: '2px dashed #3478F6', background: 'rgba(52,120,246,0.08)', pointerEvents: 'none', zIndex: 14 }} /> : null}
              {/* 功能：图片与网格共用滚动坐标，拖动和缩放结束后保存；时间：2026-10-24；作者：lq。 */}
              <SheetImageLayer images={sheet.images ?? []} colOffsets={colOffsets} rowOffsets={rowOffsets} readonly={!!props.readonly} zoom={props.zoom ?? 1}
                onChange={(id, patch) => props.onImageChange?.(id, patch)} onDelete={(id) => props.onImageDelete?.(id)} />
              <div
                style={{
                  position: 'absolute',
                  left: selectionLeft,
                  top: selectionTop,
                  width: Math.max(1, selectionWidth),
                  height: Math.max(1, selectionHeight),
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
                  border: '1px solid #fff',
                  cursor: 'crosshair',
                  zIndex: 9,
                }}
                onMouseDown={(event) => {
                  event.preventDefault();
                  event.stopPropagation();
                  const source = selectionRange(selection);
                  const rect = scrollRef.current?.getBoundingClientRect();
                  fillRef.current = {
                    source,
                    startX: event.clientX - (rect?.left ?? 0),
                    startY: event.clientY - (rect?.top ?? 0),
                  };
                  setFillPreview(source);
                }}
              />

              {editing ? (
                <FormulaInput
                  plain
                  aria-label="单元格编辑"
                  ref={inputRef}
                  value={draft}
                  onValueChange={(value) => { setDraft(value); props.onDraftChange?.(value); }}
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
                    minWidth: Math.max(80, (colOffsets[editing.col + 1] ?? 0) - (colOffsets[editing.col] ?? 0)),
                    height: Math.max(22, (rowOffsets[editing.row + 1] ?? 0) - (rowOffsets[editing.row] ?? 0)),
                    padding: '0 4px',
                    border: '2px solid var(--mo-accent)',
                    outline: 'none',
                    background: '#fff',
                    color: '#1D2939',
                    fontSize: 13,
                    fontFamily: 'var(--mo-font-ui)',
                    boxSizing: 'border-box',
                    zIndex: 20,
                    boxShadow: '0 2px 8px rgba(0,0,0,0.12)',
                  }}
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
