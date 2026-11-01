import {
  cellKey,
  normalizeRange,
  type Cell,
  type CellBorders,
  type CellRange,
  type CellStyle,
  type SheetModel,
  type WorkbookModel,
} from '@mini-office/document-model';
import { columnLabel, parseCellAddress, uid } from '@mini-office/shared';
import { formatCellValue, isDateFormat } from './format';
import { dateToSerial, parseDateInput } from './dates';
import { translateFormula } from './formula/parser';
import { editWorkbookReferences, moveSortedFormula } from './formula/references';

export interface CellCoord {
  row: number;
  col: number;
}

// 功能：区分表头选区与单元格选区，防止选区末端触发视图跳转；时间：2026-10-24；作者：lq。
export interface Selection {
  kind?: 'row' | 'col' | 'all';
  anchor: CellCoord;
  focus: CellCoord;
}

export function selectionRange(selection: Selection): CellRange {
  return normalizeRange({
    r1: selection.anchor.row,
    c1: selection.anchor.col,
    r2: selection.focus.row,
    c2: selection.focus.col,
  });
}

export function rangeRef(range: CellRange): string {
  const normalized = normalizeRange(range);
  return `${columnLabel(normalized.c1)}${normalized.r1 + 1}:${columnLabel(normalized.c2)}${normalized.r2 + 1}`;
}

export function parseRangeRef(ref: string): CellRange | null {
  const parts = ref.split(':');
  const first = parseCellAddress(parts[0]);
  const last = parseCellAddress(parts[1] ?? parts[0]);
  if (!first || !last) return null;
  return normalizeRange({ r1: first.row, c1: first.col, r2: last.row, c2: last.col });
}

export function rangeArea(range: CellRange): number {
  const normalized = normalizeRange(range);
  return (normalized.r2 - normalized.r1 + 1) * (normalized.c2 - normalized.c1 + 1);
}

export function clampRange(range: CellRange, rows: number, cols: number): CellRange {
  const normalized = normalizeRange(range);
  return {
    r1: Math.max(0, Math.min(rows - 1, normalized.r1)),
    c1: Math.max(0, Math.min(cols - 1, normalized.c1)),
    r2: Math.max(0, Math.min(rows - 1, normalized.r2)),
    c2: Math.max(0, Math.min(cols - 1, normalized.c2)),
  };
}

export function intersectRanges(a: CellRange, b: CellRange): CellRange | null {
  const range = {
    r1: Math.max(a.r1, b.r1),
    c1: Math.max(a.c1, b.c1),
    r2: Math.min(a.r2, b.r2),
    c2: Math.min(a.c2, b.c2),
  };
  return range.r1 <= range.r2 && range.c1 <= range.c2 ? range : null;
}

export function getSheet(workbook: WorkbookModel, sheetId: string): SheetModel | undefined {
  return workbook.sheets.find((sheet) => sheet.id === sheetId);
}

export function activeSheet(workbook: WorkbookModel): SheetModel {
  return workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId) ?? workbook.sheets[0];
}

export function updateSheet(
  workbook: WorkbookModel,
  sheetId: string,
  updater: (sheet: SheetModel) => SheetModel,
): WorkbookModel {
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => (sheet.id === sheetId ? updater(sheet) : sheet)),
  };
}

export function cellAt(sheet: SheetModel, row: number, col: number): Cell | undefined {
  return sheet.cells[cellKey(row, col)];
}

export function cellStyle(workbook: WorkbookModel, cell: Cell | undefined): CellStyle {
  return (cell?.styleId ? workbook.styles[cell.styleId] : undefined) ?? workbook.styles.default ?? workbook.styles.s0 ?? {};
}

export function cellDisplay(workbook: WorkbookModel, sheet: SheetModel, row: number, col: number): string {
  const computed = workbook.computed?.[sheet.id]?.[cellKey(row, col)];
  if (computed) return computed.display;
  const cell = cellAt(sheet, row, col);
  if (!cell) return '';
  return formatCellValue(cell.value, cellStyle(workbook, cell), workbook.date1904);
}

export function cellRawInput(cell: Cell | undefined): string {
  if (!cell) return '';
  if (cell.formula) return cell.formula;
  if (cell.value === null || cell.value === undefined) return '';
  if (typeof cell.value === 'boolean') return cell.value ? 'TRUE' : 'FALSE';
  return String(cell.value);
}

function styleKey(style: CellStyle): string {
  const { xfIndex: _xfIndex, ...rest } = style;
  return JSON.stringify(rest);
}

export function findOrCreateStyle(
  workbook: WorkbookModel,
  style: CellStyle,
): { workbook: WorkbookModel; styleId: string } {
  const key = styleKey(style);
  for (const [id, existing] of Object.entries(workbook.styles)) {
    if (styleKey(existing) === key) return { workbook, styleId: id };
  }
  const styleId = `style-${uid()}`;
  return {
    workbook: { ...workbook, styles: { ...workbook.styles, [styleId]: style } },
    styleId,
  };
}

interface CellUpdate {
  row: number;
  col: number;
  cell: Cell | null;
}

export function applyCellUpdates(
  workbook: WorkbookModel,
  sheetId: string,
  updates: CellUpdate[],
): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => {
    const cells = { ...sheet.cells };
    let rows = sheet.dimension.rows;
    let cols = sheet.dimension.cols;
    for (const update of updates) {
      const key = cellKey(update.row, update.col);
      if (update.cell) {
        cells[key] = update.cell;
        rows = Math.max(rows, update.row + 1);
        cols = Math.max(cols, update.col + 1);
      } else {
        delete cells[key];
      }
    }
    return {
      ...sheet,
      cells,
      dimension: {
        rows: Math.min(1_048_576, Math.max(rows + 20, 100)),
        cols: Math.min(16_384, Math.max(cols + 10, 26)),
      },
    };
  });
}

export function createCellFromInput(input: string, workbook: WorkbookModel, existing?: Cell): Cell {
  const styleId = existing?.styleId;
  if (input.startsWith('=')) return { value: null, formula: input, styleId };
  if (input === '') return { value: null, styleId };
  const trimmed = input.trim();
  if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(trimmed)) {
    return { value: Number(trimmed), styleId };
  }
  if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)%$/.test(trimmed)) {
    return { value: Number(trimmed.slice(0, -1)) / 100, styleId };
  }
  if (/^TRUE$/i.test(trimmed)) return { value: true, styleId };
  if (/^FALSE$/i.test(trimmed)) return { value: false, styleId };
  if (existing?.styleId) {
    const style = workbook.styles[existing.styleId];
    if (style?.numFmt && isDateFormat(style.numFmt)) {
      const date = parseDateInput(trimmed);
      if (date) return { value: dateToSerial(date, workbook.date1904), styleId };
    }
  }
  return { value: input, styleId };
}

export function setCellInput(
  workbook: WorkbookModel,
  sheetId: string,
  row: number,
  col: number,
  input: string,
): WorkbookModel {
  const sheet = getSheet(workbook, sheetId);
  if (!sheet) return workbook;
  const existing = sheet.cells[cellKey(row, col)];
  const cell = createCellFromInput(input, workbook, existing);
  const isEmpty = (cell.value === null || cell.value === undefined) && !cell.formula;
  if (isEmpty && !cell.styleId) {
    return applyCellUpdates(workbook, sheetId, [{ row, col, cell: null }]);
  }
  return applyCellUpdates(workbook, sheetId, [{ row, col, cell }]);
}

export function clearRangeContents(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
): WorkbookModel {
  const sheet = getSheet(workbook, sheetId);
  if (!sheet) return workbook;
  const normalized = normalizeRange(range);
  const updates: CellUpdate[] = [];
  for (let row = normalized.r1; row <= normalized.r2; row += 1) {
    for (let col = normalized.c1; col <= normalized.c2; col += 1) {
      const existing = sheet.cells[cellKey(row, col)];
      if (!existing) continue;
      updates.push({
        row,
        col,
        cell: existing.styleId ? { value: null, styleId: existing.styleId } : null,
      });
    }
  }
  return applyCellUpdates(workbook, sheetId, updates);
}

export function applyStyleToRange(
  workbook: WorkbookModel,
  sheetId: string,
  range: CellRange,
  patch: Partial<CellStyle>,
): WorkbookModel {
  const sheet = getSheet(workbook, sheetId);
  if (!sheet) return workbook;
  const normalized = normalizeRange(range);
  let next = workbook;
  const styleCache = new Map<string, string>();
  const updates: CellUpdate[] = [];
  for (let row = normalized.r1; row <= normalized.r2; row += 1) {
    for (let col = normalized.c1; col <= normalized.c2; col += 1) {
      const existing = sheet.cells[cellKey(row, col)];
      const base = cellStyle(next, existing);
      // 功能：修改单条边框时保留单元格其余三条边，避免选区格式操作互相覆盖；时间：2026-10-24；作者：lq。
      const merged = {
        ...base,
        ...patch,
        ...(patch.borders && base.borders
          ? { borders: { ...base.borders, ...patch.borders } }
          : {}),
      } as CellStyle;
      let styleId = styleCache.get(styleKey(merged));
      if (!styleId) {
        const found = findOrCreateStyle(next, merged);
        next = found.workbook;
        styleId = found.styleId;
        styleCache.set(styleKey(merged), styleId);
      }
      updates.push({
        row,
        col,
        cell: { value: existing?.value ?? null, formula: existing?.formula, styleId },
      });
    }
  }
  return applyCellUpdates(next, sheetId, updates);
}

export function mergeSelection(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  const normalized = normalizeRange(range);
  if (normalized.r1 === normalized.r2 && normalized.c1 === normalized.c2) return workbook;
  return updateSheet(workbook, sheetId, (sheet) => ({
    ...sheet,
    merges: [
      ...sheet.merges.filter(
        (merge) =>
          merge.r2 < normalized.r1 ||
          merge.r1 > normalized.r2 ||
          merge.c2 < normalized.c1 ||
          merge.c1 > normalized.c2,
      ),
      normalized,
    ],
  }));
}

export function unmergeSelection(workbook: WorkbookModel, sheetId: string, range: CellRange): WorkbookModel {
  const normalized = normalizeRange(range);
  return updateSheet(workbook, sheetId, (sheet) => ({
    ...sheet,
    merges: sheet.merges.filter(
      (merge) =>
        merge.r2 < normalized.r1 ||
        merge.r1 > normalized.r2 ||
        merge.c2 < normalized.c1 ||
        merge.c1 > normalized.c2,
    ),
  }));
}

export function usedRange(sheet: SheetModel): CellRange {
  let maxRow = 0;
  let maxCol = 0;
  let found = false;
  for (const key of Object.keys(sheet.cells)) {
    const [row, col] = key.split(':').map(Number);
    maxRow = Math.max(maxRow, row);
    maxCol = Math.max(maxCol, col);
    found = true;
  }
  for (const merge of sheet.merges) {
    maxRow = Math.max(maxRow, merge.r2);
    maxCol = Math.max(maxCol, merge.c2);
    found = true;
  }
  if (sheet.autoFilter) {
    maxRow = Math.max(maxRow, sheet.autoFilter.r2);
    maxCol = Math.max(maxCol, sheet.autoFilter.c2);
    found = true;
  }
  return found ? { r1: 0, c1: 0, r2: maxRow, c2: maxCol } : { r1: 0, c1: 0, r2: 0, c2: 0 };
}

export function insertRows(workbook: WorkbookModel, sheetId: string, at: number, count = 1): WorkbookModel {
  // 功能：先维护全工作簿引用，再移动单元格；时间：2026-10-18；作者：lq。
  if (!Number.isInteger(at) || at < 0 || !Number.isInteger(count) || count < 1) return workbook;
  const target = getSheet(workbook, sheetId);
  if (!target) return workbook;
  const adjusted = editWorkbookReferences(workbook, { kind: 'rows', sheet: target.name, at, count, deleting: false });
  return updateSheet(adjusted, sheetId, (sheet) => {
    const cells: SheetModel['cells'] = {};
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const [row, col] = key.split(':').map(Number);
      cells[cellKey(row >= at ? row + count : row, col)] = cell;
    }
    const rowHeights: SheetModel['rowHeights'] = {};
    for (const [row, height] of Object.entries(sheet.rowHeights)) {
      const index = Number(row);
      rowHeights[index >= at ? index + count : index] = height;
    }
    const hiddenRows: SheetModel['hiddenRows'] = {};
    for (const row of Object.keys(sheet.hiddenRows)) {
      const index = Number(row);
      hiddenRows[index >= at ? index + count : index] = true;
    }
    return {
      ...sheet,
      cells,
      rowHeights,
      hiddenRows,
      merges: sheet.merges.map((merge) =>
        merge.r1 >= at
          ? { ...merge, r1: merge.r1 + count, r2: merge.r2 + count }
          : merge.r2 >= at
            ? { ...merge, r2: merge.r2 + count }
            : merge,
      ),
      freeze:
        sheet.freeze && sheet.freeze.rows >= at
          ? { ...sheet.freeze, rows: sheet.freeze.rows + count }
          : sheet.freeze,
      autoFilter:
        sheet.autoFilter && sheet.autoFilter.r1 >= at
          ? { ...sheet.autoFilter, r1: sheet.autoFilter.r1 + count, r2: sheet.autoFilter.r2 + count }
          : sheet.autoFilter,
      dimension: { ...sheet.dimension, rows: sheet.dimension.rows + count },
    };
  });
}

export function deleteRows(workbook: WorkbookModel, sheetId: string, at: number, count = 1): WorkbookModel {
  const end = at + count - 1;
  // 功能：先维护全工作簿引用，再移动单元格；时间：2026-10-18；作者：lq。
  if (!Number.isInteger(at) || at < 0 || !Number.isInteger(count) || count < 1) return workbook;
  const target = getSheet(workbook, sheetId);
  if (!target) return workbook;
  const adjusted = editWorkbookReferences(workbook, { kind: 'rows', sheet: target.name, at, count, deleting: true });
  return updateSheet(adjusted, sheetId, (sheet) => {
    const cells: SheetModel['cells'] = {};
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const [row, col] = key.split(':').map(Number);
      if (row >= at && row <= end) continue;
      cells[cellKey(row > end ? row - count : row, col)] = cell;
    }
    const rowHeights: SheetModel['rowHeights'] = {};
    for (const [row, height] of Object.entries(sheet.rowHeights)) {
      const index = Number(row);
      if (index >= at && index <= end) continue;
      rowHeights[index > end ? index - count : index] = height;
    }
    const hiddenRows: SheetModel['hiddenRows'] = {};
    for (const row of Object.keys(sheet.hiddenRows)) {
      const index = Number(row);
      if (index >= at && index <= end) continue;
      hiddenRows[index > end ? index - count : index] = true;
    }
    return {
      ...sheet,
      cells,
      rowHeights,
      hiddenRows,
      merges: sheet.merges
        .map((merge) =>
          merge.r1 > end
            ? { ...merge, r1: merge.r1 - count, r2: merge.r2 - count }
            : merge.r2 > end
              ? { ...merge, r2: merge.r2 - count }
              : merge.r1 >= at && merge.r2 <= end
                ? null
                : merge,
        )
        .filter((merge): merge is CellRange => !!merge),
      autoFilter:
        sheet.autoFilter && sheet.autoFilter.r1 >= at
          ? {
              ...sheet.autoFilter,
              r1: Math.max(0, sheet.autoFilter.r1 - count),
              r2: Math.max(0, sheet.autoFilter.r2 - count),
            }
          : sheet.autoFilter,
      dimension: { ...sheet.dimension, rows: Math.max(100, sheet.dimension.rows - count) },
    };
  });
}

export function insertColumns(workbook: WorkbookModel, sheetId: string, at: number, count = 1): WorkbookModel {
  // 功能：先维护全工作簿引用，再移动单元格；时间：2026-10-18；作者：lq。
  if (!Number.isInteger(at) || at < 0 || !Number.isInteger(count) || count < 1) return workbook;
  const target = getSheet(workbook, sheetId);
  if (!target) return workbook;
  const adjusted = editWorkbookReferences(workbook, { kind: 'columns', sheet: target.name, at, count, deleting: false });
  return updateSheet(adjusted, sheetId, (sheet) => {
    const cells: SheetModel['cells'] = {};
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const [row, col] = key.split(':').map(Number);
      cells[cellKey(row, col >= at ? col + count : col)] = cell;
    }
    const colWidths: SheetModel['colWidths'] = {};
    for (const [col, width] of Object.entries(sheet.colWidths)) {
      const index = Number(col);
      colWidths[index >= at ? index + count : index] = width;
    }
    const hiddenCols: SheetModel['hiddenCols'] = {};
    for (const col of Object.keys(sheet.hiddenCols)) {
      const index = Number(col);
      hiddenCols[index >= at ? index + count : index] = true;
    }
    return {
      ...sheet,
      cells,
      colWidths,
      hiddenCols,
      merges: sheet.merges.map((merge) =>
        merge.c1 >= at
          ? { ...merge, c1: merge.c1 + count, c2: merge.c2 + count }
          : merge.c2 >= at
            ? { ...merge, c2: merge.c2 + count }
            : merge,
      ),
      freeze:
        sheet.freeze && sheet.freeze.cols >= at
          ? { ...sheet.freeze, cols: sheet.freeze.cols + count }
          : sheet.freeze,
      dimension: { ...sheet.dimension, cols: sheet.dimension.cols + count },
    };
  });
}

export function deleteColumns(workbook: WorkbookModel, sheetId: string, at: number, count = 1): WorkbookModel {
  const end = at + count - 1;
  // 功能：先维护全工作簿引用，再移动单元格；时间：2026-10-18；作者：lq。
  if (!Number.isInteger(at) || at < 0 || !Number.isInteger(count) || count < 1) return workbook;
  const target = getSheet(workbook, sheetId);
  if (!target) return workbook;
  const adjusted = editWorkbookReferences(workbook, { kind: 'columns', sheet: target.name, at, count, deleting: true });
  return updateSheet(adjusted, sheetId, (sheet) => {
    const cells: SheetModel['cells'] = {};
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const [row, col] = key.split(':').map(Number);
      if (col >= at && col <= end) continue;
      cells[cellKey(row, col > end ? col - count : col)] = cell;
    }
    const colWidths: SheetModel['colWidths'] = {};
    for (const [col, width] of Object.entries(sheet.colWidths)) {
      const index = Number(col);
      if (index >= at && index <= end) continue;
      colWidths[index > end ? index - count : index] = width;
    }
    const hiddenCols: SheetModel['hiddenCols'] = {};
    for (const col of Object.keys(sheet.hiddenCols)) {
      const index = Number(col);
      if (index >= at && index <= end) continue;
      hiddenCols[index > end ? index - count : index] = true;
    }
    return {
      ...sheet,
      cells,
      colWidths,
      hiddenCols,
      merges: sheet.merges
        .map((merge) =>
          merge.c1 > end
            ? { ...merge, c1: merge.c1 - count, c2: merge.c2 - count }
            : merge.c2 > end
              ? { ...merge, c2: merge.c2 - count }
              : merge.c1 >= at && merge.c2 <= end
                ? null
                : merge,
        )
        .filter((merge): merge is CellRange => !!merge),
      dimension: { ...sheet.dimension, cols: Math.max(26, sheet.dimension.cols - count) },
    };
  });
}

export function setFreeze(workbook: WorkbookModel, sheetId: string, rows: number, cols: number): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => ({
    ...sheet,
    freeze: rows > 0 || cols > 0 ? { rows, cols } : undefined,
  }));
}

export function setColumnWidth(workbook: WorkbookModel, sheetId: string, col: number, width: number): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => ({
    ...sheet,
    colWidths: { ...sheet.colWidths, [col]: Math.max(24, Math.min(800, width)) },
  }));
}

export function setRowHeight(workbook: WorkbookModel, sheetId: string, row: number, height: number): WorkbookModel {
  return updateSheet(workbook, sheetId, (sheet) => ({
    ...sheet,
    rowHeights: { ...sheet.rowHeights, [row]: Math.max(14, Math.min(400, height)) },
  }));
}

export interface SheetStats {
  count: number;
  numericCount: number;
  sum: number;
  average: number;
  min: number;
  max: number;
}

export function selectionStats(workbook: WorkbookModel, sheet: SheetModel, range: CellRange): SheetStats {
  const normalized = normalizeRange(range);
  let count = 0;
  let numericCount = 0;
  let sum = 0;
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;
  for (let row = normalized.r1; row <= normalized.r2; row += 1) {
    if (sheet.hiddenRows[row]) continue;
    for (let col = normalized.c1; col <= normalized.c2; col += 1) {
      if (sheet.hiddenCols[col]) continue;
      const cell = sheet.cells[cellKey(row, col)];
      const computed = workbook.computed?.[sheet.id]?.[cellKey(row, col)];
      const value = computed?.value ?? cell?.value;
      if (value === null || value === undefined || value === '') continue;
      count += 1;
      if (typeof value === 'number') {
        numericCount += 1;
        sum += value;
        min = Math.min(min, value);
        max = Math.max(max, value);
      }
    }
  }
  return {
    count,
    numericCount,
    sum,
    average: numericCount > 0 ? sum / numericCount : 0,
    min: numericCount > 0 ? min : 0,
    max: numericCount > 0 ? max : 0,
  };
}

export interface ClipboardCell {
  row: number;
  col: number;
  cell: Cell;
  style: CellStyle;
}

export interface SheetClipboard {
  sourceSheetId: string;
  range: CellRange;
  cells: ClipboardCell[];
}

export function extractClipboard(workbook: WorkbookModel, sheet: SheetModel, range: CellRange): SheetClipboard {
  const normalized = normalizeRange(range);
  const cells: ClipboardCell[] = [];
  for (let row = normalized.r1; row <= normalized.r2; row += 1) {
    for (let col = normalized.c1; col <= normalized.c2; col += 1) {
      const cell = sheet.cells[cellKey(row, col)];
      if (!cell) continue;
      cells.push({
        row,
        col,
        cell: { ...cell },
        style: { ...cellStyle(workbook, cell) },
      });
    }
  }
  return { sourceSheetId: sheet.id, range: normalized, cells };
}

export function clipboardToTsv(workbook: WorkbookModel, sheet: SheetModel, clipboard: SheetClipboard): string {
  const { range } = clipboard;
  const rows: string[] = [];
  for (let row = range.r1; row <= range.r2; row += 1) {
    const fields: string[] = [];
    for (let col = range.c1; col <= range.c2; col += 1) {
      fields.push(cellDisplay(workbook, sheet, row, col).replace(/[\t\r\n]/g, ' '));
    }
    rows.push(fields.join('\t'));
  }
  return rows.join('\n');
}

export function parseTsv(text: string): string[][] {
  return text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\n$/, '')
    .split('\n')
    .map((line) => line.split('\t'));
}

export function pasteClipboard(
  workbook: WorkbookModel,
  sheet: SheetModel,
  target: CellRange,
  clipboard: SheetClipboard,
): WorkbookModel {
  const normalized = normalizeRange(clipboard.range);
  const width = normalized.c2 - normalized.c1 + 1;
  const height = normalized.r2 - normalized.r1 + 1;
  const targetRange = normalizeRange(target);
  let next = workbook;
  const updates: CellUpdate[] = [];
  const styleCache = new Map<string, string>();
  for (let row = targetRange.r1; row <= targetRange.r2; row += 1) {
    for (let col = targetRange.c1; col <= targetRange.c2; col += 1) {
      const sourceRow = normalized.r1 + ((row - targetRange.r1) % height);
      const sourceCol = normalized.c1 + ((col - targetRange.c1) % width);
      const source = clipboard.cells.find((item) => item.row === sourceRow && item.col === sourceCol);
      if (!source) {
        updates.push({ row, col, cell: null });
        continue;
      }
      const style = source.style;
      let styleId = styleCache.get(styleKey(style));
      if (!styleId) {
        const found = findOrCreateStyle(next, style);
        next = found.workbook;
        styleId = found.styleId;
        styleCache.set(styleKey(style), styleId);
      }
      updates.push({
        row,
        col,
        cell: {
          value: source.cell.value,
          formula: source.cell.formula
            ? translateFormula(source.cell.formula, row - source.row, col - source.col)
            : undefined,
          styleId,
        },
      });
    }
  }
  return applyCellUpdates(next, sheet.id, updates);
}

export function pasteTsv(
  workbook: WorkbookModel,
  sheet: SheetModel,
  target: CellCoord,
  rows: string[][],
): WorkbookModel {
  let next = workbook;
  const updates: CellUpdate[] = [];
  rows.forEach((rowValues, rowOffset) => {
    rowValues.forEach((value, colOffset) => {
      const row = target.row + rowOffset;
      const col = target.col + colOffset;
      const existing = sheet.cells[cellKey(row, col)];
      updates.push({ row, col, cell: createCellFromInput(value, next, existing) });
    });
  });
  return applyCellUpdates(next, sheet.id, updates);
}

/**
 * 功能：拖动填充柄——单列/单行数字按等差数列延展，其他内容按原样重复填充；
 * 时间：2026-11-01；作者：lq。
 *
 * 与 Excel 一致：单格数字向下拖得到 1、2、3…，多格数字按步长延展，文本与公式按周期重复。
 */
export function fillRange(
  workbook: WorkbookModel,
  sheetId: string,
  source: CellRange,
  target: CellRange,
): WorkbookModel {
  const sheet = getSheet(workbook, sheetId);
  if (!sheet) return workbook;
  const src = normalizeRange(source);
  const dst = normalizeRange(target);
  if (src.r1 === dst.r1 && src.r2 === dst.r2 && src.c1 === dst.c1 && src.c2 === dst.c2) return workbook;
  const clipboard = extractClipboard(workbook, sheet, src);
  let next = pasteClipboard(workbook, sheet, dst, clipboard);

  const vertical = dst.c1 === src.c1 && dst.c2 === src.c2 && (dst.r2 > src.r2 || dst.r1 < src.r1);
  const horizontal = dst.r1 === src.r1 && dst.r2 === src.r2 && (dst.c2 > src.c2 || dst.c1 < src.c1);

  /** 功能：按等差数列推算目标位置的值，源序列含非数字时返回 null 走重复填充；时间：2026-11-01；作者：lq。 */
  const seriesValue = (values: Array<number | null>, offset: number): number | null => {
    if (values.length === 0 || values.some((value) => value === null)) return null;
    const numbers = values as number[];
    if (numbers.length === 1) return numbers[0] + offset;
    const step = (numbers[numbers.length - 1] - numbers[0]) / (numbers.length - 1);
    const base = offset > 0 ? numbers[numbers.length - 1] : numbers[0];
    return base + step * offset;
  };

  const updates: CellUpdate[] = [];
  if (vertical) {
    for (let col = src.c1; col <= src.c2; col += 1) {
      const values: Array<number | null> = [];
      for (let row = src.r1; row <= src.r2; row += 1) {
        const cell = sheet.cells[cellKey(row, col)];
        values.push(typeof cell?.value === 'number' && !cell.formula ? cell.value : null);
      }
      for (let row = dst.r1; row <= dst.r2; row += 1) {
        if (row >= src.r1 && row <= src.r2) continue;
        const offset = row > src.r2 ? row - src.r2 : row - src.r1;
        const value = seriesValue(values, offset);
        if (value === null) continue;
        updates.push({ row, col, cell: { value, styleId: sheet.cells[cellKey(row, col)]?.styleId } });
      }
    }
  } else if (horizontal) {
    for (let row = src.r1; row <= src.r2; row += 1) {
      const values: Array<number | null> = [];
      for (let col = src.c1; col <= src.c2; col += 1) {
        const cell = sheet.cells[cellKey(row, col)];
        values.push(typeof cell?.value === 'number' && !cell.formula ? cell.value : null);
      }
      for (let col = dst.c1; col <= dst.c2; col += 1) {
        if (col >= src.c1 && col <= src.c2) continue;
        const offset = col > src.c2 ? col - src.c2 : col - src.c1;
        const value = seriesValue(values, offset);
        if (value === null) continue;
        updates.push({ row, col, cell: { value, styleId: sheet.cells[cellKey(row, col)]?.styleId } });
      }
    }
  }
  return updates.length > 0 ? applyCellUpdates(next, sheet.id, updates) : next;
}

// 功能：表头选择使用实际数据宽度，跳过合并标题，避免将空白整行整列当成数据范围；时间：2026-10-24；作者：lq。
export function tableRangeForSelection(sheet: SheetModel, selection: Selection): CellRange {
  const selected = selectionRange(selection);
  const used = usedRange(sheet);
  if (selection.kind === 'row') return { ...used, r1: Math.min(used.r2, Math.max(used.r1, selected.r1)) };
  if (!selection.kind && (selected.r1 !== selected.r2 || selected.c1 !== selected.c2)) return selected;
  const target = { ...used };
  const title = sheet.merges.find((merge) => merge.r1 === target.r1 && merge.c1 <= target.c1 && merge.c2 >= target.c2);
  if (title && title.r2 < target.r2) target.r1 = title.r2 + 1;
  return target;
}

export function sortRange(
  workbook: WorkbookModel,
  sheet: SheetModel,
  range: CellRange,
  sortColumn: number,
  ascending: boolean,
): WorkbookModel {
  const normalized = normalizeRange(range);
  // 功能：记录原始行号，排序后同步平移相对公式；时间：2026-10-18；作者：lq。
  const rows: Array<{ row: number; cells: Array<{ col: number; cell: Cell | undefined }>; key: string | number }> = [];
  for (let row = normalized.r1; row <= normalized.r2; row += 1) {
    const entry = { row, cells: [] as Array<{ col: number; cell: Cell | undefined }>, key: 0 as string | number };
    for (let col = normalized.c1; col <= normalized.c2; col += 1) {
      const cell = sheet.cells[cellKey(row, col)];
      entry.cells.push({ col, cell });
      if (col === sortColumn) {
        const computed = workbook.computed?.[sheet.id]?.[cellKey(row, col)];
        const value = computed?.value ?? cell?.value ?? '';
        entry.key = typeof value === 'number' ? value : String(value);
      }
    }
    rows.push(entry);
  }
  rows.sort((a, b) => {
    // 功能：升序和降序均将空白放在数据末尾，保持整行关联；时间：2026-10-24；作者：lq。
    if (a.key === '' || b.key === '') return a.key === b.key ? 0 : a.key === '' ? 1 : -1;
    if (typeof a.key === 'number' && typeof b.key === 'number') {
      return ascending ? a.key - b.key : b.key - a.key;
    }
    const compare = String(a.key).localeCompare(String(b.key), 'zh-CN', { numeric: true });
    return ascending ? compare : -compare;
  });
  const updates: CellUpdate[] = [];
  rows.forEach((entry, rowOffset) => {
    entry.cells.forEach(({ col, cell }) => {
      const row = normalized.r1 + rowOffset;
      updates.push({ row, col, cell: cell?.formula
        ? { ...cell, formula: moveSortedFormula(cell.formula, row - entry.row) }
        : cell ?? null });
    });
  });
  return applyCellUpdates(workbook, sheet.id, updates);
}

// 功能：统一 Excel 常用边框预设，工具栏和格式弹窗共用；时间：2026-10-24；作者：lq。
export type BorderPreset = 'all' | 'outline' | 'bottom' | 'top' | 'left' | 'right' | 'none' | 'inside' | 'thickOutline' | 'doubleBottom' | 'thickBottom' | 'topBottom' | 'topThickBottom' | 'topDoubleBottom';
export function applyBordersToRange(workbook: WorkbookModel, sheetId: string, target: CellRange, preset: BorderPreset, side: NonNullable<CellBorders['top']>): WorkbookModel {
  const range = normalizeRange(target);
  if (preset === 'none') return applyStyleToRange(workbook, sheetId, range, { borders: undefined });
  if (preset === 'all') return applyStyleToRange(workbook, sheetId, range, { borders: { top: side, right: side, bottom: side, left: side } });
  const edges = (next: WorkbookModel, edge: keyof CellBorders, border = side): WorkbookModel => {
    const boundary = edge === 'top' ? { ...range, r2: range.r1 } : edge === 'bottom' ? { ...range, r1: range.r2 } : edge === 'left' ? { ...range, c2: range.c1 } : { ...range, c1: range.c2 };
    return applyStyleToRange(next, sheetId, boundary, { borders: { [edge]: border } });
  };
  if (['top', 'bottom', 'left', 'right'].includes(preset)) return edges(workbook, preset as keyof CellBorders);
  if (preset === 'doubleBottom' || preset === 'thickBottom') return edges(workbook, 'bottom', { ...side, style: preset === 'doubleBottom' ? 'double' : 'thick' });
  let next = workbook;
  if (preset === 'inside') {
    if (range.r2 > range.r1) {
      next = applyStyleToRange(next, sheetId, { ...range, r2: range.r2 - 1 }, { borders: { bottom: side } });
      next = applyStyleToRange(next, sheetId, { ...range, r1: range.r1 + 1 }, { borders: { top: side } });
    }
    if (range.c2 > range.c1) {
      next = applyStyleToRange(next, sheetId, { ...range, c2: range.c2 - 1 }, { borders: { right: side } });
      next = applyStyleToRange(next, sheetId, { ...range, c1: range.c1 + 1 }, { borders: { left: side } });
    }
    return next;
  }
  const outlineSide = preset === 'thickOutline' ? { ...side, style: 'thick' as const } : side;
  next = edges(next, 'top', outlineSide);
  const bottom = preset === 'topDoubleBottom' ? { ...side, style: 'double' as const } : preset === 'topThickBottom' ? { ...side, style: 'thick' as const } : outlineSide;
  next = edges(next, 'bottom', bottom);
  if (preset === 'outline' || preset === 'thickOutline') { next = edges(next, 'left', outlineSide); next = edges(next, 'right', outlineSide); }
  return next;
}

// 功能：一次提交单元格格式及选区边框，外框仅写入边缘，支持整体撤销；时间：2026-10-24；作者：lq。
export function applyCellFormat(workbook: WorkbookModel, sheetId: string, range: CellRange, patch: Partial<CellStyle>, border?: { preset: BorderPreset; side: NonNullable<CellBorders['top']> }): WorkbookModel {
  const next = applyStyleToRange(workbook, sheetId, range, patch);
  return border ? applyBordersToRange(next, sheetId, range, border.preset, border.side) : next;
}


export function applyFilter(
  workbook: WorkbookModel,
  sheet: SheetModel,
  column: number,
  allowed: Set<string> | null,
  filterRange: CellRange,
): WorkbookModel {
  const normalized = normalizeRange(filterRange);
  // 功能：每次按全部列条件重新计算可见行，并保留用户手动隐藏的行；时间：2026-10-24；作者：lq。
  const filterCriteria = { ...sheet.filterCriteria };
  if (allowed === null) delete filterCriteria[column];
  else filterCriteria[column] = [...allowed];
  const filterBaseHiddenRows = sheet.filterBaseHiddenRows ?? { ...sheet.hiddenRows };
  const hiddenRows: SheetModel['hiddenRows'] = { ...filterBaseHiddenRows };
  const conditions = Object.entries(filterCriteria).map(([col, values]) => [Number(col), new Set(values)] as const);
  for (let row = normalized.r1; row <= normalized.r2; row += 1) {
    if (conditions.some(([col, values]) => !values.has(cellDisplay(workbook, sheet, row, col)))) hiddenRows[row] = true;
  }
  return updateSheet(workbook, sheet.id, (current) => ({ ...current, hiddenRows, filterCriteria, filterBaseHiddenRows }));
}

export function uniqueColumnValues(
  workbook: WorkbookModel,
  sheet: SheetModel,
  column: number,
  range: CellRange,
): string[] {
  const values = new Set<string>();
  for (let row = range.r1; row <= range.r2; row += 1) {
    values.add(cellDisplay(workbook, sheet, row, column));
  }
  return [...values].sort((a, b) => a.localeCompare(b, 'zh-CN', { numeric: true }));
}

export function addSheet(workbook: WorkbookModel, name?: string): WorkbookModel {
  const sheet: SheetModel = {
    id: `sheet-${uid()}`,
    name: name ?? nextSheetName(workbook),
    state: 'visible',
    order: workbook.sheets.length,
    cells: {},
    colWidths: {},
    rowHeights: {},
    merges: [],
    hiddenRows: {},
    hiddenCols: {},
    dimension: { rows: 100, cols: 26 },
    defaultColWidth: 64,
    defaultRowHeight: 21,
    showGridLines: true,
  };
  return {
    ...workbook,
    sheets: [...workbook.sheets, sheet],
    activeSheetId: sheet.id,
  };
}

export function nextSheetName(workbook: WorkbookModel): string {
  const used = new Set(workbook.sheets.map((sheet) => sheet.name));
  let index = workbook.sheets.length + 1;
  while (used.has(`Sheet${index}`)) index += 1;
  return `Sheet${index}`;
}

export function renameSheet(workbook: WorkbookModel, sheetId: string, name: string): WorkbookModel {
  // 功能：验证名称唯一性并维护跨表公式；时间：2026-10-18；作者：lq。
  const target = getSheet(workbook, sheetId);
  const nextName = name.trim();
  if (!target || nextName === target.name) return workbook;
  if (!nextName || nextName.length > 31 || /[\[\]:*?/\\]/.test(nextName) || /^'|'$/.test(nextName)) {
    throw new Error('工作表名应为 1–31 个字符，不能包含 []:*?/\\ 或以单引号开头、结尾。');
  }
  if (workbook.sheets.some((sheet) => sheet.id !== sheetId && sheet.name.toLowerCase() === nextName.toLowerCase())) {
    throw new Error('工作表名称已存在。');
  }
  const adjusted = editWorkbookReferences(workbook, { kind: 'rename', sheet: target.name, name: nextName });
  return updateSheet(adjusted, sheetId, (sheet) => ({ ...sheet, name: nextName }));
}

export function deleteSheet(workbook: WorkbookModel, sheetId: string): WorkbookModel {
  if (workbook.sheets.length <= 1) return workbook;
  // 功能：删除工作表后将相关跨表引用标记为 REF；时间：2026-10-18；作者：lq。
  const target = getSheet(workbook, sheetId);
  if (!target) return workbook;
  const adjusted = editWorkbookReferences(workbook, { kind: 'delete-sheet', sheet: target.name });
  const sheets = adjusted.sheets
    .filter((sheet) => sheet.id !== sheetId)
    .map((sheet, index) => ({ ...sheet, order: index }));
  return {
    ...adjusted,
    sheets,
    activeSheetId: workbook.activeSheetId === sheetId ? sheets[0].id : workbook.activeSheetId,
  };
}
