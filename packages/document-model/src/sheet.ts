export type CellScalar = string | number | boolean | null;

export type HorizontalAlign = 'general' | 'left' | 'center' | 'right';
export type VerticalAlign = 'top' | 'middle' | 'bottom';

export interface BorderSide {
  style: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'none';
  color?: string;
}

export interface CellBorders {
  top?: BorderSide;
  right?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
}

export interface CellStyle {
  fontFamily?: string;
  fontSizePt?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  fill?: string;
  horizontalAlign?: HorizontalAlign;
  verticalAlign?: VerticalAlign;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  numFmt?: string;
  numFmtId?: number;
  borders?: CellBorders;
  /** Original cellXfs index, used to preserve untouched style references. */
  xfIndex?: number;
}

export interface Cell {
  value: CellScalar;
  formula?: string;
  styleId?: string;
}

export interface ComputedCell {
  value: CellScalar;
  display: string;
  error?: string;
}

export interface CellRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface FreezePane {
  rows: number;
  cols: number;
}

/**
 * 功能：保存 Excel 页面布局、打印缩放、页边距和打印辅助项；时间：2026-10-24；作者：lq。
 */
export interface SheetPageSetup {
  orientation: 'portrait' | 'landscape';
  scalePercent: number;
  fitToWidth?: number;
  fitToHeight?: number;
  firstPageNumber?: number;
  margins: { left: number; right: number; top: number; bottom: number; header: number; footer: number };
  printGridLines: boolean;
  printHeadings: boolean;
  /**
   * 功能：打印标题（顶端标题行），形如 `1:2`，每页都会重复输出这些行；空表示沿用“首行当表头”；
   * 时间：2026-11-01；作者：lq。
   */
  printTitleRows?: string;
}

/** 功能：返回与 Excel 新工作表一致的页面设置默认值；时间：2026-10-24；作者：lq。 */
export function defaultSheetPageSetup(): SheetPageSetup {
  return {
    orientation: 'landscape',
    scalePercent: 100,
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    printGridLines: true,
    printHeadings: false,
  };
}

// 功能：保存工作表浮动图片的内容、单元格锚点及像素尺寸；时间：2026-10-24；作者：lq。
export interface SheetImage {
  id: string;
  name: string;
  src: string;
  row: number;
  col: number;
  offsetX: number;
  offsetY: number;
  width: number;
  height: number;
}

export interface SheetModel {
  id: string;
  name: string;
  state: 'visible' | 'hidden' | 'veryHidden';
  order: number;
  cells: Record<string, Cell>;
  colWidths: Record<number, number>;
  rowHeights: Record<number, number>;
  merges: CellRange[];
  freeze?: FreezePane;
  hiddenRows: Record<number, true>;
  hiddenCols: Record<number, true>;
  autoFilter?: CellRange;
  // 功能：保存各列筛选条件与筛选前隐藏行，支持组合筛选、撤销和恢复；时间：2026-10-24；作者：lq。
  filterCriteria?: Record<number, string[]>;
  filterBaseHiddenRows?: Record<number, true>;
  dimension: { rows: number; cols: number };
  defaultColWidth: number;
  defaultRowHeight: number;
  showGridLines: boolean;
  tabColor?: string;
  // 功能：图片随工作簿快照保存，支持撤销与恢复；时间：2026-10-24；作者：lq。
  images?: SheetImage[];
  pageSetup?: SheetPageSetup;
}

export interface WorkbookModel {
  version: 1;
  sheets: SheetModel[];
  activeSheetId: string;
  styles: Record<string, CellStyle>;
  themeColors: Record<string, string>;
  themeFonts: { major: string; minor: string };
  defaultFont: string;
  defaultFontSizePt: number;
  date1904?: boolean;
  definedNames?: Record<string, string>;
  csv?: boolean;
  meta: { title?: string; author?: string; created?: string; modified?: string };
  /** Transient calculation cache; never serialized to XLSX. */
  computed?: Record<string, Record<string, ComputedCell>>;
  stats?: { cells: number; sheets: number };
}

export const DEFAULT_COL_WIDTH = 88;
export const DEFAULT_ROW_HEIGHT = 24;

export function createSheet(name: string, order: number): SheetModel {
  return {
    id: `sheet-${order + 1}`,
    name,
    state: 'visible',
    order,
    cells: {},
    colWidths: {},
    rowHeights: {},
    merges: [],
    hiddenRows: {},
    hiddenCols: {},
    dimension: { rows: 100, cols: 26 },
    defaultColWidth: DEFAULT_COL_WIDTH,
    defaultRowHeight: DEFAULT_ROW_HEIGHT,
    showGridLines: true,
    pageSetup: defaultSheetPageSetup(),
  };
}

export function createBlankWorkbook(): WorkbookModel {
  const sheet = createSheet('Sheet1', 0);
  return {
    version: 1,
    sheets: [sheet],
    activeSheetId: sheet.id,
    styles: {
      default: {
        fontFamily: '等线',
        fontSizePt: 11,
        numFmt: 'General',
        horizontalAlign: 'general',
        verticalAlign: 'bottom',
      },
    },
    themeColors: {},
    themeFonts: { major: '等线 Light', minor: '等线' },
    defaultFont: '等线',
    defaultFontSizePt: 11,
    meta: { title: '未命名工作簿' },
    stats: { cells: 0, sheets: 1 },
  };
}

export function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

export function parseCellKey(key: string): { row: number; col: number } {
  const [row, col] = key.split(':').map(Number);
  return { row, col };
}

export function rangeContains(range: CellRange, row: number, col: number): boolean {
  return row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2;
}

export function normalizeRange(range: CellRange): CellRange {
  return {
    r1: Math.min(range.r1, range.r2),
    r2: Math.max(range.r1, range.r2),
    c1: Math.min(range.c1, range.c2),
    c2: Math.max(range.c1, range.c2),
  };
}
