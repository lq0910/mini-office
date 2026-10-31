import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type MouseEvent as ReactMouseEvent,
} from 'react';
import {
  cellKey,
  defaultSheetPageSetup,
  type CellRange,
  type CellStyle,
  type SheetModel,
  type SheetImage,
  type SheetSession,
  type SheetPageSetup,
  type WorkbookModel,
} from '@mini-office/document-model';
import { useWorkspaceStore } from '@mini-office/document-model';
import { printPages, renderPrintPreview } from '@mini-office/platform-adapter';
import {
  Button,
  ContextMenu,
  Dialog,
  Field,
  IconButton,
  Input,
  ProgressBar,
  SegmentedControl,
  Select,
  Switch,
  type MenuItem,
} from '@mini-office/ui';
import {
  ChevronDown,
  ChevronUp,
  FileWarning,
  Search,
  X,
} from 'lucide-react';
import { exportElementsToPdf, type ExportProgress } from '@mini-office/export-service';
import {
  columnLabel,
  ensureExtension,
  formatBytes,
  stripExtension,
  uid,
} from '@mini-office/shared';
// 功能：接入表头筛选浮层；时间：2026-10-24；作者：lq。
import { ColumnFilterMenu } from './ColumnFilterMenu';
// 功能：接入单元格格式和原位工作表标签操作；时间：2026-10-24；作者：lq。
import { completeFormulaParentheses } from './formula/reference-picking';
import { CellFormatDialog } from './CellFormatDialog';
import { SheetTabs } from './SheetTabs';
// 功能：原生文件选择失败时提供可见的选择图片入口；时间：2026-10-24；作者：lq。
import { openSheetImagePicker } from './image-picker';
import { SheetGrid, type EditingCell, type GridAction } from './SheetGrid';
import { SheetToolbar, type BorderPreset, type BorderWeight } from './SheetToolbar';
// 功能：接入 Excel 原版缩放快捷键；时间：2026-10-24；作者：lq。
import { normalizeSheetZoom, SHEET_ZOOM_MAX, SHEET_ZOOM_MIN, useSheetZoomShortcuts } from './zoom-shortcuts';
import { recalculateWorkbook } from './formula/engine';
// 功能：打印分页独立成模块，按真实行列尺寸切页；时间：2026-10-29；作者：lq。
import { buildSheetPrintPages, sheetPageSize } from './print';
import {
  activeSheet as activeSheetOf,
  addSheet,
  applyCellUpdates,
  applyFilter,
  applyStyleToRange,
  applyCellFormat,
  applyBordersToRange,
  cellAt,
  cellDisplay,
  cellRawInput,
  cellStyle,
  clearRangeContents,
  deleteColumns,
  deleteRows,
  deleteSheet,
  extractClipboard,
  findOrCreateStyle,
  insertColumns,
  insertRows,
  mergeSelection,
  parseRangeRef,
  pasteClipboard,
  pasteTsv,
  parseTsv,
  rangeArea,
  rangeRef,
  renameSheet,
  selectionRange,
  selectionStats,
  setCellInput,
  setColumnWidth,
  setFreeze,
  setRowHeight,
  sortRange,
  tableRangeForSelection,
  unmergeSelection,
  uniqueColumnValues,
  updateSheet,
  usedRange,
  type Selection,
  type SheetClipboard,
} from './sheet-utils';

interface HistoryEntry {
  past: WorkbookModel[];
  future: WorkbookModel[];
  lastKey?: string;
  lastTime: number;
}

const historyStore = new Map<string, HistoryEntry>();

function historyFor(id: string): HistoryEntry {
  const existing = historyStore.get(id);
  if (existing) return existing;
  const created: HistoryEntry = { past: [], future: [], lastTime: 0 };
  historyStore.set(id, created);
  return created;
}

interface FilterState {
  column: number;
  range: CellRange;
  autoFilterRange: CellRange;
  values: string[];
  selected: Set<string>;
  anchor: { left: number; bottom: number };
}

interface ExportState {
  open: boolean;
  name: string;
  progress: ExportProgress | null;
  busy: boolean;
}

export function SheetEditor({ session }: { session: SheetSession }) {
  const workbook = session.data as WorkbookModel;
  const readonly = session.readonly;
  const updateData = useWorkspaceStore((state) => state.updateData);
  const updateView = useWorkspaceStore((state) => state.updateView);
  const settings = useWorkspaceStore((state) => state.settings);
  const updateSettings = useWorkspaceStore((state) => state.updateSettings);
  const exportBytes = useWorkspaceStore((state) => state.exportBytes);
  const openCompatibility = useWorkspaceStore((state) => state.openDialog);
  const notify = useWorkspaceStore((state) => state.notify);

  const [sheetId, setSheetId] = useState(() => {
    const saved = session.view.activeSheetId;
    return saved && workbook.sheets.some((sheet) => sheet.id === saved)
      ? saved
      : workbook.activeSheetId;
  });
  const sheet = workbook.sheets.find((item) => item.id === sheetId) ?? workbook.sheets[0];
  const [selection, setSelection] = useState<Selection>({
    anchor: { row: 0, col: 0 },
    focus: { row: 0, col: 0 },
  });
  const [editing, setEditing] = useState<EditingCell | null>(null);
  const [formulaText, setFormulaText] = useState('');
  const [historyVersion, setHistoryVersion] = useState(0);
  const [zoom, setZoom] = useState(() => session.view.zoom ?? 1);
  // 功能：保存当前边框工具的线宽选择，供后续边框预设复用；时间：2026-10-24；作者：lq。
  const [borderWeight, setBorderWeight] = useState<BorderWeight>('thin');
  // 功能：边框颜色和线型随当前编辑会话保留，供连续设置使用；时间：2026-10-24；作者：lq。
  const [borderColor, setBorderColor] = useState('#1D2939');
  const [findOpen, setFindOpen] = useState(false);
  const [findQuery, setFindQuery] = useState('');
  const [findActive, setFindActive] = useState(0);
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [filterState, setFilterState] = useState<FilterState | null>(null);
  const [filterOpen, setFilterOpen] = useState(false);
  // 功能：自定义排序可选择字段、方向和表头；时间：2026-10-24；作者：lq。
  const [sortDraft, setSortDraft] = useState<{ target: CellRange; column: number; ascending: boolean; header: boolean } | null>(null);
  // 功能：保存页面设置弹窗草稿，取消时不污染当前工作表；时间：2026-10-24；作者：lq。
  const [pageSetupOpen, setPageSetupOpen] = useState(false);
  const [pageSetupDraft, setPageSetupDraft] = useState<SheetPageSetup | null>(null);
  // 功能：打印前展示全部分页并允许直接切换纸张方向；时间：2026-10-29；作者：lq。
  const [printPreviewOpen, setPrintPreviewOpen] = useState(false);
  const [printPreviewSetup, setPrintPreviewSetup] = useState<SheetPageSetup | null>(null);
  const [printPreviewCount, setPrintPreviewCount] = useState(0);
  // 功能：与 Word 预览一致的缩放切换，方便在小窗口里看实际字号；时间：2026-10-29；作者：lq。
  const [printPreviewScaleMode, setPrintPreviewScaleMode] = useState<'fit' | 'actual'>('fit');
  const printPreviewRef = useRef<HTMLDivElement>(null);
  const printPreviewCleanupRef = useRef<(() => void) | null>(null);
  const [contextMenu, setContextMenu] = useState<{
    x: number;
    y: number;
    kind: 'cell' | 'row' | 'col';
    row: number;
    col: number;
  } | null>(null);
  // 功能：格式弹窗使用独立草稿，图片文件由当前单元格定位插入；时间：2026-10-24；作者：lq。
  const [formatTarget, setFormatTarget] = useState<{ tab: '数字' | '边框'; range: CellRange; style: CellStyle; value: string | number | boolean | null } | null>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);
  const [imagePickerError, setImagePickerError] = useState<string | null>(null);
  const imageTargetRef = useRef<{ sheetId: string; row: number; col: number } | null>(null);
  const [sheetDialog, setSheetDialog] = useState<{ sheetId: string; name: string } | null>(null);
  const [exportState, setExportState] = useState<ExportState>({
    open: false,
    name: '',
    progress: null,
    busy: false,
  });
  const [replaceText, setReplaceText] = useState('');
  const clipboardRef = useRef<SheetClipboard | null>(null);
  const clipboardTsvRef = useRef('');
  const cancelExportRef = useRef(false);
  // 功能：取消旧导航的延迟聚焦，避免快速进入公式编辑后被网格抢走焦点；时间：2026-10-24；作者：lq。
  const gridFocusFrame = useRef<number | null>(null);
  const editorRef = useRef<HTMLDivElement>(null);
  const sheetScrollRef = useRef<HTMLDivElement>(null);

  const latestWorkbook = useCallback((): WorkbookModel => {
    const latest = useWorkspaceStore.getState().sessions[session.id]?.data as WorkbookModel | undefined;
    return latest ?? workbook;
  }, [session.id, workbook]);

  const history = historyFor(session.id);
  void historyVersion;
  const canUndo = history.past.length > 0;
  const canRedo = history.future.length > 0;

  const pushHistory = useCallback(
    (previous: WorkbookModel, key?: string) => {
      const stack = historyFor(session.id);
      const now = Date.now();
      if (key && stack.lastKey === key && now - stack.lastTime < 900) {
        stack.lastTime = now;
        return;
      }
      stack.past.push(previous);
      if (stack.past.length > 60) stack.past.shift();
      stack.future = [];
      stack.lastKey = key;
      stack.lastTime = now;
      setHistoryVersion((value) => value + 1);
    },
    [session.id],
  );

  const commit = useCallback(
    (next: WorkbookModel, options?: { key?: string; history?: boolean }) => {
      if (readonly) return;
      const previous = latestWorkbook();
      if (options?.history !== false) pushHistory(previous, options?.key);
      const recalculated = recalculateWorkbook(next);
      updateData(session.id, recalculated);
    },
    [latestWorkbook, pushHistory, readonly, session.id, updateData],
  );

  const undo = useCallback(() => {
    if (readonly) return;
    const stack = historyFor(session.id);
    const previous = stack.past.pop();
    if (!previous) return;
    stack.future.push(latestWorkbook());
    stack.lastKey = undefined;
    updateData(session.id, previous);
    setEditing(null);
    setHistoryVersion((value) => value + 1);
  }, [latestWorkbook, readonly, session.id, updateData]);

  const redo = useCallback(() => {
    if (readonly) return;
    const stack = historyFor(session.id);
    const next = stack.future.pop();
    if (!next) return;
    stack.past.push(latestWorkbook());
    stack.lastKey = undefined;
    updateData(session.id, next);
    setEditing(null);
    setHistoryVersion((value) => value + 1);
  }, [latestWorkbook, readonly, session.id, updateData]);

  const active = selection.kind ? selection.anchor : selection.focus;
  const range = selectionRange(selection);
  const activeCell = cellAt(sheet, active.row, active.col);

  // 功能：按 Excel 的 40%–250%范围统一缩放，并把比例保存到当前工作表视图；时间：2026-10-24；作者：lq。
  const changeZoom = useCallback(
    (value: number) => {
      const next = normalizeSheetZoom(value);
      setZoom(next);
      updateView(session.id, { activeSheetId: sheet.id, zoom: next });
    },
    [session.id, sheet.id, updateView],
  );

  useSheetZoomShortcuts({
    editorRef,
    scrollRef: sheetScrollRef,
    zoom,
    onZoomChange: changeZoom,
    enabled: true,
  });

  useEffect(() => {
    if (editing) return;
    setFormulaText(cellRawInput(cellAt(sheet, active.row, active.col)));
  }, [active.col, active.row, editing, sheet, workbook]);

  useEffect(() => {
    updateView(session.id, { activeSheetId: sheet.id, zoom });
  }, [sheet.id, session.id, updateView, zoom]);

  const cancelGridFocus = useCallback(() => {
    if (gridFocusFrame.current !== null) cancelAnimationFrame(gridFocusFrame.current);
    gridFocusFrame.current = null;
  }, []);
  useEffect(() => cancelGridFocus, [cancelGridFocus]);
  const focusGrid = useCallback(() => {
    cancelGridFocus();
    gridFocusFrame.current = requestAnimationFrame(() => {
      gridFocusFrame.current = null;
      const element = document.querySelector<HTMLElement>(`[data-sheet-grid="${session.id}"]`);
      element?.focus();
    });
  }, [cancelGridFocus, session.id]);

  const handleSelectionChange = useCallback((next: Selection) => {
    setSelection(next);
  }, []);

  const startEdit = useCallback(
    (row: number, col: number, initialText?: string) => {
      if (readonly) return;
      cancelGridFocus();
      const current = latestWorkbook();
      const currentSheet = current.sheets.find((item) => item.id === sheet.id) ?? sheet;
      const raw = initialText ?? cellRawInput(cellAt(currentSheet, row, col));
      setEditing({ row, col, initialText: raw });
      setFormulaText(raw);
      setSelection({ anchor: { row, col }, focus: { row, col } });
    },
    [cancelGridFocus, latestWorkbook, readonly, sheet],
  );

  const commitEdit = useCallback(
    (value: string, move: 'down' | 'right' | 'up' | 'left' | 'none') => {
      if (!editing || readonly) return;
      const current = latestWorkbook();
      const next = setCellInput(current, sheet.id, editing.row, editing.col, completeFormulaParentheses(value));
      const moved =
        move === 'down'
          ? { row: Math.min(sheet.dimension.rows - 1, editing.row + 1), col: editing.col }
          : move === 'right'
            ? { row: editing.row, col: Math.min(sheet.dimension.cols - 1, editing.col + 1) }
            : move === 'up'
              ? { row: Math.max(0, editing.row - 1), col: editing.col }
              : move === 'left'
                ? { row: editing.row, col: Math.max(0, editing.col - 1) }
                : { row: editing.row, col: editing.col };
      commit(next, { key: `cell:${editing.row}:${editing.col}` });
      setEditing(null);
      setSelection({ anchor: moved, focus: moved });
      focusGrid();
    },
    [commit, editing, focusGrid, latestWorkbook, readonly, sheet.dimension.cols, sheet.dimension.rows, sheet.id],
  );

  const cancelEdit = useCallback(() => {
    setEditing(null);
    setFormulaText(cellRawInput(cellAt(sheet, active.row, active.col)));
    focusGrid();
  }, [active.col, active.row, focusGrid, sheet]);

  const applyStyle = useCallback(
    (patch: Partial<CellStyle>) => {
      if (readonly) return;
      let target = range;
      if (rangeArea(target) > 20_000) {
        target = usedRange(sheet);
        notify('info', '选区过大', '已将格式应用到当前数据区域。');
      }
      const next = applyStyleToRange(latestWorkbook(), sheet.id, target, patch);
      commit(next, { key: `style:${JSON.stringify(patch).slice(0, 40)}` });
    },
    [commit, latestWorkbook, notify, range, readonly, sheet],
  );

  // 功能：右键格式设置保持整个选区，避免修改未选中的单元格；时间：2026-10-24；作者：lq。
  const openCellFormat = useCallback((tab: '数字' | '边框' = '数字') => {
    if (readonly) return;
    const target = rangeArea(range) > 20_000 ? usedRange(sheet) : range;
    if (target !== range) notify('info', '选区过大', '已将格式设置限定到当前数据区域。');
    setFormatTarget({ tab, range: target, style: cellStyle(workbook, cellAt(sheet, active.row, active.col)), value: workbook.computed?.[sheet.id]?.[cellKey(active.row, active.col)]?.value ?? cellAt(sheet, active.row, active.col)?.value ?? null });
  }, [active.col, active.row, notify, range, readonly, sheet, workbook]);

  // 功能：兼容 Excel 的设置单元格格式快捷键；时间：2026-10-24；作者：lq。
  useEffect(() => {
    const keydown = (event: KeyboardEvent) => {
      if (!(event.metaKey || event.ctrlKey) || event.key !== '1' || !editorRef.current?.contains(document.activeElement)) return;
      event.preventDefault(); openCellFormat();
    };
    document.addEventListener('keydown', keydown);
    return () => document.removeEventListener('keydown', keydown);
  }, [openCellFormat]);

  const requestImage = useCallback(() => {
    if (readonly) return;
    if (workbook.csv) { notify('info', '请先另存为 Excel 工作簿', 'CSV 文件只能保存文本，不能保存图片。'); return; }
    imageTargetRef.current = { sheetId: sheet.id, row: active.row, col: active.col };
    // 功能：保留点击事件的用户激活状态，避免延迟打开文件选择器被浏览器拦截；时间：2026-10-24；作者：lq。
    try { openSheetImagePicker(imageInputRef.current); }
    catch (error) {
      setImagePickerError(error instanceof Error ? error.message : '浏览器未能自动打开文件选择器。');
    }
  }, [active.col, active.row, notify, readonly, sheet.id, workbook.csv]);

  // 功能：读取本地图片并限制初始显示尺寸，内容嵌入工作簿而不依赖原文件；时间：2026-10-24；作者：lq。
  const insertImageFile = useCallback(async (file: File) => {
    const target = imageTargetRef.current;
    if (!target || readonly) return;
    setImagePickerError(null);
    try {
      const src = await new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result)); reader.onerror = () => reject(new Error('图片读取失败')); reader.readAsDataURL(file); });
      if (!/^data:image\/(png|jpeg|gif);base64,/.test(src)) throw new Error('请选择 PNG、JPEG 或 GIF 图片。');
      const size = await new Promise<{ width: number; height: number }>((resolve, reject) => { const image = new Image(); image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight }); image.onerror = () => reject(new Error('图片无法解码，请更换文件')); image.src = src; });
      const factor = Math.min(1, 360 / size.width, 260 / size.height);
      const image: SheetImage = { id: uid(), name: file.name, src, row: target.row, col: target.col, offsetX: 0, offsetY: 0, width: Math.max(8, size.width * factor), height: Math.max(8, size.height * factor) };
      commit(updateSheet(latestWorkbook(), target.sheetId, (item) => ({ ...item, images: [...(item.images ?? []), image] })), { key: `image-insert:${image.id}` });
    } catch (error) { notify('warning', '无法插入图片', error instanceof Error ? error.message : String(error)); }
  }, [commit, latestWorkbook, notify, readonly]);

  // 功能：边框预设按整个选区的边缘应用，内部和复合边框支持单步撤销；时间：2026-10-24；作者：lq。
  const applyBorders = useCallback((preset: BorderPreset) => {
    if (readonly) return;
    const target = rangeArea(range) > 20_000 ? usedRange(sheet) : range;
    if (target !== range) notify('info', '选区过大', '已将边框限定到当前数据区域。');
    commit(applyBordersToRange(latestWorkbook(), sheet.id, target, preset, { style: borderWeight, color: borderColor }), { key: 'borders' });
  }, [borderColor, borderWeight, commit, latestWorkbook, notify, range, readonly, sheet]);

  // 功能：打开页面设置草稿，提供方向、缩放、页边距、打印网格线和首个页码；时间：2026-10-24；作者：lq。
  const openPageSetup = useCallback(() => {
    const current = sheet.pageSetup ?? defaultSheetPageSetup();
    setPageSetupDraft({ ...current, margins: { ...current.margins } });
    setPageSetupOpen(true);
  }, [sheet.pageSetup]);

  const applyPageSetup = useCallback(() => {
    if (readonly || !pageSetupDraft) return;
    const normalized: SheetPageSetup = {
      ...pageSetupDraft,
      scalePercent: Math.max(10, Math.min(400, Math.round(pageSetupDraft.scalePercent || 100))),
      fitToWidth: pageSetupDraft.fitToWidth && pageSetupDraft.fitToWidth > 0 ? Math.round(pageSetupDraft.fitToWidth) : undefined,
      fitToHeight: pageSetupDraft.fitToHeight && pageSetupDraft.fitToHeight > 0 ? Math.round(pageSetupDraft.fitToHeight) : undefined,
      firstPageNumber: pageSetupDraft.firstPageNumber && pageSetupDraft.firstPageNumber > 0 ? Math.round(pageSetupDraft.firstPageNumber) : undefined,
      margins: Object.fromEntries(
        Object.entries(pageSetupDraft.margins).map(([key, value]) => [key, Math.max(0, Math.min(5, Number(value) || 0))]),
      ) as SheetPageSetup['margins'],
    };
    commit(updateSheet(latestWorkbook(), sheet.id, (item) => ({ ...item, pageSetup: normalized })), { key: 'page-setup' });
    setPageSetupOpen(false);
    setPageSetupDraft(null);
    notify('success', '页面设置已应用');
  }, [commit, latestWorkbook, notify, pageSetupDraft, readonly, sheet.id]);

  const patchPageSetupDraft = useCallback((patch: Partial<SheetPageSetup>) => {
    setPageSetupDraft((current) => (current ? { ...current, ...patch } : current));
  }, []);

  const patchPageMargins = useCallback((key: keyof SheetPageSetup['margins'], value: number) => {
    setPageSetupDraft((current) => (current ? { ...current, margins: { ...current.margins, [key]: value } } : current));
  }, []);

  const toggleMerge = useCallback(() => {
    if (readonly) return;
    const merged = sheet.merges.some(
      (merge) =>
        merge.r1 <= range.r2 && merge.r2 >= range.r1 && merge.c1 <= range.c2 && merge.c2 >= range.c1,
    );
    const next = merged
      ? unmergeSelection(latestWorkbook(), sheet.id, range)
      : mergeSelection(latestWorkbook(), sheet.id, range);
    commit(next, { key: 'merge' });
  }, [commit, latestWorkbook, range, readonly, sheet]);

  const freeze = useCallback(() => {
    const current = sheet.freeze;
    const activeFrozen = !!current && (current.rows > 0 || current.cols > 0);
    if (activeFrozen) {
      commit(setFreeze(latestWorkbook(), sheet.id, 0, 0));
    } else {
      commit(setFreeze(latestWorkbook(), sheet.id, active.row, active.col));
    }
  }, [active.col, active.row, commit, latestWorkbook, sheet]);

  /** 功能：结构变更无法安全维护引用时保留原文档并提示原因；时间：2026-10-18；作者：lq。 */
  const changeStructure = useCallback((change: () => void) => {
    try { change(); } catch (error) {
      notify('warning', '未执行结构变更', error instanceof Error ? error.message : String(error));
    }
  }, [notify]);

  const insertRow = useCallback(() => {
    changeStructure(() => commit(insertRows(latestWorkbook(), sheet.id, active.row, 1), { key: 'structure' }));
  }, [active.row, changeStructure, commit, latestWorkbook, sheet.id]);
  const deleteRow = useCallback(() => {
    changeStructure(() => commit(deleteRows(latestWorkbook(), sheet.id, active.row, 1), { key: 'structure' }));
  }, [active.row, changeStructure, commit, latestWorkbook, sheet.id]);
  const insertColumn = useCallback(() => {
    changeStructure(() => commit(insertColumns(latestWorkbook(), sheet.id, active.col, 1), { key: 'structure' }));
  }, [active.col, changeStructure, commit, latestWorkbook, sheet.id]);
  const deleteColumn = useCallback(() => {
    changeStructure(() => commit(deleteColumns(latestWorkbook(), sheet.id, active.col, 1), { key: 'structure' }));
  }, [active.col, changeStructure, commit, latestWorkbook, sheet.id]);

  const sortByActive = useCallback(
    (ascending: boolean, column = active.col, explicitTarget?: CellRange) => {
      const current = latestWorkbook();
      const currentSheet = current.sheets.find((item) => item.id === sheet.id) ?? sheet;
      const singleCell = range.r1 === range.r2 && range.c1 === range.c2;
      let target = explicitTarget ?? (currentSheet.autoFilter ? { ...currentSheet.autoFilter, r1: currentSheet.autoFilter.r1 + 1 } : tableRangeForSelection(currentSheet, selection));
      if (!explicitTarget && singleCell && currentSheet.autoFilter) {
        // 功能：单元格排序默认保留首行为表头，避免把字段名排入数据；时间：2026-10-24；作者：lq。
        target = { ...currentSheet.autoFilter, r1: currentSheet.autoFilter.r1 + 1 };
      } else if (!explicitTarget && !currentSheet.autoFilter && (singleCell || selection.kind) && target.r2 > target.r1) {
        target = { ...target, r1: target.r1 + 1 };
      }
      // 功能：单列快速排序扩展到关联字段，防止只移动一列造成记录错位；时间：2026-10-24；作者：lq。
      if (!explicitTarget && !currentSheet.autoFilter && target.c1 === target.c2) {
        const used = usedRange(currentSheet);
        target = { ...target, c1: used.c1, c2: used.c2 };
      }
      const sortColumn = Math.min(target.c2, Math.max(target.c1, column));
      if (target.r2 <= target.r1) {
        notify('info', '没有可排序的数据', '请选择包含数据的区域。');
        return;
      }
      if (currentSheet.merges.some((merge) => merge.r1 <= target.r2 && merge.r2 >= target.r1 && merge.c1 <= target.c2 && merge.c2 >= target.c1)) {
        notify('warning', '无法排序', '排序区域包含合并单元格，请先取消合并。');
        return;
      }
      changeStructure(() => {
        let next = sortRange(current, currentSheet, target, sortColumn, ascending);
        const sorted = next.sheets.find((item) => item.id === sheet.id)!;
        const first = Object.entries(sorted.filterCriteria ?? {})[0];
        if (first && sorted.autoFilter) next = applyFilter(next, sorted, Number(first[0]), new Set(first[1]), { ...sorted.autoFilter, r1: sorted.autoFilter.r1 + 1 });
        commit(next, { key: 'sort' });
        setSortDraft(null);
        setFilterOpen(false);
        notify('success', ascending ? '已按当前列升序排序' : '已按当前列降序排序');
      });
    },
    [active.col, changeStructure, commit, latestWorkbook, notify, range, selection, sheet],
  );

  // 功能：表头按钮打开当前列菜单，恢复该列已应用的勾选条件；时间：2026-10-24；作者：lq。
  const openFilterColumn = useCallback((column: number, anchor: { left: number; bottom: number }) => {
    if (readonly) return;
    const current = latestWorkbook();
    const currentSheet = current.sheets.find((item) => item.id === sheet.id) ?? sheet;
    const target = currentSheet.autoFilter;
    if (!target) return;
    const dataRange = { ...target, r1: target.r1 + 1 };
    const values = uniqueColumnValues(current, currentSheet, column, dataRange);
    setFilterState({ column, range: dataRange, autoFilterRange: target, values,
      selected: new Set(currentSheet.filterCriteria?.[column] ?? values), anchor });
    setFilterOpen(true);
  }, [latestWorkbook, readonly, sheet]);

  // 功能：清除条件时保留表头下拉按钮，关闭筛选时恢复原有隐藏行；时间：2026-10-24；作者：lq。
  const clearFilter = useCallback((disable = false) => {
    const current = latestWorkbook();
    commit(updateSheet(current, sheet.id, (item) => ({ ...item,
      hiddenRows: item.filterBaseHiddenRows ?? item.hiddenRows,
      filterCriteria: undefined, filterBaseHiddenRows: undefined,
      autoFilter: disable ? undefined : item.autoFilter,
    })), { key: 'filter-clear' });
    setFilterOpen(false);
  }, [commit, latestWorkbook, sheet.id]);

  const openFilter = useCallback(() => {
    if (readonly) return;
    if (sheet.autoFilter) { clearFilter(true); return; }
    const target = tableRangeForSelection(sheet, selection);
    if (target.r2 <= target.r1) { notify('info', '没有可筛选的数据', '请选择包含表头和数据行的区域。'); return; }
    commit(updateSheet(latestWorkbook(), sheet.id, (item) => ({ ...item, autoFilter: target })), { key: 'filter-enable' });
  }, [clearFilter, commit, latestWorkbook, notify, readonly, selection, sheet]);

  const applyFilterDialog = useCallback((clearColumn = false, selected = filterState?.selected) => {
    if (!filterState || !selected || readonly) return;
    const current = latestWorkbook();
    const currentSheet = current.sheets.find((item) => item.id === sheet.id)!;
    const allowed = clearColumn || selected.size === filterState.values.length ? null : selected;
    commit(applyFilter(current, currentSheet, filterState.column, allowed, filterState.range), { key: 'filter' });
    setFilterOpen(false);
  }, [commit, filterState, latestWorkbook, readonly, sheet.id]);

  const reapplyFilter = useCallback(() => {
    const current = latestWorkbook();
    const currentSheet = current.sheets.find((item) => item.id === sheet.id)!;
    const first = Object.entries(currentSheet.filterCriteria ?? {})[0];
    if (first && currentSheet.autoFilter) commit(applyFilter(current, currentSheet, Number(first[0]), new Set(first[1]), { ...currentSheet.autoFilter, r1: currentSheet.autoFilter.r1 + 1 }), { key: 'filter' });
  }, [commit, latestWorkbook, sheet.id]);

  const openCustomSort = useCallback(() => {
    const target = sheet.autoFilter ?? (tableRangeForSelection(sheet, selection));
    setSortDraft({ target, column: Math.min(target.c2, Math.max(target.c1, active.col)), ascending: true, header: true });
  }, [active.col, selection, sheet]);

  const findHits = useMemo(() => {
    if (!findQuery) return [] as Array<{ row: number; col: number }>;
    const target = usedRange(sheet);
    const hits: Array<{ row: number; col: number }> = [];
    const needle = caseSensitive ? findQuery : findQuery.toLowerCase();
    for (let row = target.r1; row <= target.r2; row += 1) {
      for (let col = target.c1; col <= target.c2; col += 1) {
        const text = cellDisplay(workbook, sheet, row, col);
        if (!text) continue;
        const haystack = caseSensitive ? text : text.toLowerCase();
        if (haystack.includes(needle)) hits.push({ row, col });
        if (hits.length >= 2000) return hits;
      }
    }
    return hits;
  }, [caseSensitive, findQuery, sheet, workbook]);

  const goToHit = useCallback(
    (index: number) => {
      if (findHits.length === 0) return;
      const normalized = ((index % findHits.length) + findHits.length) % findHits.length;
      setFindActive(normalized);
      const hit = findHits[normalized];
      const next = { anchor: { row: hit.row, col: hit.col }, focus: { row: hit.row, col: hit.col } };
      setSelection(next);
      setFormulaText(cellRawInput(cellAt(sheet, hit.row, hit.col)));
      focusGrid();
    },
    [findHits, focusGrid, sheet],
  );

  const replaceCurrent = useCallback(() => {
    if (readonly || findHits.length === 0 || !findQuery) return;
    const hit = findHits[Math.min(findActive, findHits.length - 1)];
    const cell = cellAt(sheet, hit.row, hit.col);
    if (cell?.formula) {
      notify('warning', '公式单元格', '查找替换不会修改公式本身，请直接编辑公式。');
      return;
    }
    const display = cellDisplay(workbook, sheet, hit.row, hit.col);
    const haystack = caseSensitive ? display : display.toLowerCase();
    const needle = caseSensitive ? findQuery : findQuery.toLowerCase();
    const index = haystack.indexOf(needle);
    if (index < 0) return;
    const replaced = display.slice(0, index) + replaceText + display.slice(index + findQuery.length);
    const next = setCellInput(latestWorkbook(), sheet.id, hit.row, hit.col, replaced);
    commit(next, { key: `replace:${hit.row}:${hit.col}` });
  }, [caseSensitive, commit, findActive, findHits, findQuery, latestWorkbook, notify, readonly, replaceText, sheet, workbook]);

  const replaceAll = useCallback(() => {
    if (readonly || !findQuery) return;
    const current = latestWorkbook();
    let total = 0;
    const updates: Array<{ row: number; col: number; cell: ReturnType<typeof createReplacedCell> }> = [];
    const target = usedRange(sheet);
    const needle = caseSensitive ? findQuery : findQuery.toLowerCase();
    for (let row = target.r1; row <= target.r2; row += 1) {
      for (let col = target.c1; col <= target.c2; col += 1) {
        const cell = cellAt(sheet, row, col);
        if (!cell || cell.formula) continue;
        const display = cellDisplay(workbook, sheet, row, col);
        const haystack = caseSensitive ? display : display.toLowerCase();
        if (!haystack.includes(needle)) continue;
        const replaced = caseSensitive
          ? display.split(findQuery).join(replaceText)
          : replaceCaseInsensitive(display, findQuery, replaceText);
        if (replaced !== display) {
          updates.push({ row, col, cell: createReplacedCell(replaced, cell) });
          total += 1;
        }
      }
    }
    if (updates.length === 0) {
      notify('info', '没有可替换的内容');
      return;
    }
    commit(applyCellUpdates(current, sheet.id, updates), { key: 'replace-all' });
    notify('success', `已替换 ${total} 个单元格`);
  }, [caseSensitive, commit, findQuery, latestWorkbook, notify, readonly, replaceText, sheet, workbook]);

  const copySelection = useCallback(
    async (cut: boolean) => {
      const current = latestWorkbook();
      const currentSheet = current.sheets.find((item) => item.id === sheet.id) ?? sheet;
      const clipboard = extractClipboard(current, currentSheet, range);
      clipboardRef.current = clipboard;
      const tsv = clipboardToTsvText(current, currentSheet, clipboard);
      clipboardTsvRef.current = tsv;
      try {
        await navigator.clipboard.writeText(tsv);
      } catch {
        // Clipboard permissions may be denied; the internal clipboard still works.
      }
      if (cut && !readonly) {
        commit(clearRangeContents(current, sheet.id, range), { key: 'cut' });
      }
    },
    [commit, latestWorkbook, range, readonly, sheet],
  );

  const pasteSelection = useCallback(async () => {
    if (readonly) return;
    let text = '';
    try {
      text = await navigator.clipboard.readText();
    } catch {
      text = '';
    }
    const current = latestWorkbook();
    const currentSheet = current.sheets.find((item) => item.id === sheet.id) ?? sheet;
    if (text && text.trim() && text !== clipboardTsvRef.current) {
      const rows = parseTsv(text);
      commit(pasteTsv(current, currentSheet, active, rows), { key: 'paste' });
      notify('success', `已粘贴 ${rows.length} 行`);
      return;
    }
    const clipboard = clipboardRef.current;
    if (!clipboard) {
      notify('info', '剪贴板为空');
      return;
    }
    const target =
      range.r1 === range.r2 && range.c1 === range.c2
        ? {
            r1: active.row,
            c1: active.col,
            r2: active.row + (clipboard.range.r2 - clipboard.range.r1),
            c2: active.col + (clipboard.range.c2 - clipboard.range.c1),
          }
        : range;
    commit(pasteClipboard(current, currentSheet, target, clipboard), { key: 'paste' });
  }, [active, commit, latestWorkbook, notify, range, readonly, sheet]);

  const handleAction = useCallback(
    (action: GridAction) => {
      if (action === 'copy') void copySelection(false);
      else if (action === 'cut') void copySelection(true);
      else if (action === 'paste') void pasteSelection();
      else if (action === 'undo') undo();
      else if (action === 'redo') redo();
      else if (action === 'selectAll') {
        const target = usedRange(sheet);
        setSelection({ kind: 'all', anchor: { row: target.r1, col: target.c1 }, focus: { row: target.r2, col: target.c2 } });
      } else if (action === 'clear') {
        if (!readonly) commit(clearRangeContents(latestWorkbook(), sheet.id, range), { key: 'clear' });
      } else if (action === 'find') setFindOpen(true);
    },
    [commit, copySelection, latestWorkbook, pasteSelection, range, readonly, redo, sheet, undo],
  );

  const handleFill = useCallback(
    (source: CellRange, target: CellRange) => {
      if (readonly) return;
      const current = latestWorkbook();
      const currentSheet = current.sheets.find((item) => item.id === sheet.id) ?? sheet;
      const clipboard = extractClipboard(current, currentSheet, source);
      const next = pasteClipboard(current, currentSheet, target, clipboard);
      commit(next, { key: 'fill' });
    },
    [commit, latestWorkbook, readonly, sheet],
  );

  const handleCellContextMenu = useCallback(
    (event: ReactMouseEvent, row: number, col: number) => {
      setContextMenu({ x: event.clientX, y: event.clientY, kind: 'cell', row, col });
    },
    [],
  );

  const handleHeaderContextMenu = useCallback(
    (event: ReactMouseEvent, kind: 'row' | 'col', index: number) => {
      setContextMenu({ x: event.clientX, y: event.clientY, kind, row: kind === 'row' ? index : active.row, col: kind === 'col' ? index : active.col });
    },
    [active.col, active.row],
  );

  const contextItems: MenuItem[] = contextMenu
    ? contextMenu.kind === 'cell'
      ? [
          { id: 'copy', label: '复制', shortcut: '⌘C', onSelect: () => void copySelection(false) },
          { id: 'cut', label: '剪切', shortcut: '⌘X', onSelect: () => void copySelection(true) },
          { id: 'paste', label: '粘贴', shortcut: '⌘V', onSelect: () => void pasteSelection() },
          { type: 'separator' },
          { id: 'insert-row', label: '在上方插入行', onSelect: insertRow },
          { id: 'insert-col', label: '在左侧插入列', onSelect: insertColumn },
          { id: 'delete-row', label: '删除当前行', danger: true, onSelect: deleteRow },
          { id: 'delete-col', label: '删除当前列', danger: true, onSelect: deleteColumn },
          { type: 'separator' },
          { id: 'clear', label: '清除内容', onSelect: () => handleAction('clear') },
          { id: 'freeze', label: '冻结到当前单元格', onSelect: freeze },
          { type: 'separator' },
          { id: 'format-cell', label: '设置单元格格式…', shortcut: '⌘1', disabled: readonly, onSelect: openCellFormat },
          { id: 'insert-image', label: '插入图片…', disabled: readonly, onSelect: requestImage },
        ]
      : [
          ...(contextMenu.kind === 'row'
            ? [
                { id: 'insert', label: '在上方插入行', onSelect: insertRow },
                { id: 'delete', label: '删除行', danger: true, onSelect: deleteRow },
              ]
            : [
                { id: 'insert', label: '在左侧插入列', onSelect: insertColumn },
                { id: 'delete', label: '删除列', danger: true, onSelect: deleteColumn },
              ]),
          { type: 'separator' },
          { id: 'sort-asc', label: '升序排序', onSelect: () => sortByActive(true) },
          { id: 'sort-desc', label: '降序排序', onSelect: () => sortByActive(false) },
          { id: 'filter', label: '筛选…', onSelect: openFilter },
          { id: 'format-cell', label: '设置单元格格式…', disabled: readonly, onSelect: openCellFormat },
        ]
    : [];

  const activeSheetIndex = workbook.sheets.findIndex((item) => item.id === sheet.id);
  const stats = useMemo(() => selectionStats(workbook, sheet, range), [range, sheet, workbook]);

  const switchSheet = useCallback(
    (nextId: string) => {
      setSheetId(nextId);
      updateView(session.id, { activeSheetId: nextId });
      setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } });
      setEditing(null);
    },
    [session.id, updateView, workbook],
  );

  const addNewSheet = useCallback(() => {
    if (readonly) return;
    const next = addSheet(latestWorkbook());
    commit(next);
    const newSheet = next.sheets[next.sheets.length - 1];
    setSheetId(newSheet.id);
    setSelection({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } });
  }, [commit, latestWorkbook, readonly]);

  const confirmSheetDialog = useCallback(() => {
    if (!sheetDialog || readonly) return;
    changeStructure(() => {
      const next = deleteSheet(latestWorkbook(), sheetDialog.sheetId);
      commit(next);
      if (sheetDialog.sheetId === sheet.id) setSheetId(next.sheets.find((item) => item.state === 'visible')!.id);
      setSheetDialog(null);
    });
  }, [changeStructure, commit, latestWorkbook, readonly, sheetDialog, sheet.id]);

  const openExport = useCallback(() => {
    setExportState({ open: true, name: `${stripExtension(session.displayName)}.pdf`, progress: null, busy: false });
  }, [session.displayName]);

  const runExport = useCallback(async () => {
    if (exportState.busy) return;
    cancelExportRef.current = false;
    setExportState((state) => ({ ...state, busy: true, progress: { phase: 'preparing', current: 0, total: 1, message: '正在准备表格…' } }));
    await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
    const container = document.createElement('div');
    container.dataset.sheetPrint = 'true';
    try {
      const pages = buildSheetPrintPages(workbook, sheet, container);
      if (pages.length === 0) throw new Error('没有可导出的内容');
      const bytes = await exportElementsToPdf({
        pages,
        fileName: ensureExtension(stripExtension(exportState.name || session.displayName), '.pdf'),
        title: session.displayName,
        scale: 2,
        pageWidthPx: sheetPageSize(workbook.sheets.find((item) => item.id === sheet.id)?.pageSetup ?? defaultSheetPageSetup()).width,
        pageHeightPx: sheetPageSize(workbook.sheets.find((item) => item.id === sheet.id)?.pageSetup ?? defaultSheetPageSetup()).height,
        onProgress: (progress) => setExportState((state) => ({ ...state, progress })),
        isCancelled: () => cancelExportRef.current,
      });
      await exportBytes(
        bytes,
        ensureExtension(stripExtension(exportState.name || session.displayName), '.pdf'),
        'application/pdf',
      );
      setExportState({ open: false, name: '', progress: null, busy: false });
    } catch (error) {
      if ((error as { code?: string })?.code === 'CANCELLED') notify('info', '已取消导出');
      else notify('error', '导出 PDF 失败', error instanceof Error ? error.message : String(error));
      setExportState((state) => ({ ...state, busy: false, progress: null }));
    } finally {
      container.remove();
    }
  }, [exportBytes, exportState.busy, exportState.name, notify, session.displayName, sheet, workbook]);

  const printSheet = useCallback(async (requestedSetup?: SheetPageSetup) => {
    const setup = requestedSetup ?? sheet.pageSetup ?? defaultSheetPageSetup();
    const container = document.createElement('div');
    container.dataset.sheetPrint = 'true';
    try {
      // 功能：打印确认后的方向、缩放和边距使用同一份页面设置生成页面，页数即预览页数；时间：2026-10-29；作者：lq。
      const pages = buildSheetPrintPages(workbook, { ...sheet, pageSetup: setup }, container);
      const size = sheetPageSize(setup);
      const ok = await printPages({
        pages,
        widthPx: size.width,
        heightPx: size.height,
        title: sheet.name,
      });
      if (!ok) notify('error', '打印不可用', '当前平台不支持打印，请改用“导出 PDF”。');
    } finally {
      container.remove();
    }
  }, [notify, sheet, workbook]);

  const openPrintPreview = useCallback(() => {
    printPreviewCleanupRef.current?.();
    setPrintPreviewSetup({ ...(sheet.pageSetup ?? defaultSheetPageSetup()), margins: { ...(sheet.pageSetup ?? defaultSheetPageSetup()).margins } });
    setPrintPreviewOpen(true);
  }, [sheet.pageSetup]);

  useEffect(() => {
    if (!printPreviewOpen || !printPreviewSetup || !printPreviewRef.current) return;
    const host = printPreviewRef.current;
    host.replaceChildren();
    const staging = document.createElement('div');
    const previewSheet = { ...sheet, pageSetup: printPreviewSetup };
    const pages = buildSheetPrintPages(workbook, previewSheet, staging);
    const size = sheetPageSize(printPreviewSetup);
    // 功能：预览复用打印模块的公共渲染，保证与 Word 预览视觉和页数完全一致；时间：2026-10-29；作者：lq。
    const cleanupPreview = renderPrintPreview({
      host,
      pages,
      widthPx: size.width,
      heightPx: size.height,
      scale: printPreviewScaleMode === 'actual' ? 1 : undefined,
    });
    setPrintPreviewCount(pages.length);
    printPreviewCleanupRef.current = () => { staging.remove(); cleanupPreview(); };
    return () => printPreviewCleanupRef.current?.();
  }, [printPreviewOpen, printPreviewScaleMode, printPreviewSetup, sheet, workbook]);

  const confirmPrintPreview = useCallback(async () => {
    if (!printPreviewSetup) return;
    setPrintPreviewOpen(false);
    printPreviewCleanupRef.current?.();
    if (!readonly) commit(updateSheet(latestWorkbook(), sheet.id, (item) => ({ ...item, pageSetup: printPreviewSetup })), { key: 'page-setup-print' });
    await printSheet(printPreviewSetup);
  }, [commit, latestWorkbook, printPreviewSetup, printSheet, readonly, sheet.id]);

  useEffect(() => {
    const handleFind = () => setFindOpen(true);
    const handleExport = () => openExport();
    const handlePrint = () => openPrintPreview();
    window.addEventListener('mo:find', handleFind);
    window.addEventListener('mo:export-pdf', handleExport);
    window.addEventListener('mo:print', handlePrint);
    return () => {
      window.removeEventListener('mo:find', handleFind);
      window.removeEventListener('mo:export-pdf', handleExport);
      window.removeEventListener('mo:print', handlePrint);
    };
  }, [openExport, openPrintPreview]);

  const formulaCommit = useCallback(() => {
    if (readonly) return;
    if (editing) {
      commitEdit(formulaText, 'none');
      return;
    }
    // 功能：失焦与回车可能连续发生，相同内容不重复写入历史；时间：2026-10-24；作者：lq。
    const value = completeFormulaParentheses(formulaText);
    if (cellRawInput(cellAt(sheet, active.row, active.col)) === value) return;
    const next = setCellInput(latestWorkbook(), sheet.id, active.row, active.col, value);
    commit(next, { key: `cell:${active.row}:${active.col}` });
  }, [active.col, active.row, commit, commitEdit, editing, formulaText, latestWorkbook, readonly, sheet.id]);

  const handleJump = useCallback(
    (ref: string) => {
      const parsed = parseRangeRef(ref);
      if (!parsed) {
        notify('warning', '无效的单元格地址', '请输入如 A1 或 A1:C10 的地址。');
        return;
      }
      setSelection({
        anchor: { row: parsed.r1, col: parsed.c1 },
        focus: { row: parsed.r2, col: parsed.c2 },
      });
      focusGrid();
    },
    [focusGrid, notify],
  );

  return (
    <div ref={editorRef} className="flex h-full min-h-0 w-full">
      <div className="flex min-w-0 flex-1 flex-col">
        {settings.showToolbar ? (
        <SheetToolbar
          workbook={workbook}
          sheet={sheet}
          range={range}
          active={active}
          formulaText={formulaText}
          readonly={readonly}
          canUndo={canUndo}
          canRedo={canRedo}
          zoom={zoom}
          propertiesOpen={settings.showProperties}
          onFormulaChange={setFormulaText}
          onFormulaFocus={cancelGridFocus}
          onFormulaCommit={formulaCommit}
          onFormulaCancel={cancelEdit}
          onJump={handleJump}
          onStyle={applyStyle}
          onNumberFormat={(format) => applyStyle({ numFmt: format })}
          onBorders={applyBorders}
          borderWeight={borderWeight}
          onBorderWeight={setBorderWeight}
          borderColor={borderColor} onBorderColor={setBorderColor}
          onMoreBorders={() => openCellFormat('边框')}
          onMergeToggle={toggleMerge}
          onFreeze={freeze}
          onInsertRow={insertRow}
          onDeleteRow={deleteRow}
          onInsertColumn={insertColumn}
          onDeleteColumn={deleteColumn}
          onSort={sortByActive}
          onInsertImage={requestImage}
          onFilter={openFilter}
          onCustomSort={openCustomSort}
          onClearFilter={() => clearFilter()}
          onReapplyFilter={reapplyFilter}
          onFind={() => setFindOpen(true)}
          onExport={openExport}
          onPrint={openPrintPreview}
          onPageSetup={openPageSetup}
          onCompatibility={() => openCompatibility({ type: 'compatibility', sessionId: session.id })}
          onToggleProperties={() => updateSettings({ showProperties: !settings.showProperties })}
          onZoom={setZoom}
          onUndo={undo}
          onRedo={redo}
          onAddSheet={addNewSheet}
        />
        ) : null}

        <div className="relative min-h-0 flex-1" style={{ zoom }}>
          <SheetGrid
            gridId={session.id}
            scrollElementRef={sheetScrollRef}
            workbook={workbook}
            sheet={sheet}
            selection={selection}
            editing={editing}
            formulaDraft={formulaText} onDraftChange={setFormulaText}
            onSelectionChange={handleSelectionChange}
            onStartEdit={startEdit}
            onCommitEdit={commitEdit}
            onCancelEdit={cancelEdit}
            onColumnResize={(col, width) => commit(setColumnWidth(latestWorkbook(), sheet.id, col, width))}
            onRowResize={(row, height) => commit(setRowHeight(latestWorkbook(), sheet.id, row, height))}
            onCellContextMenu={handleCellContextMenu}
            onHeaderContextMenu={handleHeaderContextMenu}
            onAction={handleAction}
            onFill={handleFill}
            onFilterColumn={openFilterColumn}
            readonly={readonly} zoom={zoom}
            onImageChange={(id, patch) => { if (!readonly) commit(updateSheet(latestWorkbook(), sheet.id, (item) => ({ ...item, images: item.images?.map((image) => image.id === id ? { ...image, ...patch } : image) })), { key: `image:${id}` }); }}
            onImageDelete={(id) => { if (!readonly) commit(updateSheet(latestWorkbook(), sheet.id, (item) => ({ ...item, images: item.images?.filter((image) => image.id !== id) })), { key: `image-delete:${id}` }); }}
            highlightRowCol={settings.highlightRowCol}
          />
          {findOpen ? (
            <SheetFindBar
              query={findQuery}
              replaceText={replaceText}
              caseSensitive={caseSensitive}
              count={findHits.length}
              active={findActive}
              readonly={readonly}
              onQuery={(value) => {
                setFindQuery(value);
                setFindActive(0);
              }}
              onReplace={setReplaceText}
              onCaseSensitive={setCaseSensitive}
              onPrev={() => goToHit(findActive - 1)}
              onNext={() => goToHit(findActive + 1)}
              onReplaceCurrent={replaceCurrent}
              onReplaceAll={replaceAll}
              onClose={() => setFindOpen(false)}
            />
          ) : null}
        </div>

        <div className="mo-glass mo-no-print flex h-8 shrink-0 items-center gap-1 border-t border-[var(--mo-line)] bg-[var(--mo-surface)] px-2">
          {/* 功能：标签右键菜单、双击原位改名及标签颜色；时间：2026-10-24；作者：lq。 */}
          <SheetTabs sheets={workbook.sheets} activeId={sheet.id} readonly={readonly} onSwitch={switchSheet} onAdd={addNewSheet}
            onRename={(id, name) => { if (!readonly) commit(renameSheet(latestWorkbook(), id, name), { key: 'rename-sheet' }); }}
            onColor={(id, tabColor) => { if (!readonly) commit(updateSheet(latestWorkbook(), id, (item) => ({ ...item, tabColor })), { key: 'tab-color' }); }}
            onDelete={(item) => setSheetDialog({ sheetId: item.id, name: item.name })} />
          <div className="flex shrink-0 items-center gap-3 pl-3 text-[10px] tabular-nums text-[var(--mo-text-muted)]">
            {stats.numericCount > 0 ? (
              <>
                <span>求和 {formatNumber(stats.sum)}</span>
                <span>平均 {formatNumber(stats.average)}</span>
                <span>计数 {stats.count}</span>
              </>
            ) : stats.count > 0 ? (
              <span>非空 {stats.count}</span>
            ) : (
              <span>
                {columnLabel(active.col)}
                {active.row + 1}
              </span>
            )}
            <span className="text-[var(--mo-text-faint)]">工作表 {activeSheetIndex + 1}/{workbook.sheets.length}</span>
          </div>
        </div>
      </div>

      {settings.showProperties ? (
        <SheetPropertiesPanel
          workbook={workbook}
          sheet={sheet}
          range={range}
          active={active}
          session={session}
          onStyle={applyStyle}
          onClose={() => updateSettings({ showProperties: false })}
          onShowCompatibility={() => openCompatibility({ type: 'compatibility', sessionId: session.id })}
        />
      ) : null}

      {/* 功能：以单次历史提交应用格式草稿和选区边框；时间：2026-10-24；作者：lq。 */}
      {formatTarget ? <CellFormatDialog initialTab={formatTarget.tab} style={formatTarget.style} value={formatTarget.value} onClose={() => setFormatTarget(null)} onApply={(patch, border) => {
        if (!readonly) commit(applyCellFormat(latestWorkbook(), sheet.id, formatTarget.range, patch, border), { key: 'format-cells' });
        setFormatTarget(null);
      }} /> : null}
      <input ref={imageInputRef} type="file" aria-label="插入工作表图片" accept="image/png,image/jpeg,image/gif" className="sr-only" tabIndex={-1} onChange={(event) => { const file = event.target.files?.[0]; if (file) void insertImageFile(file); event.target.value = ''; }} />

      {/* 功能：浏览器拒绝自动文件选择时显示原生文件控件，两个插图入口共用重试路径；时间：2026-10-24；作者：lq。 */}
      <Dialog open={imagePickerError !== null} title="选择插入的图片" width="sm" onClose={() => setImagePickerError(null)}
        description="未能自动打开文件选择窗口，请在下方直接选择 PNG、JPEG 或 GIF 图片。"
        footer={<Button onClick={() => setImagePickerError(null)}>取消</Button>}>
        <input type="file" aria-label="直接选择插入图片" accept="image/png,image/jpeg,image/gif" className="block w-full rounded border border-[var(--mo-line)] p-3 text-[12px]" onChange={(event) => {
          const file = event.target.files?.[0]; if (file) void insertImageFile(file); event.target.value = '';
        }} />
      </Dialog>

      {contextMenu ? (
        <ContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          items={contextItems}
          onClose={() => setContextMenu(null)}
        />
      ) : null}

      <Dialog
        open={printPreviewOpen}
        onClose={() => { setPrintPreviewOpen(false); printPreviewCleanupRef.current?.(); }}
        title="打印预览"
        description={`预览全部分页（共 ${printPreviewCount} 页）并选择打印方向，确认后调用系统打印。`}
        width="xl"
        footer={<><Button onClick={() => { setPrintPreviewOpen(false); printPreviewCleanupRef.current?.(); }}>取消</Button><Button variant="primary" onClick={() => void confirmPrintPreview()}>打印</Button></>}
      >
        {printPreviewSetup ? <div className="grid grid-cols-[180px_1fr] gap-4 text-[12px]">
          <div className="space-y-3 rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface-2)] p-3">
            <Field label="方向">
              <Select value={printPreviewSetup.orientation} onChange={(event) => setPrintPreviewSetup((current) => current ? { ...current, orientation: event.target.value as SheetPageSetup['orientation'] } : current)}>
                <option value="portrait">纵向</option>
                <option value="landscape">横向</option>
              </Select>
            </Field>
            <Field label="缩放比例（%）">
              <Input type="number" min={10} max={400} value={printPreviewSetup.scalePercent} onChange={(event) => setPrintPreviewSetup((current) => current ? { ...current, scalePercent: Number(event.target.value) || 100, fitToWidth: undefined } : current)} />
            </Field>
            <label className="flex items-center gap-2"><input type="checkbox" checked={!!printPreviewSetup.fitToWidth} onChange={(event) => setPrintPreviewSetup((current) => current ? { ...current, fitToWidth: event.target.checked ? 1 : undefined } : current)} />调整为一页宽</label>
            <div className="space-y-1.5">
              <div className="text-[11px] text-[var(--mo-text-muted)]">预览比例</div>
              <SegmentedControl
                value={printPreviewScaleMode}
                onChange={setPrintPreviewScaleMode}
                options={[
                  { value: 'fit', label: '适应页宽' },
                  { value: 'actual', label: '实际大小' },
                ]}
              />
            </div>
            <p className="leading-5 text-[var(--mo-text-muted)]">共 {printPreviewCount} 页。列宽超出纸张时自动横向分页，合并单元格跨页会保留边框与底纹。</p>
          </div>
          <div ref={printPreviewRef} className="relative min-h-[300px] max-h-[460px] overflow-auto rounded-md border border-[var(--mo-line)] bg-[#667085]" aria-label="工作表打印预览" />
        </div> : null}
      </Dialog>

      <Dialog
        open={pageSetupOpen}
        onClose={() => { setPageSetupOpen(false); setPageSetupDraft(null); }}
        title="页面设置"
        description="设置当前工作表的打印方向、缩放、页边距和打印辅助项。"
        width="md"
        footer={
          <>
            <Button onClick={() => { setPageSetupOpen(false); setPageSetupDraft(null); }}>取消</Button>
            <Button variant="primary" onClick={applyPageSetup}>确定</Button>
          </>
        }
      >
        {pageSetupDraft ? (
          <div className="space-y-4 text-[12px]">
            <section className="grid grid-cols-2 gap-3">
              <Field label="方向">
                <Select
                  value={pageSetupDraft.orientation}
                  onChange={(event) => patchPageSetupDraft({ orientation: event.target.value as SheetPageSetup['orientation'] })}
                >
                  <option value="portrait">纵向</option>
                  <option value="landscape">横向</option>
                </Select>
              </Field>
              <Field label="缩放比例（%）">
                <Input
                  type="number"
                  min={10}
                  max={400}
                  value={pageSetupDraft.scalePercent}
                  onChange={(event) => patchPageSetupDraft({ scalePercent: Number(event.target.value) })}
                />
              </Field>
            </section>
            <section className="grid grid-cols-2 gap-3">
              <Field label="调整为页宽（0 为自动）">
                <Input
                  type="number"
                  min={0}
                  value={pageSetupDraft.fitToWidth ?? ''}
                  onChange={(event) => patchPageSetupDraft({ fitToWidth: Number(event.target.value) || undefined })}
                />
              </Field>
              <Field label="调整为页高（0 为自动）">
                <Input
                  type="number"
                  min={0}
                  value={pageSetupDraft.fitToHeight ?? ''}
                  onChange={(event) => patchPageSetupDraft({ fitToHeight: Number(event.target.value) || undefined })}
                />
              </Field>
            </section>
            <section className="grid grid-cols-3 gap-2">
              {([
                ['left', '左'], ['right', '右'], ['top', '上'], ['bottom', '下'], ['header', '页眉'], ['footer', '页脚'],
              ] as Array<[keyof SheetPageSetup['margins'], string]>).map(([key, label]) => (
                <Field key={key} label={`${label}边距（英寸）`}>
                  <Input
                    type="number"
                    min={0}
                    max={5}
                    step="0.05"
                    value={pageSetupDraft.margins[key]}
                    onChange={(event) => patchPageMargins(key, Number(event.target.value))}
                  />
                </Field>
              ))}
            </section>
            <section className="space-y-2 border-t border-[var(--mo-line)] pt-3">
              <Field label="首页页码（空为自动）">
                <Input
                  type="number"
                  min={1}
                  value={pageSetupDraft.firstPageNumber ?? ''}
                  onChange={(event) => patchPageSetupDraft({ firstPageNumber: Number(event.target.value) || undefined })}
                />
              </Field>
            </section>
            {/* 功能：打印辅助项单独成组并显式标注文字，去掉此前右侧无标签的悬浮开关；时间：2026-10-29；作者：lq。 */}
            <section className="space-y-1.5 border-t border-[var(--mo-line)] pt-3">
              <h3 className="text-[11px] font-semibold text-[var(--mo-text-muted)]">打印</h3>
              <label className="flex h-8 items-center justify-between rounded-md px-1 text-[12px] hover:bg-[var(--mo-surface-2)]">
                <span>打印网格线</span>
                <Switch
                  checked={pageSetupDraft.printGridLines}
                  onChange={(checked) => patchPageSetupDraft({ printGridLines: checked })}
                  label="打印网格线"
                />
              </label>
              <label className="flex h-8 items-center justify-between rounded-md px-1 text-[12px] hover:bg-[var(--mo-surface-2)]">
                <span>打印行列标题</span>
                <Switch
                  checked={pageSetupDraft.printHeadings}
                  onChange={(checked) => patchPageSetupDraft({ printHeadings: checked })}
                  label="打印行列标题"
                />
              </label>
            </section>
          </div>
        ) : null}
      </Dialog>

      {/* 功能：筛选浮层紧邻字段表头，支持搜索、多选和列排序；时间：2026-10-24；作者：lq。 */}
      {filterOpen && filterState ? <ColumnFilterMenu
        anchor={filterState.anchor}
        title={cellDisplay(workbook, sheet, filterState.autoFilterRange.r1, filterState.column) || `${columnLabel(filterState.column)} 列`}
        values={filterState.values} selected={filterState.selected}
        onSelected={(selected) => setFilterState((state) => state ? { ...state, selected } : state)}
        onSort={(ascending) => sortByActive(ascending, filterState.column, filterState.range)}
        onApply={(selected) => applyFilterDialog(false, selected)} onClear={() => applyFilterDialog(true)} onClose={() => setFilterOpen(false)}
      /> : null}
      {/* 功能：自定义排序明确字段、表头和方向，整行数据一起移动；时间：2026-10-24；作者：lq。 */}
      <Dialog open={!!sortDraft} onClose={() => setSortDraft(null)} title="自定义排序" width="sm"
        description={sortDraft ? `排序区域：${rangeRef(sortDraft.target)}` : ''}
        footer={<><Button onClick={() => setSortDraft(null)}>取消</Button><Button variant="primary" onClick={() => {
          if (sortDraft) sortByActive(sortDraft.ascending, sortDraft.column, { ...sortDraft.target, r1: sortDraft.target.r1 + (sortDraft.header ? 1 : 0) });
        }}>确定</Button></>}>
        {sortDraft ? <div className="space-y-3">
          <label className="flex items-center gap-2 text-[12px]"><input type="checkbox" checked={sortDraft.header} onChange={(event) => setSortDraft({ ...sortDraft, header: event.target.checked })} />数据包含标题</label>
          <Field label="排序依据"><Select aria-label="排序依据" value={sortDraft.column} onChange={(event) => setSortDraft({ ...sortDraft, column: Number(event.target.value) })}>
            {Array.from({ length: sortDraft.target.c2 - sortDraft.target.c1 + 1 }, (_, i) => i + sortDraft.target.c1).map((col) => <option key={col} value={col}>{sortDraft.header ? cellDisplay(workbook, sheet, sortDraft.target.r1, col) || `${columnLabel(col)} 列` : `${columnLabel(col)} 列`}</option>)}
          </Select></Field>
          <Field label="次序"><Select aria-label="排序次序" value={String(sortDraft.ascending)} onChange={(event) => setSortDraft({ ...sortDraft, ascending: event.target.value === 'true' })}><option value="true">升序（A → Z，小 → 大）</option><option value="false">降序（Z → A，大 → 小）</option></Select></Field>
        </div> : null}
      </Dialog>

      {/* 功能：删除工作表单独确认，重命名在标签内完成；时间：2026-10-24；作者：lq。 */}
      <Dialog open={sheetDialog !== null} onClose={() => setSheetDialog(null)} title="删除工作表" width="sm"
        footer={<><Button onClick={() => setSheetDialog(null)}>取消</Button><Button variant="danger" onClick={confirmSheetDialog}>删除</Button></>}>
        <p className="text-[12px]">确定删除工作表“{sheetDialog?.name}”吗？此操作可通过撤销恢复。</p>
      </Dialog>

      <Dialog
        open={exportState.open}
        onClose={() => {
          if (!exportState.busy) setExportState((state) => ({ ...state, open: false }));
        }}
        title="导出为 PDF"
        description="将当前工作表的数据区域按 A4 横向分页导出，保留字体、颜色、对齐与边框。"
        width="sm"
        footer={
          <>
            <Button onClick={() => setExportState((state) => ({ ...state, open: false }))} disabled={exportState.busy}>
              取消
            </Button>
            <Button variant="primary" loading={exportState.busy} onClick={() => void runExport()}>
              开始导出
            </Button>
          </>
        }
      >
        <div className="space-y-3">
          <Field label="文件名">
            <Input
              value={exportState.name}
              onChange={(event) => setExportState((state) => ({ ...state, name: event.target.value }))}
              disabled={exportState.busy}
            />
          </Field>
          <div className="rounded-md bg-[var(--mo-surface-2)] px-3 py-2 text-[11px] leading-5 text-[var(--mo-text-muted)]">
            工作表：{sheet.name} · 数据区域 {rangeRef(usedRange(sheet))} · 纸张 A4 {(sheet.pageSetup ?? defaultSheetPageSetup()).orientation === 'portrait' ? '纵向' : '横向'}
          </div>
          {exportState.progress ? (
            <div className="space-y-1.5">
              <ProgressBar
                value={
                  exportState.progress.total > 0
                    ? exportState.progress.current / exportState.progress.total
                    : 0
                }
              />
              <div className="flex justify-between text-[10px] text-[var(--mo-text-muted)]">
                <span>{exportState.progress.message}</span>
                <span>
                  {exportState.progress.current}/{exportState.progress.total}
                </span>
              </div>
            </div>
          ) : null}
          {exportState.busy ? (
            <Button
              size="sm"
              variant="ghost"
              className="text-[var(--mo-danger)]"
              onClick={() => {
                cancelExportRef.current = true;
              }}
            >
              取消导出
            </Button>
          ) : null}
        </div>
      </Dialog>
    </div>
  );
}

function formatNumber(value: number): string {
  if (!Number.isFinite(value)) return '—';
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(4)));
}

function replaceCaseInsensitive(text: string, needle: string, replacement: string): string {
  if (!needle) return text;
  const lower = text.toLowerCase();
  const target = needle.toLowerCase();
  let result = '';
  let index = 0;
  while (index < text.length) {
    const found = lower.indexOf(target, index);
    if (found < 0) {
      result += text.slice(index);
      break;
    }
    result += text.slice(index, found) + replacement;
    index = found + needle.length;
  }
  return result;
}

function createReplacedCell(value: string, existing: { styleId?: string }) {
  return { value, styleId: existing.styleId };
}

function clipboardToTsvText(
  workbook: WorkbookModel,
  sheet: SheetModel,
  clipboard: SheetClipboard,
): string {
  const rows: string[] = [];
  for (let row = clipboard.range.r1; row <= clipboard.range.r2; row += 1) {
    const fields: string[] = [];
    for (let col = clipboard.range.c1; col <= clipboard.range.c2; col += 1) {
      fields.push(cellDisplay(workbook, sheet, row, col));
    }
    rows.push(fields.join('\t'));
  }
  return rows.join('\n');
}

function SheetFindBar(props: {
  query: string;
  replaceText: string;
  caseSensitive: boolean;
  count: number;
  active: number;
  readonly: boolean;
  onQuery: (value: string) => void;
  onReplace: (value: string) => void;
  onCaseSensitive: (value: boolean) => void;
  onPrev: () => void;
  onNext: () => void;
  onReplaceCurrent: () => void;
  onReplaceAll: () => void;
  onClose: () => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  useEffect(() => {
    inputRef.current?.focus();
  }, []);
  return (
    <div className="mo-pop-in absolute right-4 top-3 z-30 w-[330px] rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface)] p-2.5 shadow-pop">
      <div className="flex items-center gap-1.5">
        <div className="relative flex-1">
          <Search size={13} className="pointer-events-none absolute left-2 top-1/2 -translate-y-1/2 text-[var(--mo-text-faint)]" />
          <Input
            ref={inputRef}
            value={props.query}
            onChange={(event) => props.onQuery(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.preventDefault();
                if (event.shiftKey) props.onPrev();
                else props.onNext();
              }
              if (event.key === 'Escape') props.onClose();
            }}
            placeholder="查找内容"
            className="h-7 pl-7 text-[12px]"
          />
        </div>
        <span className="w-14 text-center text-[10px] tabular-nums text-[var(--mo-text-muted)]">
          {props.count > 0 ? `${props.active + 1}/${props.count}` : '0/0'}
        </span>
        <IconButton label="上一个" onClick={props.onPrev} disabled={props.count === 0}>
          <ChevronUp size={14} />
        </IconButton>
        <IconButton label="下一个" onClick={props.onNext} disabled={props.count === 0}>
          <ChevronDown size={14} />
        </IconButton>
        <IconButton label="关闭" onClick={props.onClose}>
          <X size={14} />
        </IconButton>
      </div>
      {!props.readonly ? (
        <div className="mt-2 flex items-center gap-1.5">
          <Input
            value={props.replaceText}
            onChange={(event) => props.onReplace(event.target.value)}
            placeholder="替换为"
            className="h-7 flex-1 text-[12px]"
          />
          <Button size="sm" onClick={props.onReplaceCurrent} disabled={props.count === 0}>
            替换
          </Button>
          <Button size="sm" onClick={props.onReplaceAll} disabled={props.count === 0}>
            全部
          </Button>
        </div>
      ) : null}
      <label className="mt-2 flex cursor-pointer items-center gap-1.5 text-[11px] text-[var(--mo-text-muted)]">
        <input
          type="checkbox"
          checked={props.caseSensitive}
          onChange={(event) => props.onCaseSensitive(event.target.checked)}
          className="h-3 w-3 accent-[var(--mo-accent)]"
        />
        区分大小写
      </label>
    </div>
  );
}

function SheetPropertiesPanel({
  workbook,
  sheet,
  range,
  active,
  session,
  onStyle,
  onClose,
  onShowCompatibility,
}: {
  workbook: WorkbookModel;
  sheet: SheetModel;
  range: CellRange;
  active: { row: number; col: number };
  session: SheetSession;
  onStyle: (patch: Partial<CellStyle>) => void;
  onClose: () => void;
  onShowCompatibility: () => void;
}) {
  const cell = cellAt(sheet, active.row, active.col);
  const style = cellStyle(workbook, cell);
  return (
    <aside className="mo-glass mo-no-print flex w-[268px] shrink-0 flex-col overflow-y-auto border-l border-[var(--mo-line)] bg-[var(--mo-surface)]">
      <div className="flex h-10 items-center justify-between border-b border-[var(--mo-line)] px-3">
        <span className="text-[12px] font-semibold">单元格属性</span>
        <IconButton label="关闭属性面板" onClick={onClose}>
          <X size={14} />
        </IconButton>
      </div>
      <div className="space-y-4 p-3 text-[11px]">
        <section className="space-y-1.5">
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">地址</span>
            <span className="font-mono">
              {columnLabel(active.col)}
              {active.row + 1}
            </span>
          </div>
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">选区</span>
            <span className="font-mono">{rangeRef(range)}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">内容</span>
            <span className="max-w-[150px] truncate">{cellRawInput(cell) || '（空）'}</span>
          </div>
        </section>
        <section className="space-y-2 border-t border-[var(--mo-line)] pt-3">
          <h3 className="font-semibold text-[var(--mo-text-muted)]">字体</h3>
          <Field label="字体">
            <Select value={style.fontFamily ?? workbook.defaultFont} onChange={(event) => onStyle({ fontFamily: event.target.value })}>
              {[workbook.defaultFont, '宋体', '微软雅黑', '等线', '思源黑体', 'Arial', 'Times New Roman'].filter(
                (value, index, list) => list.indexOf(value) === index,
              ).map((font) => (
                <option key={font} value={font}>
                  {font}
                </option>
              ))}
            </Select>
          </Field>
          <div className="grid grid-cols-2 gap-2">
            <Field label="字号">
              <Input
                type="number"
                value={style.fontSizePt ?? workbook.defaultFontSizePt}
                onChange={(event) => onStyle({ fontSizePt: Number(event.target.value) || 11 })}
              />
            </Field>
            <Field label="水平对齐">
              <Select
                value={style.horizontalAlign ?? 'general'}
                onChange={(event) => onStyle({ horizontalAlign: event.target.value as CellStyle['horizontalAlign'] })}
              >
                <option value="general">常规</option>
                <option value="left">左对齐</option>
                <option value="center">居中</option>
                <option value="right">右对齐</option>
              </Select>
            </Field>
          </div>
        </section>
        <section className="space-y-2 border-t border-[var(--mo-line)] pt-3">
          <h3 className="font-semibold text-[var(--mo-text-muted)]">数字格式</h3>
          <Input
            value={style.numFmt ?? 'General'}
            onChange={(event) => onStyle({ numFmt: event.target.value })}
          />
        </section>
        <section className="space-y-1.5 border-t border-[var(--mo-line)] pt-3">
          <h3 className="font-semibold text-[var(--mo-text-muted)]">工作表</h3>
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">名称</span>
            <span>{sheet.name}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">有效单元格</span>
            <span>{Object.keys(sheet.cells).length}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">冻结</span>
            <span>
              {sheet.freeze
                ? `${sheet.freeze.rows} 行 / ${sheet.freeze.cols} 列`
                : '未冻结'}
            </span>
          </div>
          <div className="flex justify-between">
            <span className="text-[var(--mo-text-muted)]">文件大小</span>
            <span>{formatBytes(session.sourceSize)}</span>
          </div>
        </section>
        <button
          type="button"
          onClick={onShowCompatibility}
          className="flex w-full items-center gap-2 rounded-md bg-[var(--mo-surface-2)] px-2 py-2 text-left text-[11px] hover:bg-[var(--mo-surface-3)]"
        >
          <FileWarning size={14} className="text-[var(--mo-warning)]" />
          <span className="flex-1">查看兼容性报告</span>
          <span className="text-[var(--mo-text-faint)]">{session.capability.warnings.length}</span>
        </button>
      </div>
    </aside>
  );
}

// 功能：打印分页实现迁移到 print.ts，此处仅保留导出以兼容既有调用；时间：2026-10-29；作者：lq。
export { buildSheetPrintPages } from './print';
