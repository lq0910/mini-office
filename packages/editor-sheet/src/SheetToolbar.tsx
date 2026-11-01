import { useRef } from 'react';
import type { BorderSide, CellRange, CellStyle, SheetModel, WorkbookModel } from '@mini-office/document-model';
import { columnLabel } from '@mini-office/shared';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDownAZ,
  ArrowUpAZ,
  Bold,
  ChevronDown,
  Columns3,
  Download,
  FileWarning,
  Italic,
  ImagePlus,
  ListFilter,
  Merge,
  PanelRight,
  PaintBucket,
  Printer,
  Redo2,
  Rows3,
  Search,
  Settings2,
  Snowflake,
  Strikethrough,
  Trash2,
  Underline,
  Undo2,
  WrapText,
  Crosshair,
  Sigma,
} from 'lucide-react';
import {
  ColorField,
  Dropdown,
  IconButton,
  Input,
  Select,
  Tooltip,
  type MenuItem,
} from '@mini-office/ui';
import { FONT_CHOICES } from '@mini-office/design-system';
import { cellAt, cellStyle } from './sheet-utils';
// 功能：在顶部公式栏复用函数智能提示；时间：2026-10-24；作者：lq。
import { FormulaInput } from './FormulaInput';
// 功能：接入完整边框菜单及预设类型；时间：2026-10-24；作者：lq。
import { BorderMenu } from './BorderMenu';
import type { BorderPreset } from './sheet-utils';
export type { BorderPreset } from './sheet-utils';

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];

export interface NumberFormatOption {
  label: string;
  format: string;
}

export const NUMBER_FORMATS: NumberFormatOption[] = [
  { label: '常规', format: 'General' },
  { label: '数值 0.00', format: '0.00' },
  { label: '千分位 #,##0.00', format: '#,##0.00' },
  { label: '整数 #,##0', format: '#,##0' },
  { label: '货币 ¥#,##0.00', format: '¥#,##0.00' },
  { label: '百分比 0.00%', format: '0.00%' },
  { label: '日期 yyyy-MM-dd', format: 'yyyy-MM-dd' },
  { label: '时间 hh:mm:ss', format: 'hh:mm:ss' },
  { label: '文本 @', format: '@' },
];

export type BorderWeight = Exclude<BorderSide['style'], 'none'>;

export interface SheetToolbarProps {
  workbook: WorkbookModel;
  sheet: SheetModel;
  range: CellRange;
  active: { row: number; col: number };
  formulaText: string;
  readonly: boolean;
  canUndo: boolean;
  canRedo: boolean;
  propertiesOpen: boolean;
  onFormulaChange: (value: string) => void;
  onFormulaCommit: () => void;
  onFormulaCancel: () => void;
  onFormulaFocus: () => void;
  onAutoSum: () => void;
  onHideRow: () => void;
  onShowRows: () => void;
  onHideColumn: () => void;
  onShowColumns: () => void;
  onJump: (ref: string) => void;
  onStyle: (patch: Partial<CellStyle>) => void;
  onNumberFormat: (format: string) => void;
  onBorders: (preset: BorderPreset) => void;
  borderWeight: BorderWeight;
  onBorderWeight: (weight: BorderWeight) => void;
  borderColor: string;
  onBorderColor: (color: string) => void;
  onMoreBorders: () => void;
  onMergeToggle: (mode?: 'center' | 'merge' | 'unmerge') => void;
  onFreeze: () => void;
  onInsertRow: () => void;
  onDeleteRow: () => void;
  onInsertColumn: () => void;
  onDeleteColumn: () => void;
  onSort: (ascending: boolean) => void;
  onFilter: () => void;
  // 功能：提供常驻插入图片按钮；时间：2026-10-24；作者：lq。
  onInsertImage: () => void;
  // 功能：提供独立排序筛选菜单及自定义排序入口；时间：2026-10-24；作者：lq。
  onCustomSort: () => void;
  onClearFilter: () => void;
  onReapplyFilter: () => void;
  onFind: () => void;
  onExport: () => void;
  onPrint: () => void;
  onPageSetup: () => void;
  onCompatibility: () => void;
  onToggleProperties: () => void;
  /** 功能：阅读模式（十字高亮行列）开关状态与切换入口；时间：2026-10-30；作者：lq。 */
  crosshair: boolean;
  onToggleCrosshair: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onAddSheet: () => void;
}

export function SheetToolbar(props: SheetToolbarProps) {
  const { workbook, sheet, range, active, readonly } = props;
  const formulaRef = useRef<HTMLInputElement>(null);
  const style = cellStyle(workbook, cellAt(sheet, active.row, active.col));
  const rangeLabel =
    range.r1 === range.r2 && range.c1 === range.c2
      ? `${columnLabel(range.c1)}${range.r1 + 1}`
      : `${columnLabel(range.c1)}${range.r1 + 1}:${columnLabel(range.c2)}${range.r2 + 1}`;
  const merged = sheet.merges.some(
    (merge) =>
      merge.r1 <= range.r2 && merge.r2 >= range.r1 && merge.c1 <= range.c2 && merge.c2 >= range.c1,
  );
  const frozen = !!sheet.freeze && (sheet.freeze.rows > 0 || sheet.freeze.cols > 0);

  const dataMenu: MenuItem[] = [
    { type: 'label', label: '行' },
    { id: 'insert-row', label: '在上方插入行', icon: <Rows3 size={13} />, disabled: readonly, onSelect: props.onInsertRow },
    { id: 'delete-row', label: '删除当前行', icon: <Trash2 size={13} />, disabled: readonly, onSelect: props.onDeleteRow },
    { id: 'hide-row', label: '隐藏行', disabled: readonly, onSelect: props.onHideRow },
    { id: 'show-rows', label: '取消隐藏行', disabled: readonly, onSelect: props.onShowRows },
    { type: 'separator' },
    { type: 'label', label: '列' },
    { id: 'insert-col', label: '在左侧插入列', icon: <Columns3 size={13} />, disabled: readonly, onSelect: props.onInsertColumn },
    { id: 'delete-col', label: '删除当前列', icon: <Trash2 size={13} />, disabled: readonly, onSelect: props.onDeleteColumn },
    { id: 'hide-col', label: '隐藏列', disabled: readonly, onSelect: props.onHideColumn },
    { id: 'show-cols', label: '取消隐藏列', disabled: readonly, onSelect: props.onShowColumns },
  ];

  // 功能：参照 Excel 将排序与筛选集中到独立下拉菜单；时间：2026-10-24；作者：lq。
  const sortFilterMenu: MenuItem[] = [
    { id: 'asc', label: '升序', icon: <ArrowUpAZ size={14} />, disabled: readonly, onSelect: () => props.onSort(true) },
    { id: 'desc', label: '降序', icon: <ArrowDownAZ size={14} />, disabled: readonly, onSelect: () => props.onSort(false) },
    { id: 'custom', label: '自定义排序…', disabled: readonly, onSelect: props.onCustomSort },
    { type: 'separator' },
    { id: 'filter', label: sheet.autoFilter ? '关闭筛选' : '筛选', icon: <ListFilter size={14} />, disabled: readonly, onSelect: props.onFilter },
    { id: 'clear', label: '清除筛选条件', disabled: readonly || !sheet.autoFilter, onSelect: props.onClearFilter },
    { id: 'reapply', label: '重新应用', disabled: readonly || !sheet.autoFilter, onSelect: props.onReapplyFilter },
  ];

  const formatValue = style.numFmt ?? 'General';

  // 功能：fx 入口展开常用函数清单，而不是直接写入 SUM；时间：2026-11-01；作者：lq。
  const insertFunction = (name: string) => {
    const input = formulaRef.current;
    const text = props.formulaText;
    if (text.startsWith('=')) {
      const cursor = input?.selectionStart ?? text.length;
      props.onFormulaChange(`${text.slice(0, cursor)}${name}(${text.slice(cursor)}`);
      const caret = cursor + name.length + 1;
      requestAnimationFrame(() => { input?.focus(); input?.setSelectionRange(caret, caret); });
      return;
    }
    props.onFormulaChange(`=${name}(`);
    const caret = name.length + 2;
    requestAnimationFrame(() => { input?.focus(); input?.setSelectionRange(caret, caret); });
  };

  const functionMenu: MenuItem[] = [
    { type: 'label', label: '数学与统计' },
    ...([
      ['SUM', '求和'], ['AVERAGE', '平均值'], ['COUNT', '计数'], ['MAX', '最大值'], ['MIN', '最小值'],
      ['ROUND', '四舍五入'], ['ABS', '绝对值'], ['SUBTOTAL', '分类汇总'],
    ] as Array<[string, string]>).map(([name, description]) => ({
      id: `fn-${name}`,
      label: `${name}　${description}`,
      disabled: readonly,
      onSelect: () => insertFunction(name),
    })),
    { type: 'separator' },
    { type: 'label', label: '逻辑与查找' },
    ...([
      ['IF', '条件判断'], ['IFERROR', '错误处理'], ['AND', '同时成立'], ['OR', '任一成立'],
      ['VLOOKUP', '纵向查找'], ['XLOOKUP', '查找匹配'],
    ] as Array<[string, string]>).map(([name, description]) => ({
      id: `fn-${name}`,
      label: `${name}　${description}`,
      disabled: readonly,
      onSelect: () => insertFunction(name),
    })),
    { type: 'separator' },
    { type: 'label', label: '文本与日期' },
    ...([
      ['CONCAT', '连接文本'], ['LEFT', '左侧截取'], ['RIGHT', '右侧截取'], ['LEN', '文本长度'],
      ['TEXT', '格式化文本'], ['TODAY', '今天日期'], ['NOW', '当前时间'],
    ] as Array<[string, string]>).map(([name, description]) => ({
      id: `fn-${name}`,
      label: `${name}　${description}`,
      disabled: readonly,
      onSelect: () => insertFunction(name),
    })),
  ];

  const mergeMenu: MenuItem[] = [
    { id: 'merge-center', label: '合并后居中', disabled: readonly, onSelect: () => props.onMergeToggle('center') },
    { id: 'merge-only', label: '合并单元格', disabled: readonly, onSelect: () => props.onMergeToggle('merge') },
    { id: 'merge-unmerge', label: '取消合并', disabled: readonly, onSelect: () => props.onMergeToggle('unmerge') },
  ];


  return (
    <div data-sheet-toolbar className="mo-glass mo-no-print shrink-0 border-b border-[var(--mo-line)] bg-[var(--mo-surface)]">
      {/* 功能：拉开 Excel 工具栏控件间距，降低图标拥挤并保持窄窗口可横向滚动；时间：2026-10-24；作者：lq。 */}
      <div
        className="flex h-11 items-center gap-1 overflow-x-auto px-2"
        onMouseDown={(event) => {
          if ((event.target as HTMLElement).closest('button')) event.preventDefault();
        }}
      >
        <Tooltip label="撤销 ⌘Z">
          <IconButton label="撤销" disabled={!props.canUndo || readonly} onClick={props.onUndo}>
            <Undo2 size={15} />
          </IconButton>
        </Tooltip>
        <Tooltip label="重做 ⇧⌘Z">
          <IconButton label="重做" disabled={!props.canRedo || readonly} onClick={props.onRedo}>
            <Redo2 size={15} />
          </IconButton>
        </Tooltip>
        <span className="mx-1 h-4 w-px bg-[var(--mo-line)]" />
        <Select
          aria-label="字体"
          className="h-7 w-[120px] shrink-0 border-transparent bg-transparent text-[12px]"
          value={style.fontFamily ?? workbook.defaultFont}
          disabled={readonly}
          onChange={(event) => props.onStyle({ fontFamily: event.target.value })}
        >
          {[workbook.defaultFont, ...FONT_CHOICES.filter((font) => font !== '默认')].filter(
            (font, index, list) => list.indexOf(font) === index,
          ).map((font) => (
            <option key={font} value={font}>
              {font}
            </option>
          ))}
        </Select>
        <Select
          aria-label="字号"
          className="h-7 w-[58px] shrink-0 border-transparent bg-transparent text-[12px]"
          value={String(style.fontSizePt ?? workbook.defaultFontSizePt)}
          disabled={readonly}
          onChange={(event) => props.onStyle({ fontSizePt: Number(event.target.value) })}
        >
          {[...new Set([style.fontSizePt ?? workbook.defaultFontSizePt, ...FONT_SIZES])]
            .sort((a, b) => a - b)
            .map((size) => (
              <option key={size} value={size}>
                {size}
              </option>
            ))}
        </Select>
        <span className="mx-1 h-4 w-px bg-[var(--mo-line)]" />
        <IconButton label="加粗" active={!!style.bold} disabled={readonly} onClick={() => props.onStyle({ bold: !style.bold })}>
          <Bold size={15} />
        </IconButton>
        <IconButton label="斜体" active={!!style.italic} disabled={readonly} onClick={() => props.onStyle({ italic: !style.italic })}>
          <Italic size={15} />
        </IconButton>
        <IconButton label="下划线" active={!!style.underline} disabled={readonly} onClick={() => props.onStyle({ underline: !style.underline })}>
          <Underline size={15} />
        </IconButton>
        <IconButton label="删除线" active={!!style.strike} disabled={readonly} onClick={() => props.onStyle({ strike: !style.strike })}>
          <Strikethrough size={15} />
        </IconButton>
        {/* 功能：使用 Excel 风格的“A + 颜色下划线”工具栏图标，提升颜色按钮辨识度；时间：2026-10-31；作者：lq。 */}
        <ColorField
          label="字体颜色"
          value={style.color}
          icon={<span className="text-[16px] font-semibold leading-none">A</span>}
          indicatorColor={style.color ?? '#1D2939'}
          onChange={(color) => props.onStyle({ color })}
        />
        {/* 功能：使用 Excel 风格的油漆桶图标并保留当前填充色下划线；时间：2026-10-31；作者：lq。 */}
        <ColorField
          label="填充颜色"
          value={style.fill}
          icon={<PaintBucket size={15} strokeWidth={2.1} />}
          indicatorColor={style.fill ?? '#F2C94C'}
          onChange={(fill) => props.onStyle({ fill })}
        />
        {/* 功能：常用边框、线型和颜色集中到分组菜单，左侧按钮重复应用上次预设；时间：2026-10-24；作者：lq。 */}
        <BorderMenu readonly={readonly} weight={props.borderWeight} color={props.borderColor} onWeight={props.onBorderWeight} onColor={props.onBorderColor} onApply={props.onBorders} onMore={props.onMoreBorders} />
        <span className="mx-1 h-4 w-px bg-[var(--mo-line)]" />
        <IconButton label="左对齐" active={style.horizontalAlign === 'left'} disabled={readonly} onClick={() => props.onStyle({ horizontalAlign: 'left' })}>
          <AlignLeft size={15} />
        </IconButton>
        <IconButton label="居中" active={style.horizontalAlign === 'center'} disabled={readonly} onClick={() => props.onStyle({ horizontalAlign: 'center' })}>
          <AlignCenter size={15} />
        </IconButton>
        <IconButton label="右对齐" active={style.horizontalAlign === 'right'} disabled={readonly} onClick={() => props.onStyle({ horizontalAlign: 'right' })}>
          <AlignRight size={15} />
        </IconButton>
        <IconButton label="自动换行" active={!!style.wrapText} disabled={readonly} onClick={() => props.onStyle({ wrapText: !style.wrapText })}>
          <WrapText size={15} />
        </IconButton>
        <Select
          aria-label="数字格式"
          className="h-7 w-[118px] shrink-0 border-transparent bg-transparent text-[11px]"
          value={formatValue}
          disabled={readonly}
          onChange={(event) => props.onNumberFormat(event.target.value)}
        >
          {[...NUMBER_FORMATS, ...(NUMBER_FORMATS.some((item) => item.format === formatValue) ? [] : [{ label: formatValue, format: formatValue }])].map(
            (item) => (
              <option key={item.format} value={item.format}>
                {item.label}
              </option>
            ),
          )}
        </Select>
        <span className="mx-1 h-4 w-px bg-[var(--mo-line)]" />
        {/* 功能：合并按钮拆成“合并后居中 + 下拉选项”，图标与文字同时可见，避免图标含义不明；时间：2026-11-01；作者：lq。 */}
        <div className="flex shrink-0 items-stretch">
          <button
            type="button"
            aria-label={merged ? '取消合并' : '合并后居中'}
            title={merged ? '取消合并单元格' : '合并后居中：合并选区并把内容居中'}
            disabled={readonly}
            onClick={() => props.onMergeToggle(merged ? 'unmerge' : 'center')}
            className={`flex h-7 items-center gap-1 rounded-l-md border px-2 text-[11px] transition-colors disabled:opacity-40 ${
              merged
                ? 'border-[var(--mo-accent)] bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]'
                : 'border-[var(--mo-line)] text-[var(--mo-text-secondary)] hover:bg-[var(--mo-surface-2)]'
            }`}
          >
            <Merge size={14} />
            {merged ? '取消合并' : '合并居中'}
          </button>
          <Dropdown
            minWidth={168}
            items={mergeMenu}
            trigger={({ toggle }) => (
              <button
                type="button"
                aria-label="合并选项"
                aria-haspopup="menu"
                disabled={readonly}
                onClick={toggle}
                className="flex h-7 w-4 items-center justify-center rounded-r-md border border-l-0 border-[var(--mo-line)] text-[var(--mo-text-secondary)] hover:bg-[var(--mo-surface-2)] disabled:opacity-40"
              >
                <ChevronDown size={11} />
              </button>
            )}
          />
        </div>
        <IconButton label="冻结窗格" active={frozen} onClick={props.onFreeze}>
          <Snowflake size={15} />
        </IconButton>
        <Dropdown
          minWidth={210}
          items={dataMenu}
          trigger={({ toggle }) => (
            <IconButton label="数据与行列" onClick={toggle}>
              <Rows3 size={15} />
            </IconButton>
          )}
        />
        {/* 功能：独立显示排序筛选按钮，避免混入行列结构操作；时间：2026-10-24；作者：lq。 */}
        <Dropdown minWidth={190} items={sortFilterMenu} trigger={({ toggle }) => (
          <IconButton label="排序和筛选" active={!!sheet.autoFilter} onClick={toggle}><ListFilter size={15} /></IconButton>
        )} />
        {/* 功能：插入本地图片到当前单元格位置；时间：2026-10-24；作者：lq。 */}
        <IconButton label="插入图片" disabled={readonly} onClick={props.onInsertImage}><ImagePlus size={15} /></IconButton>
        <Tooltip label="查找替换 ⌘F">
          <IconButton label="查找" onClick={props.onFind}>
            <Search size={15} />
          </IconButton>
        </Tooltip>
        {/* 功能：阅读模式开关，开启后当前行列在表头与网格中高亮，方便核对数据；时间：2026-10-30；作者：lq。 */}
        <Tooltip label={props.crosshair ? '关闭阅读模式' : '阅读模式：高亮当前行列'}>
          <IconButton label="阅读模式" active={props.crosshair} onClick={props.onToggleCrosshair}>
            <Crosshair size={15} />
          </IconButton>
        </Tooltip>
        <IconButton label="属性面板" active={props.propertiesOpen} onClick={props.onToggleProperties}>
          <PanelRight size={15} />
        </IconButton>
        {/* 功能：缩放已移到右下角状态栏（与 Word 一致），工具栏只保留输出类按钮；时间：2026-10-31；作者：lq。 */}
        <div className="flex shrink-0 items-center gap-1.5">
          <IconButton label="打印" onClick={props.onPrint}>
            <Printer size={15} />
          </IconButton>
          {/* 功能：放大页面设置按钮的可视图标与点击区域，避免齿轮字符过小；时间：2026-10-24；作者：lq。 */}
          <IconButton
            label="页面设置"
            onClick={props.onPageSetup}
            className="h-8 w-8 text-[var(--mo-text-primary)]"
          >
            <Settings2 size={17} strokeWidth={2.1} />
          </IconButton>
          <IconButton label="导出 PDF" onClick={props.onExport}>
            <Download size={15} />
          </IconButton>
          <IconButton label="兼容性报告" onClick={props.onCompatibility}>
            <FileWarning size={15} />
          </IconButton>
        </div>
      </div>

      <div className="flex h-9 items-center gap-2 border-t border-[var(--mo-line)] px-2">
        <Input
          aria-label="名称框"
          className="h-6 w-[92px] shrink-0 text-center font-mono text-[11px]"
          defaultValue={rangeLabel}
          key={`${range.c1}-${range.r1}-${range.c2}-${range.r2}`}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              props.onJump(event.currentTarget.value);
              event.currentTarget.blur();
            }
          }}
        />
        <Dropdown
          minWidth={228}
          items={functionMenu}
          trigger={({ toggle }) => (
            <button
              type="button"
              aria-label="插入函数"
              title="插入函数"
              aria-haspopup="menu"
              disabled={readonly}
              className="flex h-7 shrink-0 items-center gap-0.5 rounded px-1 font-serif text-[12px] italic text-[var(--mo-text-muted)] hover:bg-[var(--mo-accent-soft)] disabled:opacity-50"
              onMouseDown={(event) => event.preventDefault()}
              onClick={toggle}
            >
              fx
              <ChevronDown size={10} />
            </button>
          )}
        />
        <IconButton label="自动求和" disabled={readonly} onClick={props.onAutoSum}><Sigma size={14} /></IconButton>
        <FormulaInput
          ref={formulaRef}
          aria-label="公式栏"
          className="h-6 min-w-0 flex-1 font-mono text-[12px]"
          value={props.formulaText}
          disabled={readonly}
          onValueChange={props.onFormulaChange}
          onFocus={props.onFormulaFocus}
          onBlur={(event) => { if (!(event.relatedTarget as HTMLElement | null)?.hasAttribute('data-sheet-formula-input')) props.onFormulaCommit(); }}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              props.onFormulaCommit();
            } else if (event.key === 'Escape') {
              event.preventDefault(); props.onFormulaCancel();
            }
          }}
          placeholder="输入值或公式，例如 =SUM(A1:A3)"
        />
      </div>
    </div>
  );
}
