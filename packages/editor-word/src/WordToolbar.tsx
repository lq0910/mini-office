import { useEffect, useId, useMemo, useState, type ReactNode } from 'react';
import { useWorkspaceStore, type DocParagraph, type DocModel, type ParagraphStyle, type RunStyle } from '@mini-office/document-model';
import { availableFontChoices } from '@mini-office/design-system';
import { MOD_KEY } from '@mini-office/shared';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  ChevronDown,
  ChevronUp,
  Download,
  FileWarning,
  ImagePlus,
  IndentDecrease,
  IndentIncrease,
  Italic,
  List,
  MoreHorizontal,
  Printer,
  Redo2,
  Save,
  SaveAll,
  Search,
  Strikethrough,
  Underline,
  Undo2,
  PanelRight,
  Scissors,
  Highlighter,
  Baseline,
  PanelTop,
  Mails,
  Paintbrush,
  Pilcrow,
  Ruler,
  TableOfContents,
} from 'lucide-react';
import { ColorField, Dropdown, IconButton, Select, Spinner, Tooltip } from '@mini-office/ui';
import type { MenuItem } from '@mini-office/ui';
import { TablePicker } from './TablePicker';
import { FontSizePicker } from './FontSizePicker';
import { LineSpacingPicker } from './LineSpacingPicker';
import { charWidthPx } from './typography';
import { StylePicker, type StylePickerProps } from './StylePicker';

export interface WordToolbarProps {
  /** 功能：工具栏只读取字号摘要，不要求把完整新文档转换为旧模型；时间：2026-11-09；作者：lq。 */
  doc: Pick<DocModel, 'defaultFontSizePt'>;
  /** 功能：M2 开发入口保留原有 UI，M3 尚未接线的操作明确禁用；时间：2026-11-09；作者：lq。 */
  migrationPending?: boolean;
  /** 功能：按功能开放已迁移入口，不因部分迁移完成而启用空回调；时间：2026-11-10；作者：lq。 */
  // 功能：图片和表格插入按活动核心 schema 逐项开放；时间：2026-11-11；作者：lq。
  migrated?: ReadonlyArray<'font' | 'paragraph' | 'pageBreak' | 'find' | 'list' | 'formatPainter' | 'pageLayout' | 'headerFooter' | 'ruler' | 'toc' | 'image' | 'table' | 'mailMerge' | 'print' | 'export'>;
  /** 功能：列表类别和零起点层级来自核心选区摘要，混合选区不伪装统一列表；时间：2026-11-10；作者：lq。 */
  listKind?: 'bullet' | 'ordered';
  listLevel?: number;
  listMixed?: boolean;
  /** 功能：列表缩进按钮改变真实层级，普通段落继续调整长度；时间：2026-11-10；作者：lq。 */
  onChangeListLevel?: (delta: number) => void;
  /** 功能：新核心直接提供格式摘要和混合字段，不创建旧文档/段落对象；时间：2026-11-10；作者：lq。 */
  paragraphStyle?: ParagraphStyle;
  mixedText?: ReadonlyArray<keyof RunStyle>;
  mixedParagraph?: ReadonlyArray<keyof ParagraphStyle>;
  readonly: boolean;
  activeParagraph?: DocParagraph;
  activeStyle: RunStyle;
  canUndo: boolean;
  canRedo: boolean;
  /** 功能：快捷保存与另存为入口及保存中状态；时间：2026-10-30；作者：lq。 */
  saving: boolean;
  onSave: () => void;
  onSaveAs: () => void;
  zoom: number;
  pageCount: number;
  currentPage: number;
  findOpen: boolean;
  propertiesOpen: boolean;
  onInlineCommand: (command: 'bold' | 'italic' | 'underline' | 'strikeThrough') => void;
  onColorFormat: (patch: Partial<RunStyle>, kind: 'color' | 'highlight') => void;
  onFontFormat: (patch: Partial<RunStyle>) => void;
  onParagraphFormat: (patch: Partial<ParagraphStyle>) => void;
  /** 功能：按选择的行列数插入表格；时间：2026-10-19；作者：lq。 */
  onInsertTable: (rows: number, columns: number) => void;
  onInsertImage: () => void;
  onPageBreak: () => void;
  /** 功能：打开页眉页脚设置；时间：2026-10-24；作者：lq。 */
  onHeaderFooter: () => void;
  /** 功能：邮件合并操作入口；时间：2026-10-24；作者：lq。 */
  onMailMerge: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onZoomChange: (zoom: number) => void;
  onFitWidth: () => void;
  onToggleFind: () => void;
  onToggleProperties: () => void;
  onExport: () => void;
  onPrint: () => void;
  onShowCompatibility: () => void;
  onAddNumbering: () => void;
  onAddBullets: () => void;
  /** 功能：格式刷入口，sticky 为双击后的连续套用模式；时间：2026-10-29；作者：lq。 */
  onFormatPainter: (sticky: boolean) => void;
  formatPainterActive: boolean;
  /** 功能：打开段落对话框，精确设置缩进与行距；时间：2026-10-29；作者：lq。 */
  onOpenParagraphDialog: () => void;
  /** 功能：样式库入口的完整配置，由编辑器组装，工具栏只负责摆放位置；时间：2026-10-29；作者：lq。 */
  stylePicker?: StylePickerProps;
  /** 功能：标尺显隐与目录插入入口；时间：2026-10-30；作者：lq。 */
  rulerOpen: boolean;
  onToggleRuler: () => void;
  onInsertToc: () => void;
  /** 功能：新核心原位重建目录，与插入入口分别启用；时间：2026-11-11；作者：lq。 */
  onUpdateToc?: () => void;
  /** 功能：切换编辑区换行符与制表符格式标记；时间：2026-10-31；作者：lq。 */
  showFormattingMarks?: boolean;
  onToggleFormattingMarks?: () => void;
  onBeforeInteraction?: () => void;
}

const RIBBON_TABS = ['开始', '插入', '布局', '引用', '邮件', '视图'] as const;
type RibbonTab = typeof RIBBON_TABS[number];

function RibbonGroup({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div role="group" aria-label={label} className="flex shrink-0 flex-col border-r border-[var(--mo-line)] px-2 last:border-r-0">
      <div className="flex h-[58px] flex-col justify-center gap-0.5">{children}</div>
      <span aria-hidden="true" className="mt-0.5 text-center text-[10px] leading-4 text-[var(--mo-text-muted)]">{label}</span>
    </div>
  );
}

function RibbonButton({ label, icon, disabled, active, onClick }: {
  label: string;
  icon: ReactNode;
  disabled?: boolean;
  active?: boolean;
  onClick?: () => void;
}) {
  return (
    <button
      type="button"
      aria-label={label}
      title={label}
      aria-pressed={active}
      disabled={disabled}
      onClick={onClick}
      className={`flex h-7 shrink-0 items-center gap-1.5 whitespace-nowrap rounded px-2 text-[11px] hover:bg-[var(--mo-surface-2)] disabled:opacity-40 ${
        active ? 'bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]' : 'text-[var(--mo-text-primary)]'
      }`}
    >
      {icon}{label}
    </button>
  );
}

export function WordToolbar(props: WordToolbarProps) {
  const { readonly, activeStyle, activeParagraph } = props;
  const [activeTab, setActiveTab] = useState<RibbonTab>('开始');
  const [collapsed, setCollapsed] = useState(false);
  const ribbonId = useId();
  const selectTab = (tab: RibbonTab) => {
    setActiveTab(tab);
    setCollapsed(false);
  };
  const paragraphStyle = props.paragraphStyle ?? activeParagraph?.style ?? {};
  // 功能：保留旧入口的全功能行为，新核心按已接线的功能逐项解禁；时间：2026-11-10；作者：lq。
  const pending = (feature: NonNullable<WordToolbarProps['migrated']>[number]) => !!props.migrationPending && !props.migrated?.includes(feature);
  // 功能：字体下拉优先使用平台适配器枚举到的真实系统字体，拿不到时才退回逐个探测；时间：2026-10-31；作者：lq。
  const installedFonts = useWorkspaceStore((state) => state.fonts);
  const installedFamilies = useMemo(() => installedFonts.map((font) => font.family), [installedFonts]);
  const [fontChoices, setFontChoices] = useState(() => availableFontChoices(activeStyle.fontFamily, installedFamilies));
  useEffect(() => {
    const refresh = () => setFontChoices(availableFontChoices(activeStyle.fontFamily, installedFamilies));
    refresh();
    document.fonts?.ready.then(refresh).catch(() => {});
  }, [activeStyle.fontFamily, installedFamilies]);
  const saveDisabled = readonly || props.saving;
  const saveShortcut = KbdLabel('S');
  const saveAsShortcut = navigator.platform.toLowerCase().includes('mac') ? '⇧⌘S' : 'Ctrl+Shift+S';
  const saveTooltip = props.saving ? '正在保存…' : readonly ? '只读文档无法保存' : `保存 ${saveShortcut}`;
  const saveItems: MenuItem[] = [
    { id: 'save', label: '保存', icon: <Save size={14} />, shortcut: saveShortcut, disabled: saveDisabled, onSelect: props.onSave },
    { type: 'separator' },
    { id: 'save-as', label: '另存为副本…', icon: <SaveAll size={14} />, shortcut: saveAsShortcut, disabled: saveDisabled, onSelect: props.onSaveAs },
  ];
  const moreItems: MenuItem[] = [
    { id: 'find', disabled: pending('find'), label: '查找替换', icon: <Search size={14} />, shortcut: KbdLabel('F'), onSelect: props.onToggleFind },
    { id: 'layout', disabled: pending('pageLayout'), label: '布局与属性', icon: <PanelRight size={14} />, onSelect: props.onToggleProperties },
    { id: 'mail-merge', label: '邮件合并 / 批量报告', icon: <Mails size={14} />, disabled: readonly || pending('mailMerge'), onSelect: props.onMailMerge },
    { id: 'image', label: '插入图片', icon: <ImagePlus size={14} />, disabled: readonly || pending('image'), onSelect: props.onInsertImage },
    { id: 'header-footer', label: '页眉、页脚与页码', icon: <PanelTop size={14} />, disabled: readonly || pending('headerFooter'), onSelect: props.onHeaderFooter },
    { id: 'page-break', label: '插入分页符', icon: <Scissors size={14} />, disabled: readonly || pending('pageBreak'), onSelect: props.onPageBreak },
    { id: 'print', disabled: pending('print'), label: '打印', icon: <Printer size={14} />, onSelect: props.onPrint },
    { id: 'export', disabled: pending('export'), label: '导出 PDF', icon: <Download size={14} />, onSelect: props.onExport },
    { id: 'toc', label: '插入目录', icon: <TableOfContents size={14} />, disabled: readonly || pending('toc'), onSelect: props.onInsertToc },
    // 功能：紧凑窗口更多菜单也提供目录更新入口；时间：2026-11-11；作者：lq。
    ...(props.onUpdateToc ? [{ id: 'toc-update', label: '更新目录', icon: <TableOfContents size={14} />, disabled: readonly || pending('toc'), onSelect: props.onUpdateToc }] : []),
    { id: 'ruler', disabled: pending('ruler'), label: props.rulerOpen ? '隐藏标尺' : '显示标尺', icon: <Ruler size={14} />, onSelect: props.onToggleRuler },
    { type: 'separator' },
    { id: 'compat', label: '兼容性报告', icon: <FileWarning size={14} />, onSelect: props.onShowCompatibility },
    { id: 'properties', disabled: pending('pageLayout'), label: props.propertiesOpen ? '隐藏属性面板' : '显示属性面板', icon: <PanelRight size={14} />, onSelect: props.onToggleProperties },
  ];

  return (
    <div
      aria-label="Word 功能区"
      className="mo-glass mo-no-print min-w-0 shrink-0 border-b border-[var(--mo-line)] bg-[var(--mo-surface)] text-[var(--mo-text-primary)]"
      onPointerDownCapture={() => props.onBeforeInteraction?.()}
      onMouseDown={(event) => {
        if ((event.target as HTMLElement).closest('button')) event.preventDefault();
      }}
    >
      <div className="flex h-8 items-center gap-0.5 px-1">
        <Dropdown
          className="shrink-0"
          minWidth={236}
          items={saveItems}
          trigger={({ open, toggle }) => (
            <div role="group" aria-label="快捷保存" className="flex shrink-0 items-center">
              <Tooltip label={saveTooltip}>
                <IconButton label="保存" data-tooltip={saveTooltip} aria-busy={props.saving} disabled={saveDisabled}
                  onClick={props.onSave} className="rounded-r-none">
                  {props.saving ? <Spinner size={15} /> : <Save size={15} />}
                </IconButton>
              </Tooltip>
              <Tooltip label="保存选项">
                <IconButton label="保存选项" aria-haspopup="menu" aria-expanded={open} disabled={saveDisabled}
                  onClick={toggle} className="w-4 rounded-l-none">
                  <ChevronDown size={11} />
                </IconButton>
              </Tooltip>
            </div>
          )}
        />
        <Tooltip label={`撤销 ${KbdLabel('Z')}`}>
          <IconButton label="撤销" disabled={!props.canUndo || readonly} onClick={props.onUndo}><Undo2 size={15} /></IconButton>
        </Tooltip>
        <Tooltip label={`重做 ${KbdLabel('Shift+Z')}`}>
          <IconButton label="重做" disabled={!props.canRedo || readonly} onClick={props.onRedo}><Redo2 size={15} /></IconButton>
        </Tooltip>
        <span aria-hidden="true" className="mx-1 h-4 w-px shrink-0 bg-[var(--mo-line)]" />
        <div role="tablist" aria-label="Word 功能选项卡" className="flex min-w-0 flex-1 items-center self-stretch">
          {RIBBON_TABS.map((tab, index) => (
            <button
              key={tab}
              type="button"
              role="tab"
              id={`${ribbonId}-tab-${index}`}
              aria-selected={activeTab === tab}
              aria-controls={`${ribbonId}-panel-${index}`}
              tabIndex={activeTab === tab ? 0 : -1}
              onClick={() => selectTab(tab)}
              onKeyDown={(event) => {
                let next: number;
                if (event.key === 'ArrowRight') next = (index + 1) % RIBBON_TABS.length;
                else if (event.key === 'ArrowLeft') next = (index + RIBBON_TABS.length - 1) % RIBBON_TABS.length;
                else if (event.key === 'Home') next = 0;
                else if (event.key === 'End') next = RIBBON_TABS.length - 1;
                else return;
                event.preventDefault();
                selectTab(RIBBON_TABS[next]);
                document.getElementById(`${ribbonId}-tab-${next}`)?.focus({ preventScroll: true });
              }}
              className={`h-full min-w-0 flex-1 border-b-2 px-1 text-[11px] sm:flex-none sm:px-3 ${
                activeTab === tab
                  ? 'border-[var(--mo-accent)] font-medium text-[var(--mo-accent)]'
                  : 'border-transparent text-[var(--mo-text-muted)] hover:bg-[var(--mo-surface-2)] hover:text-[var(--mo-text-primary)]'
              }`}
            >{tab}</button>
          ))}
        </div>
        <Dropdown align="right" minWidth={236} items={moreItems} trigger={({ open, toggle }) => (
          <IconButton label="更多操作" aria-haspopup="menu" aria-expanded={open} className="shrink-0" onClick={toggle}>
            <MoreHorizontal size={16} />
          </IconButton>
        )} />
        <IconButton label={collapsed ? '展开功能区' : '折叠功能区'} title={collapsed ? '展开功能区' : '折叠功能区'}
          aria-expanded={!collapsed} aria-controls={`${ribbonId}-panel-${RIBBON_TABS.indexOf(activeTab)}`}
          className="shrink-0" onClick={() => setCollapsed((value) => !value)}>
          {collapsed ? <ChevronDown size={14} /> : <ChevronUp size={14} />}
        </IconButton>
      </div>
      {RIBBON_TABS.map((tab, index) => (
        <div key={tab} role="tabpanel" id={`${ribbonId}-panel-${index}`} aria-labelledby={`${ribbonId}-tab-${index}`}
          hidden={collapsed || activeTab !== tab} className="mo-scroll overflow-x-auto border-t border-[var(--mo-line)] py-1">
          {!collapsed && activeTab === tab ? (
            <div className="flex w-max min-w-full items-start px-1">
              {tab === '开始' ? (
                <>
                  <RibbonGroup label="格式刷">
                    <Tooltip label="格式刷：单击复制格式并套用一次，双击可连续套用，Esc 取消">
                      <button type="button" aria-label="格式刷" aria-pressed={props.formatPainterActive} disabled={readonly || pending('formatPainter')}
                        onClick={() => props.onFormatPainter(false)} onDoubleClick={() => props.onFormatPainter(true)}
                        className={`flex h-[58px] flex-col items-center justify-center gap-1 rounded px-2 text-[11px] hover:bg-[var(--mo-surface-2)] disabled:opacity-40 ${
                          props.formatPainterActive ? 'bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]' : ''
                        }`}>
                        <Paintbrush size={20} />格式刷
                      </button>
                    </Tooltip>
                  </RibbonGroup>
                  <RibbonGroup label="字体">
                    <div className="flex items-center gap-1">
                      <Select aria-label="字体" className="h-7 w-[126px] shrink-0 text-[11px]" value={props.mixedText?.includes('fontFamily') ? '' : activeStyle.fontFamily ?? '默认'} disabled={readonly || pending('font')}
                        onChange={(event) => props.onFontFormat({ fontFamily: event.target.value === '默认' ? undefined : event.target.value })}>
                        {/* 功能：混合字体显示空值，不把第一段格式作为整段选区的格式；时间：2026-11-10；作者：lq。 */}
                        {props.mixedText?.includes('fontFamily') ? <option value="">混合字体</option> : null}
                        {fontChoices.map((font) => <option key={font} value={font}>{font}</option>)}
                      </Select>
                      <FontSizePicker value={activeStyle.fontSizePt ?? props.doc.defaultFontSizePt} mixed={props.mixedText?.includes('fontSizePt')} disabled={readonly || pending('font')}
                        onChange={(pt) => props.onFontFormat({ fontSizePt: pt })} />
                    </div>
                    <div className="flex items-center gap-0.5">
                      <Tooltip label={`加粗 ${KbdLabel('B')}`}>
                        <IconButton label="加粗" active={!!activeStyle.bold} disabled={readonly} onClick={() => props.onInlineCommand('bold')}><Bold size={15} /></IconButton>
                      </Tooltip>
                      <Tooltip label={`斜体 ${KbdLabel('I')}`}>
                        <IconButton label="斜体" active={!!activeStyle.italic} disabled={readonly} onClick={() => props.onInlineCommand('italic')}><Italic size={15} /></IconButton>
                      </Tooltip>
                      <Tooltip label={`下划线 ${KbdLabel('U')}`}>
                        <IconButton label="下划线" active={!!activeStyle.underline} disabled={readonly} onClick={() => props.onInlineCommand('underline')}><Underline size={15} /></IconButton>
                      </Tooltip>
                      <IconButton label="删除线" active={!!activeStyle.strike} disabled={readonly} onClick={() => props.onInlineCommand('strikeThrough')}><Strikethrough size={15} /></IconButton>
                      <ColorField label="字体颜色" icon={<Baseline size={16} />} indicatorColor={activeStyle.color ?? 'var(--mo-text-primary)'}
                        disabled={readonly || pending('font')} value={activeStyle.color} onChange={(color) => props.onColorFormat({ color }, 'color')} />
                      <ColorField label="文本背景颜色（突出显示）" icon={<Highlighter size={16} />} indicatorColor={activeStyle.highlight ?? 'transparent'}
                        disabled={readonly || pending('font')} value={activeStyle.highlight} onChange={(color) => props.onColorFormat({ highlight: color }, 'highlight')} />
                    </div>
                  </RibbonGroup>
                  <RibbonGroup label="段落">
                    <div className="flex items-center gap-0.5">
                      <IconButton label="项目符号" aria-pressed={!props.listMixed && props.listKind === 'bullet'} active={!props.listMixed && props.listKind === 'bullet'} title={props.listKind === 'bullet' ? `项目符号：第 ${(props.listLevel ?? 0) + 1} 级` : undefined} disabled={readonly || pending('list')} onClick={props.onAddBullets}><List size={15} /></IconButton>
                      <IconButton label="编号" aria-pressed={!props.listMixed && props.listKind === 'ordered'} active={!props.listMixed && props.listKind === 'ordered'} title={props.listKind === 'ordered' ? `编号：第 ${(props.listLevel ?? 0) + 1} 级` : undefined} disabled={readonly || pending('list')} onClick={props.onAddNumbering}><span className="text-[11px] font-semibold">1.</span></IconButton>
                      <IconButton label="减少缩进" disabled={readonly || pending('paragraph')} onClick={() => props.listKind && props.onChangeListLevel ? props.onChangeListLevel(-1) : props.onParagraphFormat({ indentLeftPx: Math.max(0, (paragraphStyle.indentLeftPx ?? 0) - 24) })}><IndentDecrease size={15} /></IconButton>
                      <IconButton label="增加缩进" disabled={readonly || pending('paragraph')} onClick={() => props.listKind && props.onChangeListLevel ? props.onChangeListLevel(1) : props.onParagraphFormat({ indentLeftPx: (paragraphStyle.indentLeftPx ?? 0) + 24 })}><IndentIncrease size={15} /></IconButton>
                      <Dropdown minWidth={212} items={[
                        { type: 'label', label: '缩进' },
                        { id: 'indent-none', label: '无缩进', checked: !paragraphStyle.firstLineIndentPx, onSelect: () => props.onParagraphFormat({ firstLineIndentPx: 0 }) },
                        { id: 'indent-2', label: '首行缩进 2 字符', checked: (paragraphStyle.firstLineIndentPx ?? 0) > 0,
                          onSelect: () => props.onParagraphFormat({ firstLineIndentPx: charWidthPx(activeStyle.fontSizePt ?? props.doc.defaultFontSizePt) * 2 }) },
                        { id: 'indent-hanging-2', label: '悬挂缩进 2 字符', checked: (paragraphStyle.firstLineIndentPx ?? 0) < 0,
                          onSelect: () => props.onParagraphFormat({ firstLineIndentPx: -charWidthPx(activeStyle.fontSizePt ?? props.doc.defaultFontSizePt) * 2 }) },
                        { type: 'separator' },
                        { id: 'paragraph-more', disabled: pending('paragraph'), label: '段落设置…', shortcut: `${MOD_KEY}⌥M`, onSelect: props.onOpenParagraphDialog },
                      ]} trigger={({ open, toggle }) => (
                        <button type="button" aria-label="段落" title="段落：缩进、间距与行距" aria-haspopup="menu" aria-expanded={open}
                          disabled={readonly || pending('paragraph')} onClick={toggle} className="flex h-7 items-center gap-1 rounded px-2 text-[11px] hover:bg-[var(--mo-surface-2)] disabled:opacity-40">
                          段落设置<ChevronDown size={11} />
                        </button>
                      )} />
                      {props.onToggleFormattingMarks ? (
                        <Tooltip label={props.showFormattingMarks ? '隐藏编辑标记' : '显示编辑标记'}>
                          <IconButton label="显示编辑标记" aria-pressed={!!props.showFormattingMarks} active={props.showFormattingMarks} onClick={props.onToggleFormattingMarks}><Pilcrow size={15} /></IconButton>
                        </Tooltip>
                      ) : null}
                    </div>
                    <div className="flex items-center gap-0.5">
                      <IconButton label="左对齐" active={!props.mixedParagraph?.includes('alignment') && (paragraphStyle.alignment === 'left' || !paragraphStyle.alignment)} disabled={readonly || pending('paragraph')} onClick={() => props.onParagraphFormat({ alignment: 'left' })}><AlignLeft size={15} /></IconButton>
                      <IconButton label="居中" active={paragraphStyle.alignment === 'center'} disabled={readonly || pending('paragraph')} onClick={() => props.onParagraphFormat({ alignment: 'center' })}><AlignCenter size={15} /></IconButton>
                      <IconButton label="右对齐" active={paragraphStyle.alignment === 'right'} disabled={readonly || pending('paragraph')} onClick={() => props.onParagraphFormat({ alignment: 'right' })}><AlignRight size={15} /></IconButton>
                      <IconButton label="两端对齐" active={paragraphStyle.alignment === 'justify'} disabled={readonly || pending('paragraph')} onClick={() => props.onParagraphFormat({ alignment: 'justify' })}><AlignJustify size={15} /></IconButton>
                      <LineSpacingPicker mixed={props.mixedParagraph?.some(key => ['lineHeight', 'lineHeightExactPx', 'lineHeightAtLeastPx'].includes(key))} style={paragraphStyle} disabled={readonly || pending('paragraph')} onChange={props.onParagraphFormat} onOpenOptions={props.onOpenParagraphDialog} />
                    </div>
                  </RibbonGroup>
                  {props.stylePicker ? (
                    <RibbonGroup label="样式"><StylePicker {...props.stylePicker} readonly={readonly || props.stylePicker.readonly} /></RibbonGroup>
                  ) : null}
                  <RibbonGroup label="查找">
                    <RibbonButton label="查找替换" icon={<Search size={15} />} active={props.findOpen} disabled={pending('find')} onClick={props.onToggleFind} />
                  </RibbonGroup>
                </>
              ) : null}
              {tab === '插入' ? (
                <>
                  <RibbonGroup label="表格与图片">
                    <div className="flex items-center gap-1">
                      <div className="relative flex items-center [&_button]:w-[74px] [&_button]:justify-start [&_button]:px-2">
                        <TablePicker disabled={readonly || pending('table')} onInsert={props.onInsertTable} />
                        <span aria-hidden="true" className="pointer-events-none absolute right-2 text-[11px]">表格</span>
                      </div>
                      <RibbonButton label="插入图片" icon={<ImagePlus size={15} />} disabled={readonly || pending('image')} onClick={props.onInsertImage} />
                    </div>
                  </RibbonGroup>
                  <RibbonGroup label="页面">
                    <RibbonButton label="插入分页符" icon={<Scissors size={15} />} disabled={readonly || pending('pageBreak')} onClick={props.onPageBreak} />
                    <RibbonButton label="页眉、页脚与页码" icon={<PanelTop size={15} />} disabled={readonly || pending('headerFooter')} onClick={props.onHeaderFooter} />
                  </RibbonGroup>
                </>
              ) : null}
              {tab === '布局' ? (
                <RibbonGroup label="页面与段落">
                  <div className="flex items-center gap-1">
                    <RibbonButton label="布局与属性" icon={<PanelRight size={15} />} active={props.propertiesOpen} disabled={pending('pageLayout')} onClick={props.onToggleProperties} />
                    <RibbonButton label="页眉、页脚与页码" icon={<PanelTop size={15} />} disabled={readonly || pending('headerFooter')} onClick={props.onHeaderFooter} />
                  </div>
                  <div className="flex items-center gap-1">
                    <RibbonButton label="段落设置" icon={<Pilcrow size={15} />} disabled={readonly || pending('paragraph')} onClick={props.onOpenParagraphDialog} />
                    <RibbonButton label="插入分页符" icon={<Scissors size={15} />} disabled={readonly || pending('pageBreak')} onClick={props.onPageBreak} />
                  </div>
                </RibbonGroup>
              ) : null}
              {tab === '引用' ? (
                <RibbonGroup label="目录">
                  <RibbonButton label="插入目录" icon={<TableOfContents size={15} />} disabled={readonly || pending('toc')} onClick={props.onInsertToc} />
                  {/* 功能：目录更新沿用同一引用标签，旧核心无回调时不显示；时间：2026-11-11；作者：lq。 */}
                  {props.onUpdateToc ? <RibbonButton label="更新目录" icon={<TableOfContents size={15} />} disabled={readonly || pending('toc')} onClick={props.onUpdateToc} /> : null}
                </RibbonGroup>
              ) : null}
              {tab === '邮件' ? (
                <RibbonGroup label="邮件合并">
                  <RibbonButton label="邮件合并" icon={<Mails size={15} />} disabled={readonly || pending('mailMerge')} onClick={props.onMailMerge} />
                </RibbonGroup>
              ) : null}
              {tab === '视图' ? (
                <>
                  <RibbonGroup label="显示">
                    <div className="flex items-center gap-1">
                      <RibbonButton label="标尺" icon={<Ruler size={15} />} active={props.rulerOpen} disabled={pending('ruler')} onClick={props.onToggleRuler} />
                      <RibbonButton label={props.propertiesOpen ? '隐藏属性面板' : '显示属性面板'} icon={<PanelRight size={15} />} active={props.propertiesOpen} disabled={pending('pageLayout')} onClick={props.onToggleProperties} />
                    </div>
                    {props.onToggleFormattingMarks ? (
                      <RibbonButton label="显示编辑标记" icon={<Pilcrow size={15} />} active={props.showFormattingMarks} onClick={props.onToggleFormattingMarks} />
                    ) : null}
                  </RibbonGroup>
                  <RibbonGroup label="输出与兼容">
                    <div className="flex items-center gap-1">
                      <RibbonButton label="打印" icon={<Printer size={15} />} disabled={pending('print')} onClick={props.onPrint} />
                      <RibbonButton label="导出 PDF" icon={<Download size={15} />} disabled={pending('export')} onClick={props.onExport} />
                    </div>
                    <RibbonButton label="兼容性报告" icon={<FileWarning size={15} className={readonly ? 'text-[var(--mo-warning)]' : ''} />} onClick={props.onShowCompatibility} />
                  </RibbonGroup>
                </>
              ) : null}
            </div>
          ) : null}
        </div>
      ))}
    </div>
  );
}

function KbdLabel(key: string): string {
  return navigator.platform.toLowerCase().includes('mac') ? `⌘${key}` : `Ctrl+${key}`;
}
