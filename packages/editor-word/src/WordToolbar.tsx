import type { DocParagraph, DocModel, ParagraphStyle, RunStyle } from '@mini-office/document-model';
import { FONT_CHOICES } from '@mini-office/design-system';
import { mergeFontChoices } from './use-system-fonts';
// 功能：按操作系统展示 Word 缩放快捷键，并使用统一比例上下限；时间：2026-10-24；作者：lq。
import { MOD_KEY } from '@mini-office/shared';
import { WORD_ZOOM_MIN, WORD_ZOOM_MAX } from './zoom-shortcuts';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Baseline,
  Bold,
  Download,
  FileDiff,
  FileWarning,
  Highlighter,
  ImagePlus,
  IndentDecrease,
  IndentIncrease,
  Italic,
  List,
  Mails,
  MessageSquarePlus,
  MoreHorizontal,
  Paintbrush,
  PanelRight,
  PanelTop,
  Pilcrow,
  Printer,
  Redo2,
  Ruler,
  Scissors,
  Search,
  Strikethrough,
  Underline,
  Undo2,
  ZoomIn,
} from 'lucide-react';
import { ColorField, Dropdown, IconButton, Kbd, Select, Tooltip } from '@mini-office/ui';
import type { MenuItem } from '@mini-office/ui';
import { TablePicker } from './TablePicker';
import { FontSizePicker } from './FontSizePicker';
import { LineSpacingPicker } from './LineSpacingPicker';
import { charWidthPx } from './typography';
import { StylePicker, type StylePickerProps } from './StylePicker';

export interface WordToolbarProps {
  doc: DocModel;
  /** 功能：系统字体列表（平台枚举），与内置列表合并进字体下拉框；时间：2026-10-30；作者：hou0。 */
  systemFonts?: string[];
  readonly: boolean;
  activeParagraph?: DocParagraph;
  activeStyle: RunStyle;
  canUndo: boolean;
  canRedo: boolean;
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
  /** 功能：Word 标尺显示开关；时间：2026-10-30；作者：hou0。 */
  showRuler: boolean;
  onToggleRuler: () => void;
  /** 功能：审阅面板（批注与修订）开关与新建批注入口；时间：2026-10-31。 */
  reviewOpen?: boolean;
  onToggleReview?: () => void;
  onAddComment?: () => void;
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
  onBeforeInteraction?: () => void;
}

function Divider() {
  return <span className="mx-1 h-4 w-px shrink-0 bg-[var(--mo-line)]" aria-hidden />;
}

export function WordToolbar(props: WordToolbarProps) {
  const { readonly, activeStyle, activeParagraph } = props;
  const paragraphStyle = activeParagraph?.style ?? {};
  const moreItems: MenuItem[] = [
    { id: 'mail-merge', label: '邮件合并 / 批量报告', icon: <Mails size={14} />, disabled: readonly, onSelect: props.onMailMerge },
    // 功能：在窄屏更多菜单中保留页眉页脚入口；时间：2026-10-24；作者：lq。
    { id: 'header-footer', label: '页眉、页脚与页码', icon: <PanelTop size={14} />, disabled: readonly, onSelect: props.onHeaderFooter },
    {
      id: 'page-break',
      label: '插入分页符',
      icon: <Scissors size={14} />,
      disabled: readonly,
      onSelect: props.onPageBreak,
    },
    {
      id: 'print',
      label: '打印',
      icon: <Printer size={14} />,
      onSelect: props.onPrint,
    },
    {
      id: 'export',
      label: '导出 PDF',
      icon: <Download size={14} />,
      onSelect: props.onExport,
    },
    { type: 'separator' },
    {
      id: 'compat',
      label: '兼容性报告',
      icon: <FileWarning size={14} />,
      onSelect: props.onShowCompatibility,
    },
    {
      id: 'properties',
      label: props.propertiesOpen ? '隐藏属性面板' : '显示属性面板',
      icon: <PanelRight size={14} />,
      onSelect: props.onToggleProperties,
    },
  ];

  return (
    <div
      className="mo-glass mo-no-print flex h-11 shrink-0 items-center gap-0.5 overflow-x-auto border-b border-[var(--mo-line)] bg-[var(--mo-surface)] px-2"
      onPointerDownCapture={() => props.onBeforeInteraction?.()}
      onMouseDown={(event) => {
        if ((event.target as HTMLElement).closest('button')) event.preventDefault();
      }}
    >
      <Tooltip label={`撤销 ${KbdLabel('Z')}`}>
        <IconButton label="撤销" disabled={!props.canUndo || readonly} onClick={props.onUndo}>
          <Undo2 size={15} />
        </IconButton>
      </Tooltip>
      <Tooltip label={`重做 ${KbdLabel('Shift+Z')}`}>
        <IconButton label="重做" disabled={!props.canRedo || readonly} onClick={props.onRedo}>
          <Redo2 size={15} />
        </IconButton>
      </Tooltip>
      <Divider />
      {/* 功能：样式库紧跟撤销重做，作为「整段排版」的第一入口；时间：2026-10-29；作者：lq。 */}
      {props.stylePicker ? <StylePicker {...props.stylePicker} /> : null}
      <Divider />
      <Select
        aria-label="字体"
        className="h-7 w-[132px] shrink-0 border-transparent bg-transparent text-[12px] hover:bg-[var(--mo-surface-2)]"
        value={activeStyle.fontFamily ?? '默认'}
        disabled={readonly}
        onChange={(event) => {
          const value = event.target.value;
          props.onFontFormat({ fontFamily: value === '默认' ? undefined : value });
        }}
      >
        {mergeFontChoices(props.systemFonts ?? [], activeStyle.fontFamily, props.doc.defaultFont).map((font) => (
          <option key={font} value={font}>
            {font}
          </option>
        ))}
      </Select>
      <FontSizePicker
        value={activeStyle.fontSizePt ?? props.doc.defaultFontSizePt}
        disabled={readonly}
        onChange={(pt) => props.onFontFormat({ fontSizePt: pt })}
      />
      <Divider />
      <Tooltip label={`加粗 ${KbdLabel('B')}`}>
        <IconButton
          label="加粗"
          active={!!activeStyle.bold}
          disabled={readonly}
          onClick={() => props.onInlineCommand('bold')}
        >
          <Bold size={15} />
        </IconButton>
      </Tooltip>
      <Tooltip label={`斜体 ${KbdLabel('I')}`}>
        <IconButton
          label="斜体"
          active={!!activeStyle.italic}
          disabled={readonly}
          onClick={() => props.onInlineCommand('italic')}
        >
          <Italic size={15} />
        </IconButton>
      </Tooltip>
      <Tooltip label={`下划线 ${KbdLabel('U')}`}>
        <IconButton
          label="下划线"
          active={!!activeStyle.underline}
          disabled={readonly}
          onClick={() => props.onInlineCommand('underline')}
        >
          <Underline size={15} />
        </IconButton>
      </Tooltip>
      <IconButton
        label="删除线"
        active={!!activeStyle.strike}
        disabled={readonly}
        onClick={() => props.onInlineCommand('strikeThrough')}
      >
        <Strikethrough size={15} />
      </IconButton>
      <ColorField
        // 功能：用文字及荧光笔图标区分颜色工具并显示当前色条；时间：2026-10-19；作者：lq。
        label="字体颜色"
        icon={<Baseline size={16} />}
        indicatorColor={activeStyle.color ?? '#1D2939'}
        disabled={readonly}
        value={activeStyle.color}
        onChange={(color) => props.onColorFormat({ color }, 'color')}
      />
      <ColorField
        label="文本背景颜色（突出显示）"
        icon={<Highlighter size={16} />}
        indicatorColor={activeStyle.highlight ?? 'transparent'}
        disabled={readonly}
        value={activeStyle.highlight}
        onChange={(color) => props.onColorFormat({ highlight: color }, 'highlight')}
      />
      {/* 功能：格式刷紧跟字体格式区，单击复制一次、双击锁定连续套用；时间：2026-10-29；作者：lq。 */}
      <Tooltip label="格式刷：单击复制格式并套用一次，双击可连续套用，Esc 取消">
        <IconButton
          label="格式刷"
          active={props.formatPainterActive}
          disabled={readonly}
          onClick={() => props.onFormatPainter(false)}
          onDoubleClick={() => props.onFormatPainter(true)}
        >
          <Paintbrush size={15} />
        </IconButton>
      </Tooltip>
      <Divider />
      <IconButton
        label="左对齐"
        active={paragraphStyle.alignment === 'left' || !paragraphStyle.alignment}
        disabled={readonly}
        onClick={() => props.onParagraphFormat({ alignment: 'left' })}
      >
        <AlignLeft size={15} />
      </IconButton>
      <IconButton
        label="居中"
        active={paragraphStyle.alignment === 'center'}
        disabled={readonly}
        onClick={() => props.onParagraphFormat({ alignment: 'center' })}
      >
        <AlignCenter size={15} />
      </IconButton>
      <IconButton
        label="右对齐"
        active={paragraphStyle.alignment === 'right'}
        disabled={readonly}
        onClick={() => props.onParagraphFormat({ alignment: 'right' })}
      >
        <AlignRight size={15} />
      </IconButton>
      <IconButton
        label="两端对齐"
        active={paragraphStyle.alignment === 'justify'}
        disabled={readonly}
        onClick={() => props.onParagraphFormat({ alignment: 'justify' })}
      >
        <AlignJustify size={15} />
      </IconButton>
      <LineSpacingPicker
        style={paragraphStyle}
        disabled={readonly}
        onChange={props.onParagraphFormat}
        onOpenOptions={props.onOpenParagraphDialog}
      />
      {/* 功能：段落分组统一入口——常用缩进档位与段落对话框合并为一个菜单，避免两个段落标记图标并排；时间：2026-10-29；作者：lq。 */}
      <Dropdown
        minWidth={212}
        items={[
          { type: 'label', label: '缩进' },
          {
            id: 'indent-none',
            label: '无缩进',
            checked: !paragraphStyle.firstLineIndentPx,
            onSelect: () => props.onParagraphFormat({ firstLineIndentPx: 0 }),
          },
          {
            id: 'indent-2',
            label: '首行缩进 2 字符',
            checked: (paragraphStyle.firstLineIndentPx ?? 0) > 0,
            onSelect: () =>
              props.onParagraphFormat({
                firstLineIndentPx: charWidthPx(activeStyle.fontSizePt ?? props.doc.defaultFontSizePt) * 2,
              }),
          },
          {
            id: 'indent-hanging-2',
            label: '悬挂缩进 2 字符',
            checked: (paragraphStyle.firstLineIndentPx ?? 0) < 0,
            onSelect: () =>
              props.onParagraphFormat({
                firstLineIndentPx: -charWidthPx(activeStyle.fontSizePt ?? props.doc.defaultFontSizePt) * 2,
              }),
          },
          { type: 'separator' },
          {
            id: 'paragraph-more',
            label: '段落设置…',
            shortcut: `${MOD_KEY}⌥M`,
            onSelect: props.onOpenParagraphDialog,
          },
        ]}
        trigger={({ toggle }) => (
          <Tooltip label="段落：缩进、间距与行距">
            <button
              type="button"
              aria-label="段落"
              disabled={readonly}
              onClick={toggle}
              className="flex h-7 shrink-0 items-center gap-1 rounded px-2 text-[11px] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)] disabled:opacity-40"
            >
              <Pilcrow size={13} />
              段落
            </button>
          </Tooltip>
        )}
      />
      <IconButton
        label="增加缩进"
        disabled={readonly}
        onClick={() =>
          props.onParagraphFormat({
            indentLeftPx: (paragraphStyle.indentLeftPx ?? 0) + 24,
          })
        }
      >
        <IndentIncrease size={15} />
      </IconButton>
      <IconButton
        label="减少缩进"
        disabled={readonly}
        onClick={() =>
          props.onParagraphFormat({
            indentLeftPx: Math.max(0, (paragraphStyle.indentLeftPx ?? 0) - 24),
          })
        }
      >
        <IndentDecrease size={15} />
      </IconButton>
      <Divider />
      <IconButton label="项目符号" disabled={readonly} onClick={props.onAddBullets}>
        <List size={15} />
      </IconButton>
      <IconButton label="编号" disabled={readonly} onClick={props.onAddNumbering}>
        <span className="text-[11px] font-semibold">1.</span>
      </IconButton>
      <TablePicker disabled={readonly} onInsert={props.onInsertTable} />
      {/* 功能：直接显示批量报告入口，避免首次使用难以发现；时间：2026-10-24；作者：lq。 */}
      <button type="button" disabled={readonly} onClick={props.onMailMerge} className="flex h-7 shrink-0 items-center gap-1 rounded px-2 text-xs hover:bg-[var(--mo-surface-2)] disabled:opacity-40"><Mails size={14} />邮件合并</button>
      {/* 功能：常驻页眉页脚按钮；时间：2026-10-24；作者：lq。 */}
      <Tooltip label="页眉、页脚与页码"><IconButton label="页眉、页脚与页码" disabled={readonly} onClick={props.onHeaderFooter}><PanelTop size={15} /></IconButton></Tooltip>
      <IconButton label="插入图片" disabled={readonly} onClick={props.onInsertImage}>
        <ImagePlus size={15} />
      </IconButton>
      <Divider />
      <Tooltip label={`查找替换 ${KbdLabel('F')}`}>
        <IconButton label="查找替换" active={props.findOpen} onClick={props.onToggleFind}>
          <Search size={15} />
        </IconButton>
      </Tooltip>
      {/* 功能：常驻打开页面布局与文档属性，支持页边距、方向和纸张大小；时间：2026-10-24；作者：lq。 */}
      <IconButton label="布局与属性" active={props.propertiesOpen} onClick={props.onToggleProperties}>
        <PanelRight size={15} />
      </IconButton>
      {/* 功能：Word 标尺显示开关；时间：2026-10-30；作者：hou0。 */}
      <IconButton label="标尺" active={props.showRuler} onClick={props.onToggleRuler}>
        <Ruler size={15} />
      </IconButton>
      {props.onAddComment ? (
        <IconButton label="新建批注" onClick={props.onAddComment}>
          <MessageSquarePlus size={15} />
        </IconButton>
      ) : null}
      {props.onToggleReview ? (
        <IconButton label="审阅（批注与修订）" active={props.reviewOpen} onClick={props.onToggleReview}>
          <FileDiff size={15} />
        </IconButton>
      ) : null}
      <Tooltip label="打印 ⌘P">
        <IconButton label="打印" onClick={props.onPrint}>
          <Printer size={15} />
        </IconButton>
      </Tooltip>
      <div className="ml-auto flex shrink-0 items-center gap-1 pl-2">
        <span className="whitespace-nowrap pr-1 text-[10px] tabular-nums text-[var(--mo-text-muted)]">
          {props.pageCount} 页 · 第 {props.currentPage} 页
        </span>
        <IconButton label="适应宽度" onClick={props.onFitWidth}>
          <ZoomIn size={15} />
        </IconButton>
        <Select
          aria-label="缩放"
          // 功能：提示原版 Word 的缩放组合键，便于用户发现；时间：2026-10-24；作者：lq。
          title={`放大 ${MOD_KEY} +；缩小 ${MOD_KEY} -；100% ${MOD_KEY} 0；${MOD_KEY} + 滚轮`}
          className="h-7 w-[70px] border-transparent bg-transparent text-[11px]"
          value={String(Math.round(props.zoom * 100))}
          onChange={(event) => props.onZoomChange(Number(event.target.value) / 100)}
        >
          {/* 功能：包含快捷键产生的任意比例，防止 120% 等值在下拉框显示为空；时间：2026-10-24；作者：lq。 */}
          {[...new Set([WORD_ZOOM_MIN * 100, 50, 75, 90, 100, 110, 125, 150, 175, 200, WORD_ZOOM_MAX * 100, Math.round(props.zoom * 100)])].sort((a, b) => a - b).map((percent) => (
            <option key={percent} value={percent}>
              {percent}%
            </option>
          ))}
        </Select>
        <IconButton
          label={props.readonly ? '只读模式' : '兼容性报告'}
          onClick={props.onShowCompatibility}
        >
          <FileWarning size={15} className={props.readonly ? 'text-[var(--mo-warning)]' : ''} />
        </IconButton>
        <Dropdown
          align="right"
          minWidth={210}
          items={moreItems}
          trigger={({ toggle }) => (
            <IconButton label="更多操作" onClick={toggle}>
              <MoreHorizontal size={16} />
            </IconButton>
          )}
        />
      </div>
    </div>
  );
}

function KbdLabel(key: string): string {
  return navigator.platform.toLowerCase().includes('mac') ? `⌘${key}` : `Ctrl+${key}`;
}
