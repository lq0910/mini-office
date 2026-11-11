/**
 * 功能：样式库面板，参考 WPS「预设样式」浮层，提供实时预览、批量套用与样式集管理；时间：2026-10-29；作者：lq。
 *
 * 相比 WPS 的三处改进：
 * 1. 卡片直接用样式自身字体渲染名称，所见即所得，不用先套用再看效果；
 * 2. 鼠标悬停时在标题右侧显示该样式的完整规格（字号、行距、缩进、段前段后），避免逐个试错；
 * 3. 面板支持方向键在样式间移动、回车套用，键盘操作不断手。
 */
import { useCallback, useMemo, useRef, useState, type KeyboardEvent as ReactKeyboardEvent, type ReactNode } from 'react';
import type { DocStyleDefinition } from '@mini-office/document-model';
import { fontStackFor } from '@mini-office/design-system';
import { cn } from '@mini-office/ui';
import { ChevronDown, ChevronUp, Eraser, FolderPlus, Pencil, Sparkles, Trash2, X } from 'lucide-react';
import { BASIC_STYLE_COUNT, type WordStyleSet } from './styles';
import { cnFontSizeName, formatPt, pxToChars, pxToPt, readLineSpacing } from './typography';

export interface StyleGalleryProps {
  presets: DocStyleDefinition[];
  documentStyles: DocStyleDefinition[];
  currentStyleId?: string;
  showMore: boolean;
  defaultFont: string;
  defaultFontSizePt: number;
  readonly: boolean;
  styleSets: WordStyleSet[];
  onApply: (id: string) => void;
  onClearFormatting: () => void;
  onNewStyle: () => void;
  onToggleMore: () => void;
  onDeleteStyle: (id: string) => void;
  /** 功能：新核心支持原定义原位重命名，可单独限制导入定义的删除；时间：2026-11-10；作者：lq。 */
  onRenameStyle?: (id: string) => void;
  canDeleteStyle?: (id: string) => boolean;
  /** 功能：新核心只导出已完整迁移的新建格式，避免把导入复杂定义降成摘要；时间：2026-11-10；作者：lq。 */
  canSaveStyleSet?: boolean;
  onSaveStyleSet: () => void;
  onApplyStyleSet: (set: WordStyleSet) => void;
  onDeleteStyleSet: (id: string) => void;
  onBeforeInteraction?: () => void;
}

/** 功能：把样式规格压缩成一行说明，供悬停提示使用；时间：2026-10-29；作者：lq。 */
export function describeStyle(style: DocStyleDefinition, defaultFontSizePt: number): string {
  const fontSize = style.run.fontSizePt ?? defaultFontSizePt;
  const sizeLabel = cnFontSizeName(fontSize) ?? `${formatPt(fontSize)} 磅`;
  const parts: string[] = [sizeLabel];
  if (style.run.bold) parts.push('加粗');
  if (style.run.italic) parts.push('斜体');
  if (style.run.fontFamily) parts.push(style.run.fontFamily);
  const alignment = style.paragraph.alignment ?? 'left';
  parts.push(
    alignment === 'center'
      ? '居中'
      : alignment === 'right'
        ? '右对齐'
        : alignment === 'justify'
          ? '两端对齐'
          : '左对齐',
  );
  const spacing = readLineSpacing(style.paragraph);
  if (spacing.mode === 'multiple') parts.push(`${formatPt(spacing.value)} 倍行距`);
  else parts.push(`${spacing.mode === 'exact' ? '固定' : '最小'} ${formatPt(spacing.value)} 磅`);
  const indent = style.paragraph.firstLineIndentPx ?? 0;
  if (indent > 0) parts.push(`首行缩进 ${formatPt(Math.round(pxToChars(indent, fontSize) * 10) / 10)} 字`);
  if (indent < 0) parts.push(`悬挂缩进 ${formatPt(Math.round(pxToChars(-indent, fontSize) * 10) / 10)} 字`);
  if ((style.paragraph.indentLeftPx ?? 0) > 0 && indent >= 0) {
    parts.push(`左缩进 ${formatPt(Math.round(pxToChars(style.paragraph.indentLeftPx ?? 0, fontSize) * 10) / 10)} 字`);
  }
  const before = style.paragraph.spaceBeforePx ?? 0;
  const after = style.paragraph.spaceAfterPx ?? 0;
  if (before || after) {
    parts.push(`段前 ${formatPt(Math.round(pxToPt(before) * 10) / 10)} / 段后 ${formatPt(Math.round(pxToPt(after) * 10) / 10)} 磅`);
  }
  if (style.list) parts.push(style.list.ordered ? '编号列表' : '项目符号');
  return parts.join(' · ');
}

function StyleCard({
  style,
  active,
  defaultFont,
  defaultFontSizePt,
  disabled,
  onApply,
  onDelete,
  onRename,
  onHover,
}: {
  style: DocStyleDefinition;
  active: boolean;
  defaultFont: string;
  defaultFontSizePt: number;
  disabled?: boolean;
  onApply: () => void;
  onDelete?: () => void;
  onRename?: () => void;
  onHover: (id: string | undefined) => void;
}) {
  const fontSize = style.run.fontSizePt ?? defaultFontSizePt;
  const alignment = style.paragraph.alignment ?? 'left';
  return (
    <div className="group relative">
      <button
        type="button"
        disabled={disabled}
        title={describeStyle(style, defaultFontSizePt)}
        aria-pressed={active}
        data-style-card="1"
        onClick={onApply}
        // 功能：阻止按下时抢走正文选区，样式才能作用于用户刚才拖选的范围；时间：2026-10-29；作者：lq。
        onMouseDown={(event) => event.preventDefault()}
        onPointerEnter={() => onHover(style.id)}
        onFocus={() => onHover(style.id)}
        onPointerLeave={() => onHover(undefined)}
        className={cn(
          'flex h-[46px] w-full items-center overflow-hidden rounded-md border px-2 transition-colors',
          active
            ? 'border-[var(--mo-accent)] bg-[var(--mo-accent-soft)]'
            : 'border-[var(--mo-line)] bg-[var(--mo-surface)] hover:border-[var(--mo-line-strong)] hover:bg-[var(--mo-surface-2)]',
          disabled && 'cursor-default opacity-50',
          onDelete && 'pr-6',
        )}
        style={{ justifyContent: alignment === 'center' ? 'center' : alignment === 'right' ? 'flex-end' : 'flex-start' }}
      >
        <span
          className="block max-w-full overflow-hidden text-ellipsis whitespace-nowrap"
          style={{
            fontFamily: fontStackFor(style.run.fontFamily ?? defaultFont),
            fontSize: `${fontSize}pt`,
            fontWeight: style.run.bold ? 700 : 400,
            fontStyle: style.run.italic ? 'italic' : 'normal',
            color: style.run.color ?? 'var(--mo-text-primary)',
            lineHeight: 1.15,
          }}
        >
          {style.name}
        </span>
      </button>
      {onRename ? <button type="button" disabled={disabled} aria-label={`重命名样式 ${style.name}`} title={`重命名样式「${style.name}」`}
        onMouseDown={event => event.preventDefault()} onClick={onRename}
        className="absolute bottom-1 right-1 rounded p-0.5 text-[var(--mo-text-faint)] opacity-0 hover:bg-[var(--mo-surface-3)] focus-visible:opacity-100 group-hover:opacity-100"><Pencil size={11} /></button> : null}
      {onDelete ? (
        <button
          type="button"
          disabled={disabled}
          aria-label={`删除样式 ${style.name}`}
          title={`删除样式「${style.name}」`}
          onMouseDown={(event) => event.preventDefault()}
          onClick={onDelete}
          className="absolute right-1 top-1 rounded p-0.5 text-[var(--mo-text-faint)] opacity-0 transition-opacity hover:bg-[var(--mo-surface-3)] hover:text-[var(--mo-danger)] focus-visible:opacity-100 group-hover:opacity-100"
        >
          <X size={11} />
        </button>
      ) : null}
    </div>
  );
}

function SectionTitle({ children, hint }: { children: string; hint?: string }) {
  return (
    <div className="flex items-baseline justify-between px-0.5 pb-1 pt-2">
      <span className="text-[11px] font-semibold uppercase tracking-wide text-[var(--mo-text-faint)]">{children}</span>
      {hint ? <span className="max-w-[62%] truncate text-[10px] text-[var(--mo-text-faint)]">{hint}</span> : null}
    </div>
  );
}

/** 功能：面板内的动作行，图标 + 主标题 + 快捷键字母，保持与 WPS 一致的视觉节奏；时间：2026-10-29；作者：lq。 */
function ActionRow({
  icon,
  label,
  hint,
  disabled,
  onClick,
}: {
  icon: React.ReactNode;
  label: string;
  hint?: string;
  disabled?: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      disabled={disabled}
      onMouseDown={(event) => event.preventDefault()}
      onClick={onClick}
      className="flex w-full items-center gap-2 rounded px-1.5 py-1.5 text-left text-[12px] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)] disabled:cursor-default disabled:opacity-40"
    >
      <span className="text-[var(--mo-text-muted)]">{icon}</span>
      <span className="flex-1 truncate">{label}</span>
      {hint ? <span className="text-[10px] tabular-nums text-[var(--mo-text-faint)]">{hint}</span> : null}
    </button>
  );
}

export function StyleGallery(props: StyleGalleryProps) {
  const { presets, documentStyles, currentStyleId, showMore, readonly } = props;
  const [hoveredId, setHoveredId] = useState<string | undefined>();
  const gridRef = useRef<HTMLDivElement>(null);

  const visiblePresets = useMemo(
    () => (showMore ? presets : presets.slice(0, BASIC_STYLE_COUNT)),
    [presets, showMore],
  );

  const hovered = useMemo(
    () => [...presets, ...documentStyles].find((style) => style.id === hoveredId),
    [documentStyles, hoveredId, presets],
  );

  const current = useMemo(
    () => [...presets, ...documentStyles].find((style) => style.id === currentStyleId),
    [currentStyleId, documentStyles, presets],
  );

  // 功能：方向键在样式卡片间移动，回车由按钮原生触发；时间：2026-10-29；作者：lq。
  const handleGridKeyDown = useCallback((event: ReactKeyboardEvent<HTMLDivElement>) => {
    const keys = ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'];
    if (!keys.includes(event.key)) return;
    const container = gridRef.current;
    if (!container) return;
    const buttons = Array.from(container.querySelectorAll<HTMLButtonElement>('button[data-style-card="1"]'));
    if (buttons.length === 0) return;
    const index = buttons.findIndex((button) => button === document.activeElement);
    const columns = 4;
    const delta =
      event.key === 'ArrowLeft' ? -1 : event.key === 'ArrowRight' ? 1 : event.key === 'ArrowUp' ? -columns : columns;
    const next = Math.min(buttons.length - 1, Math.max(0, (index < 0 ? 0 : index) + delta));
    event.preventDefault();
    buttons[next]?.focus();
  }, []);

  const renderCards = (styles: DocStyleDefinition[], deletable: boolean) => (
    <div ref={gridRef} onKeyDown={handleGridKeyDown} className="grid grid-cols-4 gap-1">
      {styles.map((style) => (
        <div key={style.id} data-style-card-host={style.id}>
          <StyleCard
            style={style}
            active={style.id === currentStyleId}
            defaultFont={props.defaultFont}
            defaultFontSizePt={props.defaultFontSizePt}
            disabled={readonly}
            onApply={() => props.onApply(style.id)}
            onDelete={deletable && (props.canDeleteStyle?.(style.id) ?? true) ? () => props.onDeleteStyle(style.id) : undefined}
            onRename={deletable && props.onRenameStyle ? () => props.onRenameStyle!(style.id) : undefined}
            onHover={setHoveredId}
          />
        </div>
      ))}
    </div>
  );

  return (
    <div
      className="flex max-h-[min(560px,calc(100dvh-120px))] flex-col text-[12px]"
      // 功能：面板内任何按下动作都先记录正文选区，避免拖选后点样式丢失范围；时间：2026-10-29；作者：lq。
      onPointerDownCapture={() => props.onBeforeInteraction?.()}
    >
      <div className="flex items-center justify-between px-0.5 pb-1">
        <span className="text-[12px] font-semibold text-[var(--mo-text-primary)]">预设样式</span>
        <span className="max-w-[64%] truncate text-[10px] text-[var(--mo-text-faint)]" title={hovered ? describeStyle(hovered, props.defaultFontSizePt) : undefined}>
          {hovered
            ? describeStyle(hovered, props.defaultFontSizePt)
            : current
              ? `当前：${current.name}`
              : '当前：自定义格式'}
        </span>
      </div>

      <div className="min-h-0 flex-1 overflow-y-auto pr-0.5">
        {renderCards(visiblePresets, false)}

        {documentStyles.length > 0 ? (
          <>
            <SectionTitle hint="本文档新建或从 DOCX 导入">本文档样式</SectionTitle>
            {renderCards(documentStyles, true)}
          </>
        ) : null}

        {props.styleSets.length > 0 ? (
          <>
            <SectionTitle hint="点击套用到本文档">我的样式集</SectionTitle>
            <div className="space-y-0.5">
              {props.styleSets.map((set) => (
                <div
                  key={set.id}
                  className="group flex items-center gap-1 rounded px-1.5 py-1 hover:bg-[var(--mo-surface-2)]"
                >
                  <button
                    type="button"
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={() => props.onApplyStyleSet(set)}
                    className="flex min-w-0 flex-1 items-center gap-2 text-left text-[12px]"
                    title={`套用样式集「${set.name}」（${set.styles.length} 个样式）`}
                  >
                    <FolderPlus size={13} className="shrink-0 text-[var(--mo-text-muted)]" />
                    <span className="truncate">{set.name}</span>
                    <span className="shrink-0 text-[10px] text-[var(--mo-text-faint)]">{set.styles.length}</span>
                  </button>
                  <button
                    type="button"
                    aria-label={`删除样式集 ${set.name}`}
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={() => props.onDeleteStyleSet(set.id)}
                    className="rounded p-0.5 text-[var(--mo-text-faint)] opacity-0 hover:bg-[var(--mo-surface-3)] hover:text-[var(--mo-danger)] group-hover:opacity-100"
                  >
                    <Trash2 size={12} />
                  </button>
                </div>
              ))}
            </div>
          </>
        ) : null}
      </div>

      <div className="mt-1 space-y-0.5 border-t border-[var(--mo-line)] pt-1">
        <ActionRow
          icon={<Sparkles size={14} />}
          label="新建样式…"
          disabled={readonly}
          onClick={props.onNewStyle}
        />
        <ActionRow
          icon={<Eraser size={14} />}
          label="清除格式"
          disabled={readonly}
          onClick={props.onClearFormatting}
        />
        <ActionRow
          icon={showMore ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
          label={showMore ? '收起更多样式' : '显示更多样式'}
          onClick={props.onToggleMore}
        />
        <ActionRow
          icon={<FolderPlus size={14} />}
          label="保存此文档为我的样式集"
          disabled={readonly || props.canSaveStyleSet === false}
          onClick={props.onSaveStyleSet}
        />
      </div>
    </div>
  );
}
