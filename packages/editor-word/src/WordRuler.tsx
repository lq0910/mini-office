import { useCallback, useEffect, useMemo, useRef, type CSSProperties, type PointerEvent as ReactPointerEvent } from 'react';
import type { ParagraphStyle, TabStop } from '@mini-office/document-model';
import { DEFAULT_TAB_STOP_PX, nextTabStop } from './tab-stops';

/**
 * 功能：Word 横向标尺——显示页面宽度、页边距、缩进标记与制表位，并支持拖动调整；时间：2026-10-30；作者：lq。
 *
 * 设计要点：
 * 1. 刻度以厘米为主单位，另叠加字符刻度提示（中文排版常用「字符」对齐），0 点对齐正文左边界。
 * 2. 所有位置都在「未缩放」的页面坐标里计算，只在渲染时乘 zoom，因此缩放不影响取值精度。
 * 3. 拖动过程用稳定的撤销 key 合并成一条历史记录，不会因为拖动产生几十条撤销。
 */

const CM_PX = 37.7952755906;
const RULER_HEIGHT = 22;

export interface WordRulerProps {
  pageWidthPx: number;
  contentLeftPx: number;
  contentWidthPx: number;
  zoom: number;
  style: ParagraphStyle;
  defaultTabStopPx?: number;
  readonly: boolean;
  /** 功能：新核心拖动期间预览草稿，结束一次提交，取消/卸载丢弃；时间：2026-11-11；作者：lq。 */
  onDragStart?: () => void;
  onDragEnd?: (cancelled: boolean) => void;
  onChangeIndent: (patch: Partial<ParagraphStyle>, key: string) => void;
  onChangeTabStops: (stops: TabStop[], key: string) => void;
}

function snapToQuarter(value: number): number {
  const step = CM_PX / 4;
  return Math.round(value / step) * step;
}

interface DragState {
  startX: number;
  initial: number;
}

export function WordRuler({
  pageWidthPx,
  contentLeftPx,
  contentWidthPx,
  zoom,
  style,
  defaultTabStopPx,
  readonly,
  onChangeIndent,
  onChangeTabStops,
  onDragStart,
  onDragEnd,
}: WordRulerProps) {
  // 功能：拖动使用开始时的回调，组件卸载时释放全局监听并取消草稿；时间：2026-11-11；作者：lq。
  const cleanupDrag = useRef<() => void>();
  useEffect(() => () => cleanupDrag.current?.(), []);
  const indentLeft = style.indentLeftPx ?? 0;
  const indentRight = style.indentRightPx ?? 0;
  const firstLine = style.firstLineIndentPx ?? 0;
  const tabStops = useMemo(
    () => (style.tabStops ?? []).slice().sort((a, b) => a.posPx - b.posPx),
    [style.tabStops],
  );

  const ticks = useMemo(() => {
    const items: Array<{ x: number; major: boolean; label?: string }> = [];
    const minorStep = CM_PX / 4;
    const first = Math.floor((-contentLeftPx) / minorStep) * minorStep;
    for (let offset = first; offset <= contentWidthPx + contentLeftPx; offset += minorStep) {
      const x = contentLeftPx + offset;
      if (x < -0.5 || x > pageWidthPx + 0.5) continue;
      const isCm = Math.abs(offset / CM_PX - Math.round(offset / CM_PX)) < 0.01;
      const label = isCm && Math.round(offset / CM_PX) > 0 ? String(Math.round(offset / CM_PX)) : undefined;
      items.push({ x, major: isCm, label });
    }
    return items;
  }, [contentLeftPx, contentWidthPx, pageWidthPx]);

  const startDrag = useCallback(
    (event: ReactPointerEvent<HTMLElement>, kind: 'left' | 'firstLine' | 'right' | 'tab', tabIndex = -1) => {
      if (readonly || event.button !== 0) return;
      event.preventDefault();
      event.stopPropagation();
      cleanupDrag.current?.();
      onDragStart?.();
      const state: DragState = { startX: event.clientX, initial: 0 };
      if (kind === 'left') state.initial = indentLeft;
      else if (kind === 'right') state.initial = indentRight;
      else if (kind === 'firstLine') state.initial = firstLine;
      else state.initial = tabStops[tabIndex]?.posPx ?? 0;

      const move = (pointer: PointerEvent) => {
        const delta = (pointer.clientX - state.startX) / (zoom || 1);
        if (kind === 'left') {
          const value = Math.max(0, Math.min(contentWidthPx - 24, snapToQuarter(state.initial + delta)));
          onChangeIndent({ indentLeftPx: value }, 'ruler-indent-left');
          return;
        }
        if (kind === 'right') {
          const value = Math.max(0, Math.min(contentWidthPx - 24, snapToQuarter(state.initial - delta)));
          onChangeIndent({ indentRightPx: value }, 'ruler-indent-right');
          return;
        }
        if (kind === 'firstLine') {
          const value = snapToQuarter(state.initial + delta);
          onChangeIndent({ firstLineIndentPx: value }, 'ruler-first-line');
          return;
        }
        const nextPos = Math.max(0, Math.min(contentWidthPx, snapToQuarter(state.initial + delta)));
        const next = tabStops.map((stop, index) => (index === tabIndex ? { ...stop, posPx: nextPos } : stop));
        onChangeTabStops(next, 'ruler-tab-move');
      };
      const cleanup = () => {
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', finish);
        window.removeEventListener('pointercancel', finish);
        cleanupDrag.current = undefined;
      };
      const finish = (pointer: PointerEvent) => {
        cleanup();
        if (kind === 'tab' && tabIndex >= 0 && pointer.type !== 'pointercancel') {
          const delta = (pointer.clientX - state.startX) / (zoom || 1);
          if (state.initial + delta < -8) onChangeTabStops(tabStops.filter((_, index) => index !== tabIndex), 'ruler-tab-remove');
        }
        onDragEnd?.(pointer.type === 'pointercancel');
      };
      cleanupDrag.current = () => { cleanup(); onDragEnd?.(true); };
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', finish);
      window.addEventListener('pointercancel', finish);
    },
    [contentWidthPx, firstLine, indentLeft, indentRight, onChangeIndent, onChangeTabStops, onDragStart, onDragEnd, readonly, tabStops, zoom],
  );

  /** 功能：在标尺空白处单击新增一个左对齐制表位，落在最近的默认制表位网格上；时间：2026-10-30；作者：lq。 */
  const addTabStop = useCallback(
    (event: ReactPointerEvent<HTMLDivElement>) => {
      if (readonly || event.button !== 0) return;
      const rect = event.currentTarget.getBoundingClientRect();
      const x = (event.clientX - rect.left) / (zoom || 1);
      const relative = x - contentLeftPx;
      if (relative < 0 || relative > contentWidthPx) return;
      const target = nextTabStop(tabStops, relative - 1, defaultTabStopPx ?? DEFAULT_TAB_STOP_PX);
      const posPx = snapToQuarter(target.posPx);
      if (tabStops.some((stop) => Math.abs(stop.posPx - posPx) < 1)) return;
      onChangeTabStops([...tabStops, { posPx, alignment: 'left' }], 'ruler-tab-add');
    },
    [contentLeftPx, contentWidthPx, defaultTabStopPx, onChangeTabStops, readonly, tabStops, zoom],
  );

  const removeTabStop = useCallback(
    (index: number) => {
      if (readonly) return;
      onChangeTabStops(tabStops.filter((_, position) => position !== index), 'ruler-tab-remove');
    },
    [onChangeTabStops, readonly, tabStops],
  );

  const markerStyle = (left: number): CSSProperties => ({
    position: 'absolute',
    left: left * zoom,
    top: 0,
    transform: 'translateX(-50%)',
    width: 11,
    height: 11,
    cursor: readonly ? 'default' : 'ew-resize',
    pointerEvents: 'auto',
  });

  return (
    <div
      className="mo-no-print mo-ruler"
      contentEditable={false}
      onPointerDown={(event) => event.stopPropagation()}
      style={{
        position: 'relative',
        width: pageWidthPx * zoom,
        height: RULER_HEIGHT,
        flex: '0 0 auto',
        background: 'var(--mo-surface-2)',
        border: '1px solid var(--mo-line)',
        borderRadius: 3,
        boxSizing: 'border-box',
        overflow: 'hidden',
        userSelect: 'none',
      }}
    >
      {/* 功能：页边距区域用更深的底色标出，正文可用宽度一眼可见；时间：2026-10-30；作者：lq。 */}
      <div style={{ position: 'absolute', inset: 0, pointerEvents: 'none' }}>
        <div style={{ position: 'absolute', left: 0, top: 0, bottom: 0, width: contentLeftPx * zoom, background: 'var(--mo-surface-3)' }} />
        <div
          style={{
            position: 'absolute',
            right: 0,
            top: 0,
            bottom: 0,
            width: (pageWidthPx - contentLeftPx - contentWidthPx) * zoom,
            background: 'var(--mo-surface-3)',
          }}
        />
      </div>

      <div
        style={{ position: 'absolute', inset: 0, cursor: readonly ? 'default' : 'text' }}
        onPointerDown={addTabStop}
        title={readonly ? undefined : '单击添加制表位；拖动制表位可移动，双击可删除'}
      >
        {ticks.map((tick, index) => (
          <span
            key={`${tick.x}-${index}`}
            style={{
              position: 'absolute',
              left: tick.x * zoom,
              bottom: 0,
              width: 1,
              height: tick.major ? 7 : 4,
              background: tick.major ? 'var(--mo-text-faint)' : 'var(--mo-line-strong)',
              pointerEvents: 'none',
            }}
          />
        ))}
        {ticks.map((tick, index) =>
          tick.label ? (
            <span
              key={`label-${tick.x}-${index}`}
              style={{
                position: 'absolute',
                left: tick.x * zoom + 2,
                top: 1,
                fontSize: 9,
                lineHeight: '10px',
                color: 'var(--mo-text-muted)',
                pointerEvents: 'none',
              }}
            >
              {tick.label}
            </span>
          ) : null,
        )}
        {/* 功能：制表位标记，左/中/右/小数点用不同符号区分；时间：2026-10-30；作者：lq。 */}
        {tabStops.map((stop, index) => (
          <span
            key={`tab-${stop.posPx}-${index}`}
            role="button"
            aria-label={`制表位 ${Math.round(stop.posPx)}px`}
            title={`${stop.alignment === 'right' ? '右对齐' : stop.alignment === 'center' ? '居中' : stop.alignment === 'decimal' ? '小数点对齐' : '左对齐'}制表位：双击删除`}
            onPointerDown={(event) => startDrag(event, 'tab', index)}
            onDoubleClick={(event) => {
              event.stopPropagation();
              removeTabStop(index);
            }}
            style={{
              ...markerStyle(contentLeftPx + stop.posPx),
              top: 12,
              height: 8,
              display: 'flex',
              alignItems: 'flex-start',
              justifyContent: 'center',
              color: 'var(--mo-accent)',
              fontSize: 9,
              lineHeight: '8px',
              fontWeight: 700,
            }}
          >
            {stop.alignment === 'right' ? '⌐' : stop.alignment === 'center' ? '⊥' : stop.alignment === 'decimal' ? '⋮' : 'L'}
          </span>
        ))}
      </div>

      {/* 功能：首行缩进（上方三角）与左缩进（下方方块）分开拖动；时间：2026-10-30；作者：lq。 */}
      <span
        role="button"
        aria-label="首行缩进"
        title="首行缩进"
        onPointerDown={(event) => startDrag(event, 'firstLine')}
        style={{ ...markerStyle(contentLeftPx + indentLeft + firstLine), top: 2, cursor: readonly ? 'default' : 'ew-resize' }}
      >
        <svg viewBox="0 0 11 7" width="11" height="7"><path d="M0 0h11L5.5 7z" fill="var(--mo-text-primary)" /></svg>
      </span>
      <span
        role="button"
        aria-label="左缩进"
        title="左缩进"
        onPointerDown={(event) => startDrag(event, 'left')}
        style={{ ...markerStyle(contentLeftPx + indentLeft), top: 12, cursor: readonly ? 'default' : 'ew-resize' }}
      >
        <svg viewBox="0 0 11 7" width="11" height="7"><path d="M0 7h11L5.5 0z" fill="var(--mo-text-primary)" /></svg>
      </span>
      <span
        role="button"
        aria-label="右缩进"
        title="右缩进"
        onPointerDown={(event) => startDrag(event, 'right')}
        style={{ ...markerStyle(contentLeftPx + contentWidthPx - indentRight), top: 2, cursor: readonly ? 'default' : 'ew-resize' }}
      >
        <svg viewBox="0 0 11 7" width="11" height="7"><path d="M0 7h11L5.5 0z" fill="var(--mo-text-primary)" /></svg>
      </span>
    </div>
  );
}
