import { useRef } from 'react';
import type { WordTabStop } from '@mini-office/document-model';
import {
  indentMarkerX,
  draggedIndent,
  rulerTicks,
  snapTabStop,
  type RulerMetrics,
  type RulerIndents,
} from './word-ruler';

export interface WordRulerProps {
  metrics: RulerMetrics;
  indents: RulerIndents;
  onChange: (indents: RulerIndents) => void;
  tabStops: WordTabStop[];
  onTabStopsChange: (next: WordTabStop[]) => void;
}

type DragAxis = 'left' | 'firstLine' | 'right';

export function WordRuler({ metrics, indents, onChange, tabStops, onTabStopsChange }: WordRulerProps) {
  const x = indentMarkerX(metrics, indents);
  const ticks = rulerTicks(metrics);
  const dragRef = useRef<{ axis: DragAxis; startX: number; startIndent: number } | null>(null);
  const tabDragRef = useRef<{ index: number; startX: number; startPos: number; moved: boolean } | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);

  const startDrag = (axis: DragAxis) => (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.pointerId != null) {
      try {
        (e.currentTarget as Element).setPointerCapture(e.pointerId);
      } catch {
        /* jsdom */
      }
    }
    const startIndent = axis === 'left' ? indents.left : axis === 'right' ? indents.right : indents.firstLine;
    dragRef.current = { axis, startX: e.clientX, startIndent };
  };

  const onMove = (e: React.PointerEvent) => {
    const drag = dragRef.current;
    if (drag) {
      const rawDelta = e.clientX - drag.startX;
      const delta = drag.axis === 'right' ? -rawDelta : rawDelta;
      const min = drag.axis === 'firstLine' ? -indents.left : 0;
      const value = draggedIndent(drag.startIndent, delta, metrics.zoom, min);
      onChange({ ...indents, [drag.axis]: value });
      return;
    }
    const tab = tabDragRef.current;
    if (tab) {
      const raw = e.clientX - tab.startX;
      if (Math.abs(raw) > 1) tab.moved = true;
      const pos = snapTabStop(tab.startPos + raw / metrics.zoom);
      onTabStopsChange(tabStops.map((t, i) => (i === tab.index ? { ...t, posPx: pos } : t)));
    }
  };

  const endDrag = () => {
    const tab = tabDragRef.current;
    if (tab) {
      if (!tab.moved) onTabStopsChange(tabStops.filter((_, i) => i !== tab.index));
      tabDragRef.current = null;
    }
    dragRef.current = null;
  };

  const startTabDrag = (index: number) => (e: React.PointerEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.pointerId != null) {
      try {
        (e.currentTarget as Element).setPointerCapture(e.pointerId);
      } catch {
        /* jsdom */
      }
    }
    tabDragRef.current = { index, startX: e.clientX, startPos: tabStops[index]?.posPx ?? 0, moved: false };
  };

  const onTrackPointerDown = (e: React.PointerEvent) => {
    if (e.target !== e.currentTarget) return;
    const rect = containerRef.current?.getBoundingClientRect();
    if (!rect) return;
    const posPx = snapTabStop((e.clientX - rect.left - metrics.contentLeft) / metrics.zoom);
    onTabStopsChange([...tabStops, { posPx, kind: 'left' }]);
  };

  return (
    <div
      data-testid="word-ruler"
      role="presentation"
      ref={containerRef}
      className="relative h-6 select-none border-b border-[var(--mo-border)] bg-[var(--mo-canvas)]"
      onPointerMove={onMove}
      onPointerUp={endDrag}
      onPointerCancel={endDrag}
    >
      {/* 内容白区：页边距之外露出灰色画布背景。 */}
      <div
        data-tab-track
        className="absolute top-0 h-full border-x border-[var(--mo-border-soft)] bg-[var(--mo-bg)]"
        style={{ left: metrics.contentLeft, width: metrics.contentWidth }}
        onPointerDown={onTrackPointerDown}
      />
      {/* 刻度线与数字 */}
      {ticks.map((t, i) => (
        <div
          key={i}
          data-tick={t.major ? 'major' : 'minor'}
          className={`absolute bottom-0 w-px ${t.major ? 'h-2' : 'h-1'} bg-[var(--mo-border)]`}
          style={{ left: t.x }}
        />
      ))}
      {ticks
        .filter((t) => t.major)
        .map((t, i) => (
          <span
            key={i}
            data-tick-label
            className="absolute bottom-2.5 -translate-x-1/2 text-[9px] leading-none text-[var(--mo-muted)]"
            style={{ left: t.x }}
          >
            {t.label}
          </span>
        ))}
      {/* 制表位：点击空白添加、拖动移动、点击已有删除 */}
      {tabStops.map((t, i) => (
        <div
          key={i}
          data-tabstop
          data-tabstop-index={i}
          role="button"
          aria-label="制表位"
          title={`制表位（${t.posPx.toFixed(0)}px）`}
          className="absolute top-0 z-10 h-3 w-3 -translate-x-1/2 cursor-ew-resize"
          style={{ left: metrics.contentLeft + t.posPx * metrics.zoom }}
          onPointerDown={startTabDrag(i)}
        >
          <span className="block h-2 w-2 border border-[var(--mo-accent)]" />
        </div>
      ))}
      {/* 首行缩进：上三角（顶部） */}
      <div
        data-ruler-axis="firstLine"
        role="separator"
        aria-label="首行缩进"
        title="首行缩进"
        className="absolute top-0 z-10 -translate-x-1/2 cursor-ew-resize"
        style={{ left: x.firstLine }}
        onPointerDown={startDrag('firstLine')}
      >
        <span className="block h-0 w-0 border-x-[5px] border-x-transparent border-t-[7px] border-t-[var(--mo-accent)]" />
      </div>
      {/* 左缩进 + 悬挂：矩形 + 上三角（底部） */}
      <div
        data-ruler-axis="left"
        role="separator"
        aria-label="左缩进"
        title="左缩进"
        className="absolute bottom-0 z-10 -translate-x-1/2 cursor-ew-resize"
        style={{ left: x.left }}
        onPointerDown={startDrag('left')}
      >
        <span className="mx-auto block h-0 w-0 border-x-[5px] border-x-transparent border-b-[7px] border-b-[var(--mo-accent)]" />
        <span className="mx-auto block h-[6px] w-[7px] bg-[var(--mo-accent)]" />
      </div>
      {/* 右缩进：下三角（底部） */}
      <div
        data-ruler-axis="right"
        role="separator"
        aria-label="右缩进"
        title="右缩进"
        className="absolute bottom-0 z-10 -translate-x-1/2 cursor-ew-resize"
        style={{ left: x.right }}
        onPointerDown={startDrag('right')}
      >
        <span className="block h-0 w-0 border-x-[5px] border-x-transparent border-b-[7px] border-b-[var(--mo-accent)]" />
      </div>
    </div>
  );
}
