/** Ruler geometry in screen px (already zoom-scaled). */
export interface RulerMetrics {
  /** Content box left edge in screen px, relative to the ruler container. */
  contentLeft: number;
  /** Content box width in screen px. */
  contentWidth: number;
  /** Zoom factor (1 = 100%). */
  zoom: number;
}

/** Paragraph indent values in model px (what the ruler edits). */
export interface RulerIndents {
  /** Paragraph `indentLeftPx` (>= 0). */
  left: number;
  /** Paragraph `firstLineIndentPx` (positive = first-line, negative = hanging). */
  firstLine: number;
  /** Paragraph `indentRightPx` (>= 0). */
  right: number;
}

/** Screen x (relative to the ruler container) of the three indent markers. */
export function indentMarkerX(m: RulerMetrics, ind: RulerIndents): RulerIndents {
  return {
    left: m.contentLeft + ind.left * m.zoom,
    firstLine: m.contentLeft + (ind.left + ind.firstLine) * m.zoom,
    right: m.contentLeft + m.contentWidth - ind.right * m.zoom,
  };
}

/**
 * New indent value after dragging a marker by `deltaScreenPx`.
 * Converts screen delta to model px via `/zoom`, rounds to whole px, clamps to `min`.
 */
export function draggedIndent(original: number, deltaScreenPx: number, zoom: number, min: number): number {
  return Math.max(min, Math.round(original + deltaScreenPx / zoom));
}

/** 1cm in CSS px at 96dpi. */
const CM_PX = 37.79527559055118;

/** A scale tick on the ruler. */
export interface RulerTick {
  /** Screen x relative to the ruler container. */
  x: number;
  /** Major tick = whole cm, carries a number label. */
  major: boolean;
  /** Number label for major ticks. */
  label?: string;
}

/** Build a cm scale across the content box: major ticks every cm, minor every 0.5cm. */
export function rulerTicks(m: RulerMetrics): RulerTick[] {
  const ticks: RulerTick[] = [];
  const maxCm = m.contentWidth / m.zoom / CM_PX;
  for (let cm = 0; cm <= maxCm + 0.0001; cm += 0.5) {
    const rounded = Math.round(cm * 2) / 2;
    const major = Math.abs(rounded - Math.round(rounded)) < 0.001;
    ticks.push({
      x: m.contentLeft + rounded * CM_PX * m.zoom,
      major,
      label: major ? String(Math.round(rounded)) : undefined,
    });
  }
  return ticks;
}

/** Snap a tab stop position (model px) to the nearest half cm. */
export function snapTabStop(posPx: number): number {
  return Math.round(posPx / (CM_PX / 2)) * (CM_PX / 2);
}
