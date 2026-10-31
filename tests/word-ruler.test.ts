import { describe, expect, it } from 'vitest';
import {
  indentMarkerX,
  draggedIndent,
  rulerTicks,
  snapTabStop,
  type RulerMetrics,
  type RulerIndents,
} from '../packages/editor-word/src/word-ruler';

describe('indentMarkerX', () => {
  const base: RulerMetrics = { contentLeft: 0, contentWidth: 1000, zoom: 1 };

  it('places markers at content edges when indents are zero', () => {
    const x = indentMarkerX(base, { left: 0, firstLine: 0, right: 0 });
    expect(x.left).toBe(0);
    expect(x.firstLine).toBe(0);
    expect(x.right).toBe(1000);
  });

  it('offsets markers by indents (zoom 1)', () => {
    const ind: RulerIndents = { left: 50, firstLine: 20, right: 30 };
    const x = indentMarkerX(base, ind);
    expect(x.left).toBe(50);
    expect(x.firstLine).toBe(70);
    expect(x.right).toBe(970);
  });

  it('scales indents by zoom', () => {
    const ind: RulerIndents = { left: 50, firstLine: 20, right: 30 };
    const x = indentMarkerX({ contentLeft: 0, contentWidth: 1000, zoom: 2 }, ind);
    expect(x.left).toBe(100);
    expect(x.firstLine).toBe(140);
    expect(x.right).toBe(940);
  });

  it('shifts all markers by contentLeft', () => {
    const ind: RulerIndents = { left: 50, firstLine: 20, right: 30 };
    const x = indentMarkerX({ contentLeft: 120, contentWidth: 1000, zoom: 1 }, ind);
    expect(x.left).toBe(170);
    expect(x.firstLine).toBe(190);
    expect(x.right).toBe(1090);
  });
});

describe('draggedIndent', () => {
  it('converts screen delta to model px at zoom 1', () => {
    expect(draggedIndent(0, 100, 1, 0)).toBe(100);
  });

  it('divides delta by zoom', () => {
    expect(draggedIndent(0, 100, 2, 0)).toBe(50);
  });

  it('clamps to min', () => {
    expect(draggedIndent(50, -200, 1, 0)).toBe(0);
  });

  it('allows hanging via negative min', () => {
    expect(draggedIndent(50, -200, 1, -50)).toBe(-50);
  });

  it('rounds to whole model px', () => {
    expect(draggedIndent(20, 7, 1, 0)).toBe(27);
  });
});

describe('rulerTicks', () => {
  it('starts a major tick (label 0) at content left', () => {
    const m: RulerMetrics = { contentLeft: 100, contentWidth: 400, zoom: 1 };
    const ticks = rulerTicks(m);
    expect(ticks[0].x).toBeCloseTo(100, 0);
    expect(ticks[0].major).toBe(true);
    expect(ticks[0].label).toBe('0');
  });

  it('labels whole-cm majors and interleaves minor ticks', () => {
    const m: RulerMetrics = { contentLeft: 0, contentWidth: 756, zoom: 1 };
    const ticks = rulerTicks(m);
    const majors = ticks.filter((t) => t.major).map((t) => t.label);
    expect(majors).toContain('1');
    expect(majors).toContain('2');
    expect(ticks.some((t) => !t.major)).toBe(true);
  });

  it('scales tick x positions by zoom', () => {
    const m: RulerMetrics = { contentLeft: 0, contentWidth: 378, zoom: 2 };
    const one = rulerTicks(m).find((t) => t.major && t.label === '1');
    expect(one).toBeDefined();
    expect(one!.x).toBeCloseTo(2 * 37.7953, 0);
  });
});

describe('snapTabStop', () => {
  it('snaps to nearest half cm', () => {
    const half = 37.7953 / 2;
    expect(snapTabStop(0)).toBe(0);
    expect(snapTabStop(half * 0.4)).toBeCloseTo(0, 0);
    expect(snapTabStop(half * 0.6)).toBeCloseTo(half, 0);
    expect(snapTabStop(half * 1.4)).toBeCloseTo(half, 0);
  });
});
