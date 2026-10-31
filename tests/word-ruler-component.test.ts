// 功能：Word 水平标尺——三个段落缩进标记的定位与拖拽；时间：2026-10-30；作者：lq。
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';
import { WordRuler } from '../packages/editor-word/src/WordRuler';
import { snapTabStop } from '../packages/editor-word/src/word-ruler';
import type { WordTabStop } from '../packages/document-model/src/word';

function renderRuler(
  metrics = { contentLeft: 0, contentWidth: 1000, zoom: 1 },
  indents = { left: 0, firstLine: 0, right: 0 },
  tabStops: WordTabStop[] = [],
) {
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
  const host = document.createElement('div');
  document.body.append(host);
  const onChange = vi.fn();
  const onTabStopsChange = vi.fn();
  const root = createRoot(host);
  act(() => root.render(createElement(WordRuler, { metrics, indents, onChange, tabStops, onTabStopsChange })));
  return { host, root, onChange, onTabStopsChange };
}

function marker(host: HTMLElement, axis: string) {
  return host.querySelector(`[data-ruler-axis="${axis}"]`) as HTMLElement;
}

function dragMarker(host: HTMLElement, axis: string, fromX: number, toX: number) {
  const el = marker(host, axis);
  const fire = (type: string, x: number) =>
    act(() => el.dispatchEvent(new MouseEvent(type, { bubbles: true, button: 0, clientX: x })));
  fire('pointerdown', fromX);
  fire('pointermove', toX);
  fire('pointerup', toX);
}

describe('WordRuler', () => {
  it('renders three indent markers at computed positions', () => {
    const { host } = renderRuler({ contentLeft: 0, contentWidth: 1000, zoom: 1 }, { left: 50, firstLine: 20, right: 30 });
    expect(marker(host, 'left').style.left).toBe('50px');
    expect(marker(host, 'firstLine').style.left).toBe('70px');
    expect(marker(host, 'right').style.left).toBe('970px');
  });

  it('places markers at content edges for zero indents', () => {
    const { host } = renderRuler({ contentLeft: 0, contentWidth: 1000, zoom: 1 }, { left: 0, firstLine: 0, right: 0 });
    expect(marker(host, 'left').style.left).toBe('0px');
    expect(marker(host, 'firstLine').style.left).toBe('0px');
    expect(marker(host, 'right').style.left).toBe('1000px');
  });

  it('dragging the first-line marker updates firstLine indent (zoom-aware)', () => {
    const { host, onChange } = renderRuler({ contentLeft: 0, contentWidth: 1000, zoom: 2 }, { left: 0, firstLine: 0, right: 0 });
    dragMarker(host, 'firstLine', 0, 100);
    expect(onChange).toHaveBeenCalled();
    expect(onChange).toHaveBeenLastCalledWith({ left: 0, firstLine: 50, right: 0 });
  });

  it('dragging the left marker updates left indent', () => {
    const { host, onChange } = renderRuler({ contentLeft: 0, contentWidth: 1000, zoom: 1 }, { left: 10, firstLine: 0, right: 0 });
    dragMarker(host, 'left', 10, 60);
    expect(onChange).toHaveBeenLastCalledWith({ left: 60, firstLine: 0, right: 0 });
  });

  it('dragging the right marker updates right indent (inverted direction)', () => {
    const { host, onChange } = renderRuler({ contentLeft: 0, contentWidth: 1000, zoom: 1 }, { left: 0, firstLine: 0, right: 10 });
    dragMarker(host, 'right', 990, 940);
    expect(onChange).toHaveBeenLastCalledWith({ left: 0, firstLine: 0, right: 60 });
  });

  it('renders tab-stop markers at contentLeft + posPx * zoom', () => {
    const { host } = renderRuler(
      { contentLeft: 100, contentWidth: 1000, zoom: 2 },
      { left: 0, firstLine: 0, right: 0 },
      [{ posPx: 50, kind: 'left' }],
    );
    const stop = host.querySelector('[data-tabstop-index="0"]') as HTMLElement;
    expect(stop.style.left).toBe('200px');
  });

  it('renders scale ticks with major/minor and labels', () => {
    const { host } = renderRuler({ contentLeft: 0, contentWidth: 1000, zoom: 1 }, { left: 0, firstLine: 0, right: 0 });
    expect(host.querySelectorAll('[data-tick]').length).toBeGreaterThan(0);
    expect(host.querySelectorAll('[data-tick="major"]').length).toBeGreaterThan(0);
    expect(host.querySelectorAll('[data-tick-label]').length).toBeGreaterThan(0);
  });

  it('clicking the bare track adds a tab stop (snapped)', () => {
    const { host, onTabStopsChange } = renderRuler({ contentLeft: 0, contentWidth: 1000, zoom: 1 }, { left: 0, firstLine: 0, right: 0 });
    const container = host.querySelector('[data-testid="word-ruler"]') as HTMLElement;
    const track = host.querySelector('[data-tab-track]') as HTMLElement;
    vi.spyOn(container, 'getBoundingClientRect').mockReturnValue({ left: 0, top: 0 } as DOMRect);
    act(() => track.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, button: 0, clientX: 100 })));
    expect(onTabStopsChange).toHaveBeenCalled();
    expect(onTabStopsChange).toHaveBeenLastCalledWith([{ posPx: snapTabStop(100), kind: 'left' }]);
  });

  it('dragging a tab marker moves it (snapped)', () => {
    const { host, onTabStopsChange } = renderRuler(
      { contentLeft: 0, contentWidth: 1000, zoom: 1 },
      { left: 0, firstLine: 0, right: 0 },
      [{ posPx: 100, kind: 'left' }],
    );
    const stop = host.querySelector('[data-tabstop-index="0"]') as HTMLElement;
    const fire = (type: string, x: number) => act(() => stop.dispatchEvent(new MouseEvent(type, { bubbles: true, button: 0, clientX: x })));
    fire('pointerdown', 100);
    fire('pointermove', 150);
    fire('pointerup', 150);
    expect(onTabStopsChange).toHaveBeenCalled();
    expect(onTabStopsChange).toHaveBeenLastCalledWith([{ posPx: snapTabStop(150), kind: 'left' }]);
  });

  it('clicking a tab marker without moving removes it', () => {
    const { host, onTabStopsChange } = renderRuler(
      { contentLeft: 0, contentWidth: 1000, zoom: 1 },
      { left: 0, firstLine: 0, right: 0 },
      [{ posPx: 100, kind: 'left' }],
    );
    const stop = host.querySelector('[data-tabstop-index="0"]') as HTMLElement;
    const fire = (type: string, x: number) => act(() => stop.dispatchEvent(new MouseEvent(type, { bubbles: true, button: 0, clientX: x })));
    fire('pointerdown', 100);
    fire('pointerup', 100);
    expect(onTabStopsChange).toHaveBeenLastCalledWith([]);
  });
});
