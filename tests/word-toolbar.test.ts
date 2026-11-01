/**
 * 功能：验证表格尺寸选择、键盘操作及颜色工具标识，防止插入固定尺寸回归。
 * 时间：2026-10-19；作者：lq
 */
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { TablePicker } from '../packages/editor-word/src/TablePicker';
import { ColorField } from '../packages/ui/src/color-field';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
  (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  container = document.createElement('div');
  document.body.append(container);
  root = createRoot(container);
});
afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

// 功能：通过真实 DOM 事件驱动弹层交互；时间：2026-10-19；作者：lq。
function click(element: Element | null) {
  expect(element).not.toBeNull();
  act(() => element!.dispatchEvent(new MouseEvent('click', { bubbles: true })));
}
function openTable(onInsert = vi.fn(), disabled = false) {
  act(() => root.render(createElement(TablePicker, { onInsert, disabled })));
  click(container.querySelector('button'));
  return onInsert;
}

describe('Word 表格选择器', () => {
  it('点击网格按实际行列提交，并关闭弹层', () => {
    const insert = openTable();
    expect(document.querySelectorAll('[data-table-cell]')).toHaveLength(80);
    click(document.querySelector('[aria-label="5 列 4 行"]'));
    expect(insert).toHaveBeenCalledWith(4, 5);
    expect(document.querySelector('[role="dialog"]')).toBeNull();
  });

  it('方向键移动选区，Escape 取消且不插入', () => {
    const insert = openTable();
    const cell = document.activeElement!;
    act(() => cell.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true })));
    expect(document.activeElement?.getAttribute('aria-label')).toBe('1 列 2 行');
    act(() => document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })));
    expect(insert).not.toHaveBeenCalled();
    expect(document.querySelector('[role="dialog"]')).toBeNull();
  });

  it('自定义表格提交输入尺寸，空值阻止提交', () => {
    const insert = openTable();
    const inputs = document.querySelectorAll<HTMLInputElement>('input[type="number"]');
    const setValue = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
    act(() => {
      setValue.call(inputs[0], '12');
      inputs[0].dispatchEvent(new Event('input', { bubbles: true }));
      setValue.call(inputs[1], '9');
      inputs[1].dispatchEvent(new Event('input', { bubbles: true }));
    });
    act(() => document.querySelector('form')!.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })));
    expect(insert).toHaveBeenCalledWith(9, 12);
    openTable(insert);
    act(() => {
      const input = document.querySelector('input')!;
      setValue.call(input, '');
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    act(() => document.querySelector('form')!.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })));
    expect(insert).toHaveBeenCalledTimes(1);
    expect(document.querySelector<HTMLButtonElement>('button[type="submit"]')?.disabled).toBe(true);
  });

  it('只读时不能打开表格选择器', () => {
    const insert = openTable(vi.fn(), true);
    expect(document.querySelector('[role="dialog"]')).toBeNull();
    expect(insert).not.toHaveBeenCalled();
  });
});

describe('Word 颜色图标', () => {
  it('图标色条显示当前色，调色板仍可选色', () => {
    const change = vi.fn();
    act(() => root.render(createElement(ColorField, {
      label: '字体颜色', value: '#D92D20', icon: createElement('span', null, 'A'), onChange: change,
    })));
    expect(container.textContent).toBe('A');
    expect(container.querySelector<HTMLElement>('[aria-hidden="true"] > span:last-child')?.style.background).toBe('rgb(217, 45, 32)');
    // 功能：主按钮直接套用当前颜色，调色板改由右侧箭头展开；时间：2026-11-01；作者：lq。
    click(container.querySelector('button[aria-label="字体颜色"]'));
    expect(change).toHaveBeenCalledWith('#D92D20');
    click(container.querySelector('button[aria-label="字体颜色调色板"]'));
    click(document.querySelector('[aria-label="#3478F6"]'));
    expect(change).toHaveBeenCalledWith('#3478F6');
  });

  // 功能：颜色入口要记住上次选过的色，主按钮一键套用，不必每次重新调色；时间：2026-11-01；作者：lq。
  it('主按钮直接套用上次选过的颜色', () => {
    const first = vi.fn();
    act(() => root.render(createElement(ColorField, {
      label: '填充颜色', value: '#111111', onChange: first,
    })));
    click(container.querySelector('button[aria-label="填充颜色调色板"]'));
    click(document.querySelector('[aria-label="#3478F6"]'));
    expect(first).toHaveBeenCalledWith('#3478F6');

    const second = vi.fn();
    act(() => root.render(createElement(ColorField, {
      label: '填充颜色', value: '#111111', onChange: second,
    })));
    click(container.querySelector('button[aria-label="填充颜色"]'));
    expect(second).toHaveBeenCalledWith('#3478F6');
  });
});
