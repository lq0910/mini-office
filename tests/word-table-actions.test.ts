// 功能：表格整体操作控件——边缘加号（右侧加列/底部加行）与整表拖柄的点击回调；时间：2026-10-30；作者：hou0。
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';
import { createBlankDoc, type DocModel, type DocTable } from '@mini-office/document-model';
import { TableGrid } from '../packages/editor-word/src/WordEditor';

function table(): DocTable {
  const cell = (id: string) => ({ id, blocks: [], colSpan: 1, rowSpan: 1 });
  return {
    type: 'table',
    id: 't1',
    columnWidthsPx: [100, 100],
    rows: [
      { id: 'r1', cells: [cell('a'), cell('b')] },
      { id: 'r2', cells: [cell('c'), cell('d')] },
    ],
  };
}

function renderGrid(handlers?: Record<string, unknown>) {
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
  const host = document.createElement('div');
  document.body.append(host);
  const doc: DocModel = createBlankDoc();
  const onTableAdd = vi.fn();
  const onTableSelectAll = vi.fn();
  const base = {
    syncToken: 0,
    onInput: vi.fn(),
    onBlur: vi.fn(),
    onKeyDown: vi.fn(),
    onTableContextMenu: vi.fn(),
    onTableResize: vi.fn(),
    onTableAdd,
    onTableSelectAll,
    ...handlers,
  };
  const root = createRoot(host);
  act(() => root.render(createElement(TableGrid, { table: table(), doc, pageNumber: 1, totalPages: 1, readonly: false, handlers: base as never })));
  return { host, root, onTableAdd, onTableSelectAll };
}

describe('表格整体操作', () => {
  it('右/下边缘加号点击后按轴回调，按钮默认隐藏、悬停带内才显示', () => {
    const { host, root, onTableAdd } = renderGrid();
    const columnButton = host.querySelector('[data-table-add="column"]') as HTMLElement;
    const rowButton = host.querySelector('[data-table-add="row"]') as HTMLElement;
    expect(columnButton).toBeTruthy();
    expect(rowButton).toBeTruthy();
    // 加号被包在悬停带里：只有鼠标进入条带时才可见。
    expect(columnButton.closest('[data-table-edge="column"]')).toBeTruthy();
    expect(rowButton.closest('[data-table-edge="row"]')).toBeTruthy();

    act(() => columnButton.dispatchEvent(new MouseEvent('click', { bubbles: true })));
    act(() => rowButton.dispatchEvent(new MouseEvent('click', { bubbles: true })));
    expect(onTableAdd.mock.calls.map((call) => call[1])).toEqual(['column', 'row']);
    expect(onTableAdd.mock.calls[0][0]).toBe('t1');
    act(() => root.unmount());
  });

  it('整表拖柄点击回调表格 id', () => {
    const { host, root, onTableSelectAll } = renderGrid();
    const grab = host.querySelector('[data-table-grab]') as HTMLElement;
    expect(grab).toBeTruthy();
    act(() => grab.dispatchEvent(new MouseEvent('click', { bubbles: true })));
    expect(onTableSelectAll).toHaveBeenCalledWith('t1');
    act(() => root.unmount());
  });

  it('单元格不再画蓝色框；多选时用淡底纹提示', () => {
    const single = renderGrid({ tableSelection: { tableId: 't1', anchorId: 'b', focusId: 'b' } });
    const singleCells = Array.from(single.host.querySelectorAll<HTMLElement>('[data-cell-id]'));
    expect(singleCells).toHaveLength(4);
    for (const cell of singleCells) {
      expect(cell.style.boxShadow).toBe('');
      expect(cell.style.backgroundColor).toBe('');
    }
    act(() => single.root.unmount());

    const multi = renderGrid({ tableSelection: { tableId: 't1', anchorId: 'a', focusId: 'd' } });
    const tinted = Array.from(multi.host.querySelectorAll<HTMLElement>('[data-cell-id]')).filter(
      (cell) => cell.style.backgroundColor.includes('rgba'),
    );
    // 整表选中后四格都有淡底纹（提示表格工具会作用到整表），但依然没有蓝框。
    expect(tinted).toHaveLength(4);
    for (const cell of tinted) expect(cell.style.boxShadow).toBe('');
    act(() => multi.root.unmount());
  });

  it('只读模式不渲染加号与拖柄', () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
    const host = document.createElement('div');
    document.body.append(host);
    const root = createRoot(host);
    const doc = createBlankDoc();
    act(() => root.render(createElement(TableGrid, {
      table: table(), doc, pageNumber: 1, totalPages: 1, readonly: true,
      handlers: { syncToken: 0, onInput: vi.fn(), onBlur: vi.fn(), onKeyDown: vi.fn(), onTableContextMenu: vi.fn(), onTableResize: vi.fn(), onTableAdd: vi.fn(), onTableSelectAll: vi.fn() } as never,
    })));
    expect(host.querySelector('[data-table-add]')).toBeNull();
    expect(host.querySelector('[data-table-grab]')).toBeNull();
    act(() => root.unmount());
  });
});
