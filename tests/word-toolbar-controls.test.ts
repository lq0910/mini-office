/**
 * 功能：验证 Word 工具栏的字号、行距、格式刷与段落入口，防止格式设置能力再次退化为固定值。
 * 时间：2026-10-29；作者：lq
 */
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WordToolbar, type WordToolbarProps } from '../packages/editor-word/src/WordToolbar';

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
  document.body.innerHTML = '';
});

function props(overrides: Partial<WordToolbarProps> = {}): WordToolbarProps {
  return {
    doc: { defaultFont: '宋体', defaultFontSizePt: 12 } as WordToolbarProps['doc'],
    readonly: false,
    activeParagraph: undefined,
    activeStyle: { fontSizePt: 16, fontFamily: '宋体' },
    canUndo: false,
    canRedo: false,
    zoom: 1,
    pageCount: 1,
    currentPage: 1,
    findOpen: false,
    propertiesOpen: false,
    onInlineCommand: vi.fn(),
    onColorFormat: vi.fn(),
    onFontFormat: vi.fn(),
    onParagraphFormat: vi.fn(),
    onInsertTable: vi.fn(),
    onInsertImage: vi.fn(),
    onPageBreak: vi.fn(),
    onHeaderFooter: vi.fn(),
    onMailMerge: vi.fn(),
    onUndo: vi.fn(),
    onRedo: vi.fn(),
    onZoomChange: vi.fn(),
    onFitWidth: vi.fn(),
    onToggleFind: vi.fn(),
    onToggleProperties: vi.fn(),
    showRuler: false,
    onToggleRuler: vi.fn(),
    onExport: vi.fn(),
    onPrint: vi.fn(),
    onShowCompatibility: vi.fn(),
    onAddNumbering: vi.fn(),
    onAddBullets: vi.fn(),
    onFormatPainter: vi.fn(),
    formatPainterActive: false,
    onOpenParagraphDialog: vi.fn(),
    ...overrides,
  };
}

function click(element: Element | null | undefined) {
  expect(element).toBeTruthy();
  act(() => (element as HTMLElement).dispatchEvent(new MouseEvent('click', { bubbles: true })));
}

describe('Word 工具栏格式设置', () => {
  it('字号显示中文字号，并在面板中同时提供中文字号与磅值', () => {
    const onFontFormat = vi.fn();
    act(() => root.render(createElement(WordToolbar, props({ onFontFormat }))));
    const trigger = container.querySelector('[aria-label="字号"]');
    expect(trigger?.textContent).toContain('三号 16');

    click(trigger);
    const panel = document.body.querySelector('[data-testid="font-size-panel"]');
    expect(panel).toBeTruthy();
    // 中文字号与数字磅值分组都在面板里，用户不必记磅值。
    expect(panel!.textContent).toContain('中文字号');
    expect(panel!.textContent).toContain('初号');
    expect(panel!.textContent).toContain('八号');
    expect(panel!.textContent).toContain('数字磅值');

    const sanHao = Array.from(panel!.querySelectorAll('button')).find((button) => button.textContent?.startsWith('三号'));
    click(sanHao);
    expect(onFontFormat).toHaveBeenCalledWith({ fontSizePt: 16 });
  });

  it('字号面板支持手动输入磅值', () => {
    const onFontFormat = vi.fn();
    act(() => root.render(createElement(WordToolbar, props({ onFontFormat }))));
    click(container.querySelector('[aria-label="字号"]'));
    const input = document.body.querySelector<HTMLInputElement>('[aria-label="字号磅值"]');
    expect(input).toBeTruthy();
    const setValue = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
    act(() => {
      setValue.call(input, '13.5');
      input!.dispatchEvent(new Event('input', { bubbles: true }));
    });
    act(() => {
      input!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    });
    expect(onFontFormat).toHaveBeenCalledWith({ fontSizePt: 13.5 });
  });

  it('行距显示固定值并可自由输入磅值', () => {
    const onParagraphFormat = vi.fn();
    act(() =>
      root.render(
        createElement(
          WordToolbar,
          props({
            onParagraphFormat,
            activeParagraph: { type: 'paragraph', id: 'p1', runs: [], style: { lineHeightExactPx: 24 } } as WordToolbarProps['activeParagraph'],
          }),
        ),
      ),
    );
    const trigger = container.querySelector('[aria-label="行距"]');
    expect(trigger?.textContent).toContain('固定 18 磅');

    click(trigger);
    const input = document.body.querySelector<HTMLInputElement>('[aria-label="固定值行距"]');
    expect(input).toBeTruthy();
    const setValue = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
    act(() => {
      setValue.call(input, '22');
      input!.dispatchEvent(new Event('input', { bubbles: true }));
    });
    const applyButton = Array.from(document.body.querySelectorAll('button')).filter((button) => button.textContent === '应用');
    click(applyButton[0]);
    expect(onParagraphFormat).toHaveBeenCalled();
    const patch = onParagraphFormat.mock.calls[0][0] as { lineHeightExactPx?: number };
    expect(patch.lineHeightExactPx).toBeCloseTo(29.33, 1);
  });

  it('格式刷单击与双击分别提交单次与连续模式', () => {
    const onFormatPainter = vi.fn();
    act(() => root.render(createElement(WordToolbar, props({ onFormatPainter }))));
    const button = container.querySelector<HTMLButtonElement>('[aria-label="格式刷"]');
    click(button);
    expect(onFormatPainter).toHaveBeenCalledWith(false);
    act(() => button!.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })));
    expect(onFormatPainter).toHaveBeenCalledWith(true);
  });

  it('段落菜单同时提供缩进档位与段落对话框入口', () => {
    const onOpenParagraphDialog = vi.fn();
    const onParagraphFormat = vi.fn();
    act(() => root.render(createElement(WordToolbar, props({ onOpenParagraphDialog, onParagraphFormat }))));

    click(container.querySelector('[aria-label="段落"]'));
    const indentItem = Array.from(document.body.querySelectorAll('button')).find((button) =>
      button.textContent?.includes('首行缩进 2 字符'),
    );
    click(indentItem);
    expect(onParagraphFormat).toHaveBeenCalled();
    const patch = onParagraphFormat.mock.calls[0][0] as { firstLineIndentPx?: number };
    // 16pt 字号下 2 字符 ≈ 42.66px。
    expect(patch.firstLineIndentPx).toBeCloseTo(42.66, 1);

    click(container.querySelector('[aria-label="段落"]'));
    const dialogItem = Array.from(document.body.querySelectorAll('button')).find((button) =>
      button.textContent?.includes('段落设置'),
    );
    click(dialogItem);
    expect(onOpenParagraphDialog).toHaveBeenCalled();
  });

  it('只读文档禁用格式设置入口', () => {
    act(() => root.render(createElement(WordToolbar, props({ readonly: true }))));
    expect(container.querySelector<HTMLButtonElement>('[aria-label="格式刷"]')?.disabled).toBe(true);
    expect(container.querySelector<HTMLButtonElement>('[aria-label="字号"]')?.disabled).toBe(true);
  });
});
