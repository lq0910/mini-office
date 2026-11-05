/**
 * 功能：验证 Word 工具栏的字号、行距、格式刷与段落入口，防止格式设置能力再次退化为固定值。
 * 时间：2026-10-29；作者：lq
 */
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WordToolbar, type WordToolbarProps } from '../packages/editor-word/src/WordToolbar';
// 功能：注入系统字体枚举结果，验证真实工具栏的字体列表与选中行为；时间：2026-11-03；作者：lq。
import { useWorkspaceStore } from '@mini-office/document-model';

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
  vi.restoreAllMocks();
});

function props(overrides: Partial<WordToolbarProps> = {}): WordToolbarProps {
  return {
    doc: { defaultFont: '宋体', defaultFontSizePt: 12 } as WordToolbarProps['doc'],
    readonly: false,
    activeParagraph: undefined,
    activeStyle: { fontSizePt: 16, fontFamily: '宋体' },
    canUndo: false,
    canRedo: false,
    saving: false,
    onSave: vi.fn(),
    onSaveAs: vi.fn(),
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
    onExport: vi.fn(),
    onPrint: vi.fn(),
    onShowCompatibility: vi.fn(),
    onAddNumbering: vi.fn(),
    onAddBullets: vi.fn(),
    onFormatPainter: vi.fn(),
    formatPainterActive: false,
    onOpenParagraphDialog: vi.fn(),
    // 功能：标尺与目录入口为工具栏常驻项，测试同样需要提供；时间：2026-10-30；作者：lq。
    rulerOpen: true,
    onToggleRuler: vi.fn(),
    onInsertToc: vi.fn(),
    ...overrides,
  };
}

function click(element: Element | null | undefined) {
  expect(element).toBeTruthy();
  act(() => (element as HTMLElement).dispatchEvent(new MouseEvent('click', { bubbles: true })));
}

function selectTab(label: string) {
  click(Array.from(container.querySelectorAll('[role="tab"]')).find((tab) => tab.textContent === label));
}

function control(label: string) {
  return container.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`);
}

function visiblePanel() {
  return container.querySelector<HTMLElement>('[role="tabpanel"]:not([hidden])');
}

describe('Word 工具栏快捷保存', () => {
  it('保存位于撤销之前，主按钮直接保存而不展开菜单', () => {
    const onSave = vi.fn();
    const onSaveAs = vi.fn();
    act(() => root.render(createElement(WordToolbar, props({ onSave, onSaveAs }))));
    const buttons = Array.from(container.querySelectorAll('button'));
    expect(buttons.slice(0, 4).map((button) => button.getAttribute('aria-label')))
      .toEqual(['保存', '保存选项', '撤销', '重做']);
    click(buttons[0]);
    expect(onSave).toHaveBeenCalledOnce();
    expect(onSaveAs).not.toHaveBeenCalled();
    expect(document.body.querySelector('[role="menu"]')).toBeNull();
  });

  it('箭头单独展开菜单，保存和另存为调用各自回调并关闭菜单', () => {
    const onSave = vi.fn();
    const onSaveAs = vi.fn();
    act(() => root.render(createElement(WordToolbar, props({ onSave, onSaveAs }))));
    const trigger = container.querySelector<HTMLButtonElement>('button[aria-label="保存选项"]')!;
    expect(trigger.getAttribute('aria-haspopup')).toBe('menu');
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
    click(trigger);
    expect(onSave).not.toHaveBeenCalled();
    expect(trigger.getAttribute('aria-expanded')).toBe('true');
    const items = document.body.querySelectorAll('[role="menuitem"]');
    expect(items).toHaveLength(2);
    expect(items[0].textContent).toContain('保存');
    expect(items[1].textContent).toContain('另存为副本…');
    click(items[0]);
    expect(onSave).toHaveBeenCalledOnce();
    expect(document.body.querySelector('[role="menu"]')).toBeNull();
    click(trigger);
    click(document.body.querySelectorAll('[role="menuitem"]')[1]);
    expect(onSaveAs).toHaveBeenCalledOnce();
    expect(onSave).toHaveBeenCalledOnce();
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
  });

  it.each([
    ['MacIntel', '⌘S', '⇧⌘S'],
    ['Win32', 'Ctrl+S', 'Ctrl+Shift+S'],
  ])('保存提示和菜单快捷键按平台展示：%s', (platform, saveKey, saveAsKey) => {
    vi.spyOn(navigator, 'platform', 'get').mockReturnValue(platform);
    act(() => root.render(createElement(WordToolbar, props())));
    expect(container.querySelector('[aria-label="保存"]')?.getAttribute('data-tooltip')).toBe(`保存 ${saveKey}`);
    click(container.querySelector('[aria-label="保存选项"]'));
    const items = document.body.querySelectorAll('[role="menuitem"]');
    expect(items[0].textContent).toContain(saveKey);
    expect(items[1].textContent).toContain(saveAsKey);
  });

  it.each([{ readonly: true }, { saving: true }])('只读或保存中禁用主按钮和箭头：%j', (state) => {
    const onSave = vi.fn();
    act(() => root.render(createElement(WordToolbar, props({ ...state, onSave }))));
    const save = container.querySelector<HTMLButtonElement>('button[aria-label="保存"]')!;
    const trigger = container.querySelector<HTMLButtonElement>('button[aria-label="保存选项"]')!;
    expect(save.disabled).toBe(true);
    expect(trigger.disabled).toBe(true);
    click(save);
    click(trigger);
    expect(onSave).not.toHaveBeenCalled();
    expect(document.body.querySelector('[role="menu"]')).toBeNull();
    if ('saving' in state) {
      expect(save.getAttribute('aria-busy')).toBe('true');
      expect(save.getAttribute('data-tooltip')).toBe('正在保存…');
    }
  });

  it('已展开的菜单随保存状态禁用，保存结束后恢复可用', () => {
    const onSave = vi.fn();
    const onSaveAs = vi.fn();
    const callbacks = { onSave, onSaveAs };
    act(() => root.render(createElement(WordToolbar, props(callbacks))));
    click(container.querySelector('[aria-label="保存选项"]'));
    act(() => root.render(createElement(WordToolbar, props({ ...callbacks, saving: true }))));
    const items = document.body.querySelectorAll<HTMLButtonElement>('[role="menuitem"]');
    expect(Array.from(items).every((item) => item.disabled)).toBe(true);
    click(items[0]);
    click(items[1]);
    expect(onSave).not.toHaveBeenCalled();
    expect(onSaveAs).not.toHaveBeenCalled();
    act(() => root.render(createElement(WordToolbar, props(callbacks))));
    expect(container.querySelector<HTMLButtonElement>('button[aria-label="保存"]')?.disabled).toBe(false);
    expect(container.querySelector('[aria-label="保存"]')?.getAttribute('aria-busy')).toBe('false');
    click(document.body.querySelectorAll('[role="menuitem"]')[0]);
    expect(onSave).toHaveBeenCalledOnce();
  });

  it('保存按钮沿用选区保护，菜单可通过 Esc 关闭', () => {
    const onBeforeInteraction = vi.fn();
    act(() => root.render(createElement(WordToolbar, props({ onBeforeInteraction }))));
    const save = container.querySelector<HTMLButtonElement>('button[aria-label="保存"]')!;
    const mouseDown = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
    act(() => {
      save.dispatchEvent(new Event('pointerdown', { bubbles: true }));
      save.dispatchEvent(mouseDown);
    });
    expect(onBeforeInteraction).toHaveBeenCalledOnce();
    expect(mouseDown.defaultPrevented).toBe(true);
    click(container.querySelector('[aria-label="保存选项"]'));
    act(() => document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })));
    expect(document.body.querySelector('[role="menu"]')).toBeNull();
  });
});

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
    const button = container.querySelector<HTMLButtonElement>('button[aria-label="格式刷"]');
    click(button);
    expect(onFormatPainter).toHaveBeenCalledWith(false);
    act(() => button!.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })));
    expect(onFormatPainter).toHaveBeenCalledWith(true);
  });

  it('段落菜单同时提供缩进档位与段落对话框入口', () => {
    const onOpenParagraphDialog = vi.fn();
    const onParagraphFormat = vi.fn();
    act(() => root.render(createElement(WordToolbar, props({ onOpenParagraphDialog, onParagraphFormat }))));

    click(control('段落'));
    const indentItem = Array.from(document.body.querySelectorAll('button')).find((button) =>
      button.textContent?.includes('首行缩进 2 字符'),
    );
    click(indentItem);
    expect(onParagraphFormat).toHaveBeenCalled();
    const patch = onParagraphFormat.mock.calls[0][0] as { firstLineIndentPx?: number };
    // 16pt 字号下 2 字符 ≈ 42.66px。
    expect(patch.firstLineIndentPx).toBeCloseTo(42.66, 1);

    click(control('段落'));
    const dialogItem = Array.from(document.body.querySelectorAll('[role="menuitem"]')).find((button) =>
      button.textContent?.includes('段落设置'),
    );
    click(dialogItem);
    expect(onOpenParagraphDialog).toHaveBeenCalled();
  });

  it('只读文档禁用格式设置入口', () => {
    act(() => root.render(createElement(WordToolbar, props({ readonly: true }))));
    expect(container.querySelector<HTMLButtonElement>('button[aria-label="格式刷"]')?.disabled).toBe(true);
    expect(container.querySelector<HTMLButtonElement>('button[aria-label="字号"]')?.disabled).toBe(true);
  });

  it('视图与引用面板直接提供标尺与目录入口，标尺按开关状态显示激活', () => {
    const onToggleRuler = vi.fn();
    const onInsertToc = vi.fn();
    act(() => root.render(createElement(WordToolbar, props({ onToggleRuler, onInsertToc, rulerOpen: false }))));

    selectTab('视图');
    const ruler = container.querySelector<HTMLButtonElement>('button[aria-label="标尺"]');
    expect(ruler).toBeTruthy();
    // 标尺关闭时按钮不显示激活态，点击后交给上层切换。
    expect(ruler!.getAttribute('aria-pressed')).toBe('false');
    click(ruler);
    expect(onToggleRuler).toHaveBeenCalled();
    selectTab('引用');
    const toc = container.querySelector<HTMLButtonElement>('button[aria-label="插入目录"]');
    expect(toc).toBeTruthy();
    click(toc);
    expect(onInsertToc).toHaveBeenCalled();
  });

  it('只读文档不允许插入目录，但仍可切换标尺', () => {
    act(() => root.render(createElement(WordToolbar, props({ readonly: true }))));
    selectTab('引用');
    expect(container.querySelector<HTMLButtonElement>('button[aria-label="插入目录"]')?.disabled).toBe(true);
    selectTab('视图');
    expect(container.querySelector<HTMLButtonElement>('button[aria-label="标尺"]')?.disabled).toBeFalsy();
  });
});

describe('Word Ribbon 选项卡与回调', () => {
  it('六个中文选项卡具备完整可访问关联，开始面板按组展示', () => {
    act(() => root.render(createElement(WordToolbar, props())));
    const tablist = container.querySelector('[role="tablist"]');
    expect(tablist?.getAttribute('aria-label')).toBe('Word 功能选项卡');
    const tabs = Array.from(tablist!.querySelectorAll<HTMLButtonElement>('[role="tab"]'));
    expect(tabs.map((tab) => tab.textContent)).toEqual(['开始', '插入', '布局', '引用', '邮件', '视图']);
    expect(tabs.map((tab) => tab.tabIndex)).toEqual([0, -1, -1, -1, -1, -1]);
    for (const tab of tabs) {
      const panel = document.getElementById(tab.getAttribute('aria-controls')!);
      expect(panel?.getAttribute('role')).toBe('tabpanel');
      expect(panel?.getAttribute('aria-labelledby')).toBe(tab.id);
      expect(panel?.hidden).toBe(tab.textContent !== '开始');
    }
    expect(Array.from(visiblePanel()!.querySelectorAll('[role="group"]')).map((group) => group.getAttribute('aria-label')))
      .toEqual(['格式刷', '字体', '段落', '查找']);
    selectTab('邮件');
    expect(visiblePanel()?.textContent).toContain('邮件合并');
    expect(control('加粗')).toBeNull();
    expect(tabs[4].getAttribute('aria-selected')).toBe('true');
  });

  it('选项卡支持方向键、首尾键，激活时同步焦点', () => {
    act(() => root.render(createElement(WordToolbar, props())));
    const tabs = Array.from(container.querySelectorAll<HTMLButtonElement>('[role="tab"]'));
    const press = (tab: HTMLButtonElement, key: string) => {
      const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true });
      act(() => tab.dispatchEvent(event));
      expect(event.defaultPrevented).toBe(true);
    };
    press(tabs[0], 'ArrowLeft');
    expect(document.activeElement).toBe(tabs[5]);
    expect(tabs[5].getAttribute('aria-selected')).toBe('true');
    press(tabs[5], 'ArrowRight');
    expect(document.activeElement).toBe(tabs[0]);
    press(tabs[0], 'End');
    expect(document.activeElement).toBe(tabs[5]);
    press(tabs[5], 'Home');
    expect(document.activeElement).toBe(tabs[0]);
    expect(tabs[0].tabIndex).toBe(0);
  });

  it('所有选项卡顶部均不再渲染页数与缩放，但保留兼容属性', () => {
    const toolbarProps = props({ zoom: 1.23, pageCount: 13, currentPage: 8 });
    act(() => root.render(createElement(WordToolbar, toolbarProps)));
    for (const tab of ['开始', '插入', '布局', '引用', '邮件', '视图']) {
      selectTab(tab);
      expect(container.querySelector('input[type="range"], select[aria-label*="缩放"]')).toBeNull();
      expect(control('适应宽度')).toBeNull();
      expect(container.textContent).not.toContain('13 页');
      expect(container.textContent).not.toContain('123%');
    }
    expect(toolbarProps.onZoomChange).not.toHaveBeenCalled();
    expect(toolbarProps.onFitWidth).not.toHaveBeenCalled();
  });

  it('折叠只隐藏面板，保存与选项卡仍常驻，选择同一或其它选项卡会展开', () => {
    const toolbarProps = props({ canUndo: true, canRedo: true });
    act(() => root.render(createElement(WordToolbar, toolbarProps)));
    click(control('折叠功能区'));
    expect(visiblePanel()).toBeNull();
    expect(control('展开功能区')?.getAttribute('aria-expanded')).toBe('false');
    expect(container.querySelectorAll('[role="tab"]')).toHaveLength(6);
    click(control('保存'));
    click(control('撤销'));
    click(control('重做'));
    expect(toolbarProps.onSave).toHaveBeenCalledOnce();
    expect(toolbarProps.onUndo).toHaveBeenCalledOnce();
    expect(toolbarProps.onRedo).toHaveBeenCalledOnce();
    selectTab('开始');
    expect(control('加粗')).toBeTruthy();
    expect(control('折叠功能区')?.getAttribute('aria-expanded')).toBe('true');
    click(control('折叠功能区'));
    selectTab('插入');
    expect(control('插入图片')).toBeTruthy();
    click(control('折叠功能区'));
    click(control('展开功能区'));
    expect(control('插入图片')).toBeTruthy();
  });

  it('字体、内联格式、列表、缩进、对齐和查找继续调用原回调', () => {
    const toolbarProps = props({ activeParagraph: { type: 'paragraph', id: 'p1', runs: [], style: { indentLeftPx: 32 } } });
    act(() => root.render(createElement(WordToolbar, toolbarProps)));
    selectTab('开始');
    for (const label of ['加粗', '斜体', '下划线', '删除线']) click(control(label));
    expect(toolbarProps.onInlineCommand).toHaveBeenNthCalledWith(1, 'bold');
    expect(toolbarProps.onInlineCommand).toHaveBeenNthCalledWith(2, 'italic');
    expect(toolbarProps.onInlineCommand).toHaveBeenNthCalledWith(3, 'underline');
    expect(toolbarProps.onInlineCommand).toHaveBeenNthCalledWith(4, 'strikeThrough');
    const font = container.querySelector<HTMLSelectElement>('select[aria-label="字体"]')!;
    act(() => {
      font.value = '默认';
      font.dispatchEvent(new Event('change', { bubbles: true }));
    });
    expect(toolbarProps.onFontFormat).toHaveBeenCalledWith({ fontFamily: undefined });
    click(control('项目符号'));
    click(control('编号'));
    expect(toolbarProps.onAddBullets).toHaveBeenCalledOnce();
    expect(toolbarProps.onAddNumbering).toHaveBeenCalledOnce();
    click(control('减少缩进'));
    click(control('增加缩进'));
    expect(toolbarProps.onParagraphFormat).toHaveBeenCalledWith({ indentLeftPx: 8 });
    expect(toolbarProps.onParagraphFormat).toHaveBeenCalledWith({ indentLeftPx: 56 });
    for (const [label, alignment] of [['左对齐', 'left'], ['居中', 'center'], ['右对齐', 'right'], ['两端对齐', 'justify']]) {
      click(control(label));
      expect(toolbarProps.onParagraphFormat).toHaveBeenCalledWith({ alignment });
    }
    click(control('查找替换'));
    expect(toolbarProps.onToggleFind).toHaveBeenCalledOnce();
  });

  // 功能：回归截图中的重复字体与内部字体问题，并验证列表整理后仍能选中及切换字体；时间：2026-11-03；作者：lq。
  it('系统字体不同字重合并后，当前字体仍选中且新选择能提交', () => {
    const previousFonts = useWorkspaceStore.getState().fonts;
    const toolbarProps = props({ activeStyle: { fontFamily: 'verdana', fontSizePt: 12 } });
    try {
      act(() => useWorkspaceStore.setState({ fonts: [
        { family: '.Hiragino Kaku Gothic Interface', style: 'Regular', source: 'system', hasChinese: false },
        { family: 'Verdana', style: 'Regular', source: 'system', hasChinese: false },
        { family: 'Verdana', style: 'Bold', source: 'system', hasChinese: false },
        { family: '宋体', style: 'Regular', source: 'system', hasChinese: true },
      ] }));
      act(() => root.render(createElement(WordToolbar, toolbarProps)));
      const font = container.querySelector<HTMLSelectElement>('select[aria-label="字体"]')!;
      expect(Array.from(font.options, (option) => option.value)).toEqual(['默认', 'verdana', '宋体']);
      expect(font.value).toBe('verdana');
      act(() => {
        font.value = '宋体';
        font.dispatchEvent(new Event('change', { bubbles: true }));
      });
      expect(toolbarProps.onFontFormat).toHaveBeenCalledWith({ fontFamily: '宋体' });
    } finally {
      act(() => useWorkspaceStore.setState({ fonts: previousFonts }));
    }
  });

  it('字体颜色与高亮仍分别提交颜色格式回调', () => {
    const toolbarProps = props();
    act(() => root.render(createElement(WordToolbar, toolbarProps)));
    click(control('字体颜色'));
    click(document.body.querySelector('[role="dialog"] button[aria-label="#D92D20"]'));
    expect(toolbarProps.onColorFormat).toHaveBeenCalledWith({ color: '#D92D20' }, 'color');
    click(control('文本背景颜色（突出显示）'));
    click(document.body.querySelector('[role="dialog"] button[aria-label="#F79009"]'));
    expect(toolbarProps.onColorFormat).toHaveBeenCalledWith({ highlight: '#F79009' }, 'highlight');
  });

  it('插入面板保留表格尺寸、图片、分页符与页眉页脚回调', () => {
    const toolbarProps = props();
    act(() => root.render(createElement(WordToolbar, toolbarProps)));
    selectTab('插入');
    click(control('插入图片'));
    click(control('插入分页符'));
    click(control('页眉、页脚与页码'));
    expect(toolbarProps.onInsertImage).toHaveBeenCalledOnce();
    expect(toolbarProps.onPageBreak).toHaveBeenCalledOnce();
    expect(toolbarProps.onHeaderFooter).toHaveBeenCalledOnce();
    click(control('插入表格'));
    click(document.body.querySelector('[data-table-cell][aria-label="2 列 3 行"]'));
    expect(toolbarProps.onInsertTable).toHaveBeenCalledWith(3, 2);
  });

  it('布局、邮件、视图只显示对应已实现功能，并维持各自回调', () => {
    const toolbarProps = props({ onToggleFormattingMarks: vi.fn(), showFormattingMarks: true });
    act(() => root.render(createElement(WordToolbar, toolbarProps)));
    click(control('显示编辑标记'));
    expect(toolbarProps.onToggleFormattingMarks).toHaveBeenCalledOnce();
    selectTab('布局');
    click(control('布局与属性'));
    click(control('段落设置'));
    expect(toolbarProps.onToggleProperties).toHaveBeenCalledOnce();
    expect(toolbarProps.onOpenParagraphDialog).toHaveBeenCalledOnce();
    selectTab('邮件');
    click(control('邮件合并'));
    expect(toolbarProps.onMailMerge).toHaveBeenCalledOnce();
    selectTab('视图');
    expect(control('标尺')?.getAttribute('aria-pressed')).toBe('true');
    expect(control('显示编辑标记')?.getAttribute('aria-pressed')).toBe('true');
    click(control('显示编辑标记'));
    click(control('显示属性面板'));
    click(control('打印'));
    click(control('导出 PDF'));
    click(control('兼容性报告'));
    expect(toolbarProps.onToggleFormattingMarks).toHaveBeenCalledTimes(2);
    expect(toolbarProps.onToggleProperties).toHaveBeenCalledTimes(2);
    expect(toolbarProps.onPrint).toHaveBeenCalledOnce();
    expect(toolbarProps.onExport).toHaveBeenCalledOnce();
    expect(toolbarProps.onShowCompatibility).toHaveBeenCalledOnce();
  });

  it('更多菜单在折叠时仍可查找、打开布局、邮件、图片与兼容报告', () => {
    const toolbarProps = props();
    act(() => root.render(createElement(WordToolbar, toolbarProps)));
    click(control('折叠功能区'));
    for (const [label, callback] of [
      ['查找替换', toolbarProps.onToggleFind],
      ['布局与属性', toolbarProps.onToggleProperties],
      ['邮件合并 / 批量报告', toolbarProps.onMailMerge],
      ['插入图片', toolbarProps.onInsertImage],
      ['兼容性报告', toolbarProps.onShowCompatibility],
    ] as const) {
      click(control('更多操作'));
      click(Array.from(document.body.querySelectorAll('[role="menuitem"]')).find((item) => item.textContent?.includes(label)));
      expect(callback).toHaveBeenCalledOnce();
      expect(visiblePanel()).toBeNull();
    }
  });

  it('只读仍可切换选项卡和查看，编辑入口及更多菜单编辑项保持禁用', () => {
    act(() => root.render(createElement(WordToolbar, props({ readonly: true, onToggleFormattingMarks: vi.fn() }))));
    for (const label of ['加粗', '斜体', '下划线', '删除线', '字体颜色', '文本背景颜色（突出显示）', '项目符号', '编号', '增加缩进', '减少缩进', '段落']) {
      expect(control(label)?.disabled).toBe(true);
    }
    expect(container.querySelector<HTMLSelectElement>('select[aria-label="字体"]')?.disabled).toBe(true);
    selectTab('插入');
    for (const label of ['插入表格', '插入图片', '插入分页符', '页眉、页脚与页码']) expect(control(label)?.disabled).toBe(true);
    selectTab('邮件');
    expect(control('邮件合并')?.disabled).toBe(true);
    selectTab('视图');
    for (const label of ['标尺', '显示编辑标记', '打印', '导出 PDF', '兼容性报告']) expect(control(label)?.disabled).toBe(false);
    click(control('更多操作'));
    const image = Array.from(document.body.querySelectorAll<HTMLButtonElement>('[role="menuitem"]')).find((item) => item.textContent?.includes('插入图片'));
    expect(image?.disabled).toBe(true);
  });

  it('选项卡、折叠及面板按钮沿用选择保护，字体选择不阻止原生交互', () => {
    const onBeforeInteraction = vi.fn();
    act(() => root.render(createElement(WordToolbar, props({ onBeforeInteraction }))));
    const tab = container.querySelector<HTMLButtonElement>('[role="tab"]')!;
    for (const button of [tab, control('折叠功能区')!, control('加粗')!]) {
      const mouseDown = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
      act(() => {
        button.dispatchEvent(new Event('pointerdown', { bubbles: true }));
        button.dispatchEvent(mouseDown);
      });
      expect(mouseDown.defaultPrevented).toBe(true);
    }
    const font = container.querySelector('select[aria-label="字体"]')!;
    const mouseDown = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
    act(() => {
      font.dispatchEvent(new Event('pointerdown', { bubbles: true }));
      font.dispatchEvent(mouseDown);
    });
    expect(mouseDown.defaultPrevented).toBe(false);
    expect(onBeforeInteraction).toHaveBeenCalledTimes(4);
  });

  it('开始面板保留样式库配置与开关回调，并服从工具栏只读状态', () => {
    const onOpenChange = vi.fn();
    const stylePicker: NonNullable<WordToolbarProps['stylePicker']> = {
      currentStyleName: '正文', open: false, onOpenChange, readonly: false,
      presets: [], documentStyles: [], showMore: false, defaultFont: '宋体', defaultFontSizePt: 12, styleSets: [],
      onApply: vi.fn(), onClearFormatting: vi.fn(), onNewStyle: vi.fn(), onToggleMore: vi.fn(),
      onDeleteStyle: vi.fn(), onSaveStyleSet: vi.fn(), onApplyStyleSet: vi.fn(), onDeleteStyleSet: vi.fn(),
    };
    act(() => root.render(createElement(WordToolbar, props({ stylePicker }))));
    expect(visiblePanel()?.querySelector('[role="group"][aria-label="样式"]')).toBeTruthy();
    expect(control('样式库')?.textContent).toContain('正文');
    click(control('样式库'));
    expect(onOpenChange).toHaveBeenCalledWith(true);
    act(() => root.render(createElement(WordToolbar, props({ stylePicker, readonly: true }))));
    expect(control('样式库')?.disabled).toBe(true);
  });
});
