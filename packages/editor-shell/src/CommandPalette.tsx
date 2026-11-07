import { useEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import {
  AppWindow,
  ArrowRight,
  Download,
  Eye,
  FilePlus2,
  FileText,
  FileWarning,
  FolderOpen,
  Home,
  Keyboard,
  Monitor,
  Moon,
  PanelRight,
  Play,
  Presentation,
  Printer,
  Puzzle,
  Ruler,
  Save,
  SaveAll,
  Search,
  Settings,
  Sparkles,
  Sun,
  Table2,
  X,
} from 'lucide-react';
import { useWorkspaceStore } from '@mini-office/document-model';
import { collectMacroCommands, usePluginRegistry } from '@mini-office/plugin-host';
import { Kbd, FileIcon } from '@mini-office/ui';
import { MOD_KEY, type OfficeKind } from '@mini-office/shared';
import { ACCENT_PRESETS, applyAppearance } from '@mini-office/design-system';
import { canRunMacro, executePluginMacro } from './plugin-macro';
import { collectShortcutBindings, formatShortcut } from './shortcuts';

interface Command {
  id: string;
  label: string;
  section: string;
  icon: React.ReactNode;
  shortcut?: string;
  /** 右侧补充说明，用于"快捷键冲突""不适用"这类需要解释的禁用原因。 */
  note?: string;
  keywords?: string;
  disabled?: boolean;
  run: () => void;
}

export function CommandPalette({
  open,
  onClose,
  onHome,
  onRequestClose,
}: {
  open: boolean;
  onClose: () => void;
  onHome: () => void;
  onRequestClose: (id: string) => void;
}) {
  const sessions = useWorkspaceStore((state) => state.sessions);
  const activeId = useWorkspaceStore((state) => state.activeId);
  const activeSession = activeId ? sessions[activeId] : undefined;
  const createDocument = useWorkspaceStore((state) => state.createDocument);
  const openFileDialog = useWorkspaceStore((state) => state.openFileDialog);
  const saveActive = useWorkspaceStore((state) => state.saveActive);
  const openDialog = useWorkspaceStore((state) => state.openDialog);
  const updateSettings = useWorkspaceStore((state) => state.updateSettings);
  const settings = useWorkspaceStore((state) => state.settings);
  const recent = useWorkspaceStore((state) => state.recent);
  const openSnapshot = useWorkspaceStore((state) => state.openSnapshot);
  // 订阅 installed 而不是调用 macroCommands()，这样安装/禁用插件后命令面板立即刷新。
  const installedPlugins = usePluginRegistry((state) => state.installed);

  const [query, setQuery] = useState('');
  const [activeIndex, setActiveIndex] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    setQuery('');
    setActiveIndex(0);
    const timer = setTimeout(() => inputRef.current?.focus(), 20);
    return () => clearTimeout(timer);
  }, [open]);

  const commands = useMemo<Command[]>(() => {
    const fire = (name: string) => window.dispatchEvent(new CustomEvent(name));
    const create = async (kind: OfficeKind) => {
      await createDocument(kind);
      onHome();
    };
    const applyThemeMode = (theme: 'light' | 'dark' | 'system') => {
      applyAppearance({ theme, accent: settings.accent, customAccent: settings.customAccent, eyeCare: settings.eyeCare });
      updateSettings({ theme });
    };
    const list: Command[] = [
      {
        id: 'open',
        label: '打开文件…',
        section: '文件',
        icon: <FolderOpen size={15} />,
        shortcut: `${MOD_KEY}O`,
        run: () => void openFileDialog(),
      },
      {
        id: 'new-docx',
        label: '新建 Word 文档',
        section: '文件',
        icon: <FileText size={15} />,
        shortcut: `${MOD_KEY}N`,
        run: () => void create('docx'),
      },
      {
        id: 'new-xlsx',
        label: '新建 Excel 表格',
        section: '文件',
        icon: <Table2 size={15} />,
        run: () => void create('xlsx'),
      },
      {
        id: 'home',
        label: '返回首页',
        section: '文件',
        icon: <Home size={15} />,
        run: onHome,
      },
      {
        id: 'save',
        label: '保存',
        section: '文件',
        icon: <Save size={15} />,
        shortcut: `${MOD_KEY}S`,
        keywords: 'save',
        disabled: !activeSession || activeSession.readonly,
        run: () => void saveActive(),
      },
      {
        id: 'save-as',
        label: '另存为副本',
        section: '文件',
        icon: <SaveAll size={15} />,
        shortcut: `⇧${MOD_KEY}S`,
        keywords: 'save as',
        disabled: !activeSession || activeSession.readonly,
        run: () => void saveActive({ asNew: true }),
      },
      {
        id: 'export-pdf',
        label: '导出 PDF',
        section: '文件',
        icon: <Download size={15} />,
        keywords: 'pdf export',
        disabled: !activeSession,
        run: () => fire('mo:export-pdf'),
      },
      {
        id: 'print',
        label: '打印',
        section: '文件',
        icon: <Printer size={15} />,
        shortcut: `${MOD_KEY}P`,
        keywords: 'print',
        disabled: !activeSession,
        run: () => fire('mo:print'),
      },
      {
        id: 'close-tab',
        label: '关闭当前标签',
        section: '文件',
        icon: <X size={15} />,
        shortcut: `${MOD_KEY}W`,
        disabled: !activeId,
        run: () => {
          if (activeId) onRequestClose(activeId);
        },
      },
      {
        id: 'find',
        label: '查找与替换',
        section: '编辑',
        icon: <Search size={15} />,
        shortcut: `${MOD_KEY}F`,
        disabled: !activeSession,
        run: () => fire('mo:find'),
      },
      {
        // 功能：样式库是整段排版的入口，放在编辑分组便于搜索到；时间：2026-10-29；作者：lq。
        id: 'word-styles',
        label: '样式库（预设样式与我的样式集）',
        section: '编辑',
        icon: <Sparkles size={15} />,
        keywords: 'style 样式 标题 正文 排版',
        disabled: activeSession?.kind !== 'docx',
        run: () => fire('mo:styles'),
      },
      {
        id: 'present',
        label: '进入全屏演示',
        section: '编辑',
        icon: <Play size={15} />,
        shortcut: 'F5',
        disabled: activeSession?.kind !== 'pptx',
        run: () => fire('mo:present'),
      },
      {
        id: 'properties',
        label: settings.showProperties ? '隐藏属性面板' : '显示属性面板',
        section: '视图',
        icon: <PanelRight size={15} />,
        run: () => updateSettings({ showProperties: !settings.showProperties }),
      },
      // 功能：标尺默认隐藏，命令面板提供显式入口，找不到工具栏按钮时也能开启；时间：2026-10-30；作者：lq。
      {
        id: 'word-ruler',
        label: settings.wordRuler === true ? '隐藏标尺' : '显示标尺',
        section: '视图',
        icon: <Ruler size={15} />,
        disabled: activeSession?.kind !== 'docx',
        run: () => updateSettings({ wordRuler: settings.wordRuler !== true }),
      },
      {
        id: 'eye-care',
        label: settings.eyeCare ? '关闭护眼模式' : '开启护眼模式',
        section: '视图',
        icon: <Eye size={15} />,
        run: () => {
          const eyeCare = !settings.eyeCare;
          applyAppearance({
            theme: settings.theme,
            accent: settings.accent,
            customAccent: settings.customAccent,
            eyeCare,
          });
          updateSettings({ eyeCare });
        },
      },
      {
        id: 'theme-light',
        label: '切换到浅色主题',
        section: '视图',
        icon: <Sun size={15} />,
        run: () => applyThemeMode('light'),
      },
      {
        id: 'theme-dark',
        label: '切换到深色主题',
        section: '视图',
        icon: <Moon size={15} />,
        run: () => applyThemeMode('dark'),
      },
      {
        id: 'theme-system',
        label: '跟随系统主题',
        section: '视图',
        icon: <Monitor size={15} />,
        run: () => applyThemeMode('system'),
      },
      ...ACCENT_PRESETS.map((preset) => ({
        id: `accent-${preset.id}`,
        label: `主题色：${preset.label}`,
        section: '视图',
        icon: (
          <span
            className="inline-block h-3.5 w-3.5 rounded-full ring-1 ring-black/10"
            style={{ background: preset.light.accent }}
          />
        ),
        keywords: `accent color ${preset.id}`,
        run: () => {
          applyAppearance({
            theme: settings.theme,
            accent: preset.id,
            customAccent: settings.customAccent,
            eyeCare: settings.eyeCare,
          });
          updateSettings({ accent: preset.id });
        },
      })),
      {
        id: 'settings',
        label: '打开设置',
        section: '视图',
        icon: <Settings size={15} />,
        shortcut: `${MOD_KEY},`,
        run: () => openDialog({ type: 'settings' }),
      },
      {
        id: 'plugins',
        label: '管理插件',
        section: '视图',
        icon: <Puzzle size={15} />,
        keywords: 'plugin 扩展 主题 安装',
        run: () => openDialog({ type: 'plugins' }),
      },
      // 功能：命令面板直达默认应用设置；时间：2026-10-29；作者：lq。
      {
        id: 'default-app',
        label: '设为默认应用',
        section: '视图',
        icon: <AppWindow size={15} />,
        keywords: 'default 默认应用 文件关联 双击打开',
        run: () => openDialog({ type: 'default-app' }),
      },
      {
        id: 'compatibility',
        label: '查看兼容性报告',
        section: '帮助',
        icon: <FileWarning size={15} />,
        disabled: !activeSession,
        run: () => {
          if (activeSession) openDialog({ type: 'compatibility', sessionId: activeSession.id });
        },
      },
      {
        id: 'shortcuts',
        label: '键盘快捷键',
        section: '帮助',
        icon: <Keyboard size={15} />,
        run: () => openDialog({ type: 'shortcuts' }),
      },
    ];
    const recentCommands: Command[] = recent.slice(0, 3).map((item) => ({
      id: `recent-${item.id}`,
      label: `打开最近：${item.displayName}`,
      section: '最近',
      // 功能：最近文档命令复用类型图标；时间：2026-11-07；作者：lq。
      icon: <FileIcon kind={item.kind} name={item.displayName} size={22} />,
      run: () => void openSnapshot(item.id),
    }));
    // 插件宏：只列当前文档类型能跑的，避免在 Word 里出现 Excel 命令。
    // 快捷键走统一的冲突检测，冲突项禁用并写明原因——展示一个按不动的快捷键比不展示更糟。
    const pluginCommands: Command[] = collectShortcutBindings(collectMacroCommands(installedPlugins)).map((binding) => ({
      id: `plugin-${binding.macro.key}`,
      label: binding.macro.command.title,
      section: `插件 · ${binding.macro.pluginName}`,
      icon: <Sparkles size={15} />,
      shortcut: binding.shortcut ? formatShortcut(binding.shortcut) : undefined,
      note: binding.conflict ?? undefined,
      keywords: `${binding.macro.command.keywords ?? ''} plugin 插件 宏 ${binding.macro.pluginName}`,
      disabled: !!binding.conflict || !canRunMacro(activeSession, binding.macro),
      run: () => {
        executePluginMacro(binding.macro);
      },
    }));
    return [...list, ...pluginCommands, ...recentCommands];
  }, [
    activeId,
    activeSession,
    createDocument,
    installedPlugins,
    onHome,
    onRequestClose,
    openDialog,
    openFileDialog,
    openSnapshot,
    recent,
    saveActive,
    settings.accent,
    settings.customAccent,
    settings.eyeCare,
    settings.showProperties,
    settings.theme,
    settings.wordRuler,
    updateSettings,
  ]);

  const filtered = useMemo(() => {
    const keyword = query.trim().toLowerCase();
    if (!keyword) return commands;
    return commands.filter((command) =>
      `${command.label} ${command.section} ${command.keywords ?? ''}`.toLowerCase().includes(keyword),
    );
  }, [commands, query]);

  useEffect(() => {
    setActiveIndex((index) => Math.min(index, Math.max(0, filtered.length - 1)));
  }, [filtered.length]);

  useEffect(() => {
    if (!open) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        onClose();
      } else if (event.key === 'ArrowDown') {
        event.preventDefault();
        setActiveIndex((index) => (filtered.length ? (index + 1) % filtered.length : 0));
      } else if (event.key === 'ArrowUp') {
        event.preventDefault();
        setActiveIndex((index) => (filtered.length ? (index - 1 + filtered.length) % filtered.length : 0));
      } else if (event.key === 'Enter') {
        event.preventDefault();
        const command = filtered[activeIndex];
        if (command && !command.disabled) {
          onClose();
          command.run();
        }
      }
    };
    document.addEventListener('keydown', onKeyDown, true);
    return () => document.removeEventListener('keydown', onKeyDown, true);
  }, [activeIndex, filtered, onClose, open]);

  useEffect(() => {
    listRef.current
      ?.querySelector<HTMLElement>(`[data-command-index="${activeIndex}"]`)
      ?.scrollIntoView({ block: 'nearest' });
  }, [activeIndex]);

  if (!open || typeof document === 'undefined') return null;

  const grouped = new Map<string, Array<{ command: Command; index: number }>>();
  filtered.forEach((command, index) => {
    const list = grouped.get(command.section) ?? [];
    list.push({ command, index });
    grouped.set(command.section, list);
  });

  return createPortal(
    <div className="fixed inset-0 z-[100] flex justify-center px-4 pt-[12vh]">
      <div className="mo-scrim" onClick={onClose} aria-hidden />
      <div className="mo-glass mo-pop-in relative z-[101] flex max-h-[66vh] w-full max-w-[560px] flex-col overflow-hidden rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface)] shadow-pop">
        <div className="flex items-center gap-2 border-b border-[var(--mo-line)] px-3.5 py-2.5">
          <Search size={15} className="shrink-0 text-[var(--mo-text-muted)]" />
          <input
            ref={inputRef}
            value={query}
            onChange={(event) => {
              setQuery(event.target.value);
              setActiveIndex(0);
            }}
            placeholder="搜索命令、文件或操作…"
            className="h-7 min-w-0 flex-1 bg-transparent text-[13px] outline-none placeholder:text-[var(--mo-text-faint)]"
          />
          <button
            type="button"
            onClick={onClose}
            className="flex h-6 w-6 items-center justify-center rounded text-[var(--mo-text-faint)] hover:bg-[var(--mo-surface-2)]"
            aria-label="关闭命令面板"
          >
            <X size={14} />
          </button>
        </div>
        <div ref={listRef} className="mo-scroll min-h-0 flex-1 p-1.5">
          {filtered.length === 0 ? (
            <div className="px-4 py-10 text-center text-[12px] text-[var(--mo-text-muted)]">
              没有找到匹配的命令
            </div>
          ) : (
            [...grouped.entries()].map(([section, items]) => (
              <div key={section} className="mb-1">
                <div className="px-2 pb-1 pt-1.5 text-[9px] font-semibold uppercase tracking-wide text-[var(--mo-text-faint)]">
                  {section}
                </div>
                {items.map(({ command, index }) => (
                  <button
                    key={command.id}
                    type="button"
                    data-command-index={index}
                    disabled={command.disabled}
                    onMouseEnter={() => setActiveIndex(index)}
                    onClick={() => {
                      if (command.disabled) return;
                      onClose();
                      command.run();
                    }}
                    className={`flex w-full items-center gap-2.5 rounded-md px-2 py-1.5 text-left text-[12px] transition-colors ${
                      index === activeIndex
                        ? 'bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]'
                        : 'text-[var(--mo-text-primary)]'
                    } ${command.disabled ? 'cursor-not-allowed opacity-35' : ''}`}
                  >
                    <span className="flex h-5 w-5 shrink-0 items-center justify-center">{command.icon}</span>
                    <span className="min-w-0 flex-1 truncate">{command.label}</span>
                    {command.note ? (
                      <span className="shrink-0 text-[9px] text-[var(--mo-warning)]">{command.note}</span>
                    ) : command.shortcut ? (
                      <span className="shrink-0 text-[10px] text-[var(--mo-text-faint)]">{command.shortcut}</span>
                    ) : (
                      <ArrowRight size={11} className="shrink-0 text-[var(--mo-text-faint)] opacity-0" />
                    )}
                  </button>
                ))}
              </div>
            ))
          )}
        </div>
        <div className="flex items-center justify-between border-t border-[var(--mo-line)] bg-[var(--mo-surface-2)]/50 px-3 py-1.5 text-[9px] text-[var(--mo-text-muted)]">
          <div className="flex items-center gap-2">
            <span className="inline-flex items-center gap-0.5"><Kbd>↑</Kbd><Kbd>↓</Kbd> 选择</span>
            <span className="inline-flex items-center gap-0.5"><Kbd>↵</Kbd> 执行</span>
            <span className="inline-flex items-center gap-0.5"><Kbd>Esc</Kbd> 关闭</span>
          </div>
          <span>{filtered.length} 个命令</span>
        </div>
      </div>
    </div>,
    document.body,
  );
}
