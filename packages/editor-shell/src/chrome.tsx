import { BrandIcon } from '@mini-office/ui';
import { useEffect, useState, type ReactNode } from 'react';
import {
  AlertTriangle,
  AppWindow,
  Copy,
  Download,
  Eye,
  Home,
  FileText,
  FolderOpen,
  Layers,
  Monitor,
  Moon,
  MoreHorizontal,
  Minus,
  Palette,
  Plus,
  Presentation,
  Search,
  Settings,
  Square,
  Sun,
  Table2,
  X,
  FileCode2,
  FileType2,
  FileKey,
} from 'lucide-react';
import { useWorkspaceStore, type DocumentSession } from '@mini-office/document-model';
import { Button, Dialog, Dropdown, IconButton, Kbd, Select, Tooltip, type MenuItem } from '@mini-office/ui';
import { coverageLabel } from '@mini-office/document-model';
import {
  MOD_KEY,
  OFFICE_META,
  formatDateTime,
  type AccentId,
  type OfficeKind,
} from '@mini-office/shared';
import { ACCENT_PRESETS, applyAppearance, resolveTheme } from '@mini-office/design-system';
import {
  closeWindow,
  isTauriRuntime,
  isWindowMaximized,
  isWindows,
  minimizeWindow,
  startWindowDrag,
  supportsWindowVibrancy,
  toggleWindowMaximize,
} from '@mini-office/platform-adapter';
import { useUpdateStore } from './update-check';

const KIND_ICONS: Record<OfficeKind, typeof FileText> = {
  docx: FileText,
  xlsx: Table2,
  pptx: Presentation,
  pdf: FileKey,
  md: FileType2,
  html: FileCode2,
};

const KIND_TINTS: Record<OfficeKind, string> = {
  docx: '#2B579A',
  xlsx: '#217346',
  pptx: '#C43E1C',
  pdf: '#E34F4F',
  md: '#6B5BD2',
  html: '#E67E22',
};

export function TitleBar({
  onHome,
  onOpenPalette,
  tabs,
  merged = false,
}: {
  onHome: () => void;
  onOpenPalette: () => void;
  /** 功能：Windows 顶栏合并时把文档标签放进标题栏同一行；时间：2026-11-01；作者：lq。 */
  tabs?: ReactNode;
  merged?: boolean;
}) {
  const settings = useWorkspaceStore((state) => state.settings);
  const updateSettings = useWorkspaceStore((state) => state.updateSettings);
  const openDialog = useWorkspaceStore((state) => state.openDialog);
  const activeId = useWorkspaceStore((state) => state.activeId);
  // 功能：有新版本时在「更多」入口亮小红点；时间：2026-10-29；作者：lq。
  const updateAvailable = useUpdateStore((state) => state.latest !== null);
  const latestTag = useUpdateStore((state) => state.latest?.tagName ?? null);

  const currentTheme = resolveTheme(settings.theme);
  const canVibrancy = supportsWindowVibrancy();
  const applyThemeMode = (theme: 'light' | 'dark' | 'system') => {
    applyAppearance({
      theme,
      accent: settings.accent,
      customAccent: settings.customAccent,
      // 功能：切换主题时保留护眼状态，避免外壳与纸张配色错位；时间：2026-10-30；作者：lq。
      eyeCare: settings.eyeCare,
    });
    updateSettings({ theme });
  };
  const toggleTheme = () => {
    applyThemeMode(currentTheme === 'dark' ? 'light' : 'dark');
  };
  const applyAccentMode = (accent: AccentId) => {
    applyAppearance({
      theme: settings.theme,
      accent,
      customAccent: settings.customAccent,
      eyeCare: settings.eyeCare,
    });
    updateSettings({ accent });
  };

  const nativeOverlay =
    isTauriRuntime() && typeof navigator !== 'undefined' && /Mac/i.test(navigator.platform);
  const isInteractive = (target: EventTarget | null) =>
    target instanceof Element &&
    !!target.closest('button, a, input, select, textarea, [role="button"], [data-no-drag]');

  return (
    <header
      className={`mo-glass mo-no-print flex shrink-0 select-none items-center gap-2 border-b border-[var(--mo-line)] bg-[var(--mo-surface)] px-2.5 ${
        merged ? 'h-10 gap-1.5 pr-0' : 'h-11'
      }`}
      style={nativeOverlay ? { paddingLeft: 78 } : undefined}
      onMouseDown={(event) => {
        if (event.button !== 0 || isInteractive(event.target)) return;
        event.preventDefault();
        void startWindowDrag();
      }}
      onDoubleClick={(event) => {
        if (isInteractive(event.target)) return;
        void toggleWindowMaximize();
      }}
    >
      <button
        type="button"
        onClick={onHome}
        className="flex shrink-0 items-center gap-2 rounded-md px-1.5 py-1 hover:bg-[var(--mo-surface-2)]"
        title="返回首页"
      >
        {/* 功能：标题栏统一品牌图标；时间：2026-10-18；作者：lq。 */}
        <BrandIcon size={merged ? 22 : 30} />
        {merged ? null : <span className="hidden text-[12px] font-semibold sm:block">Mini Office</span>}
      </button>
      {merged ? null : <div className="h-5 w-px bg-[var(--mo-line)]" />}
      {merged && tabs ? (
        // 功能：标签与标题栏合并成一行，节省 Windows 端一条纵向空间；时间：2026-11-01；作者：lq。
        <div className="flex min-w-0 flex-1 items-stretch self-stretch">{tabs}</div>
      ) : (
        <div className="flex min-w-0 flex-1 items-center px-1">
          {activeId ? <ActiveDocumentLabel /> : <span className="text-[11px] text-[var(--mo-text-muted)]">首页</span>}
        </div>
      )}
      <button
        type="button"
        onClick={onOpenPalette}
        className={`flex h-7 shrink-0 items-center gap-2 rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/60 px-2.5 text-[11px] text-[var(--mo-text-muted)] transition-colors hover:border-[var(--mo-line-strong)] hover:text-[var(--mo-text-primary)] ${
          merged ? 'max-w-[150px]' : ''
        }`}
      >
        <Search size={13} />
        <span className="hidden md:block">搜索命令</span>
        <span className="hidden items-center gap-0.5 md:flex">
          <Kbd>{MOD_KEY}</Kbd>
          <Kbd>K</Kbd>
        </span>
      </button>
      <Tooltip label={`主题：${settings.theme === 'system' ? '跟随系统' : settings.theme === 'dark' ? '深色' : '浅色'}`}>
        <IconButton label="切换主题" onClick={toggleTheme}>
          {currentTheme === 'dark' ? <Moon size={15} /> : <Sun size={15} />}
        </IconButton>
      </Tooltip>
      {/* 功能：护眼模式一键开关，放在主题切换旁边便于随手切换；时间：2026-10-30；作者：lq。 */}
      <Tooltip label={settings.eyeCare ? '关闭护眼模式' : '护眼模式：米黄纸色，低蓝光'}>
        <IconButton
          label="护眼模式"
          active={settings.eyeCare}
          onClick={() => updateSettings({ eyeCare: !settings.eyeCare })}
        >
          <Eye size={15} />
        </IconButton>
      </Tooltip>
      <Dropdown
        align="right"
        minWidth={190}
        items={[
          { type: 'label', label: '主题色' },
          ...ACCENT_PRESETS.map((preset) => ({
            id: `accent-${preset.id}`,
            label: preset.label,
            icon: (
              <span
                className="h-3.5 w-3.5 rounded-full ring-1 ring-black/10"
                style={{ background: preset[currentTheme].accent }}
              />
            ),
            checked: settings.accent === preset.id,
            onSelect: () => applyAccentMode(preset.id),
          })),
          { type: 'separator' },
          {
            id: 'accent-custom',
            label: '自定义主题色…',
            icon: <Palette size={14} />,
            shortcut: `${MOD_KEY},`,
            onSelect: () => openDialog({ type: 'settings' }),
          },
          { type: 'separator' },
          ...(canVibrancy
            ? [
                {
                  id: 'vibrancy',
                  label: '毛玻璃效果',
                  icon: <Layers size={14} />,
                  checked: settings.vibrancy,
                  onSelect: () => updateSettings({ vibrancy: !settings.vibrancy }),
                },
              ]
            : []),
        ]}
        trigger={({ toggle }) => (
          <IconButton label="主题色" onClick={toggle}>
            <Palette size={15} />
          </IconButton>
        )}
      />
      <Dropdown
        align="right"
        minWidth={190}
        items={[
          // 功能：有新版本时置顶一条入口，与小红点呼应；时间：2026-10-29；作者：lq。
          ...(updateAvailable
            ? ([
                {
                  id: 'update-available',
                  label: `发现新版本 ${latestTag}`,
                  icon: <Download size={14} />,
                  onSelect: () => openDialog({ type: 'update' }),
                },
                { type: 'separator' },
              ] satisfies MenuItem[])
            : []),
          {
            id: 'settings',
            label: '设置',
            icon: <Settings size={14} />,
            shortcut: `${MOD_KEY},`,
            onSelect: () => openDialog({ type: 'settings' }),
          },
          // 功能：菜单直达默认应用设置，桌面端可从系统层接管文档双击打开；时间：2026-10-29；作者：lq。
          {
            id: 'default-app',
            label: '设为默认应用…',
            icon: <AppWindow size={14} />,
            onSelect: () => openDialog({ type: 'default-app' }),
          },
          {
            id: 'theme-system',
            label: '跟随系统主题',
            icon: <Monitor size={14} />,
            checked: settings.theme === 'system',
            onSelect: () => applyThemeMode('system'),
          },
          {
            id: 'theme-light',
            label: '浅色主题',
            icon: <Sun size={14} />,
            checked: settings.theme === 'light',
            onSelect: () => applyThemeMode('light'),
          },
          {
            id: 'theme-dark',
            label: '深色主题',
            icon: <Moon size={14} />,
            checked: settings.theme === 'dark',
            onSelect: () => applyThemeMode('dark'),
          },
          ...(canVibrancy
            ? [
                {
                  id: 'vibrancy',
                  label: '毛玻璃效果',
                  icon: <Layers size={14} />,
                  checked: settings.vibrancy,
                  onSelect: () => updateSettings({ vibrancy: !settings.vibrancy }),
                },
              ]
            : []),
          { type: 'separator' },
          {
            id: 'shortcuts',
            label: '键盘快捷键',
            onSelect: () => openDialog({ type: 'shortcuts' }),
          },
          {
            id: 'about',
            label: '关于 Mini Office',
            onSelect: () => openDialog({ type: 'about' }),
          },
        ]}
        trigger={({ toggle }) => (
          <IconButton label={updateAvailable ? '更多（有新版本）' : '更多'} onClick={toggle}>
            <span className="relative flex items-center justify-center">
              <MoreHorizontal size={16} />
              {updateAvailable ? (
                <span
                  data-update-dot
                  className="absolute -right-1.5 -top-1 h-2 w-2 rounded-full bg-[var(--mo-danger)] ring-2 ring-[var(--mo-surface)]"
                />
              ) : null}
            </span>
          </IconButton>
        )}
      />
      {merged ? <WindowControls /> : null}
    </header>
  );
}

/**
 * 功能：自绘窗口按钮（最小化 / 最大化 / 关闭），仅在去掉系统标题栏的 Windows 桌面端渲染；
 * 时间：2026-11-01；作者：lq。
 */
function WindowControls() {
  const [maximized, setMaximized] = useState(false);
  useEffect(() => {
    let alive = true;
    const sync = () => { void isWindowMaximized().then((value) => { if (alive) setMaximized(value); }); };
    sync();
    window.addEventListener('resize', sync);
    return () => { alive = false; window.removeEventListener('resize', sync); };
  }, []);
  const buttonClass =
    'flex h-10 w-[42px] shrink-0 items-center justify-center text-[var(--mo-text-secondary)] transition-colors';
  return (
    <div className="flex h-full shrink-0 items-stretch self-stretch" data-no-drag>
      <button
        type="button"
        aria-label="最小化"
        title="最小化"
        className={`${buttonClass} hover:bg-[var(--mo-surface-2)] hover:text-[var(--mo-text-primary)]`}
        onMouseDown={(event) => event.stopPropagation()}
        onClick={() => void minimizeWindow()}
      >
        <Minus size={14} />
      </button>
      <button
        type="button"
        aria-label={maximized ? '向下还原' : '最大化'}
        title={maximized ? '向下还原' : '最大化'}
        className={`${buttonClass} hover:bg-[var(--mo-surface-2)] hover:text-[var(--mo-text-primary)]`}
        onMouseDown={(event) => event.stopPropagation()}
        onClick={() => { void toggleWindowMaximize().then(() => isWindowMaximized().then(setMaximized)); }}
      >
        {maximized ? <Copy size={12} /> : <Square size={12} />}
      </button>
      <button
        type="button"
        aria-label="关闭"
        title="关闭"
        className={`${buttonClass} hover:bg-[#C42B1C] hover:text-white`}
        onMouseDown={(event) => event.stopPropagation()}
        onClick={() => void closeWindow()}
      >
        <X size={15} />
      </button>
    </div>
  );
}

function ActiveDocumentLabel() {
  const sessions = useWorkspaceStore((state) => state.sessions);
  const activeId = useWorkspaceStore((state) => state.activeId);
  const session = activeId ? sessions[activeId] : undefined;
  if (!session) return null;
  return (
    <div className="flex min-w-0 items-center gap-2">
      <span className="truncate text-[12px] text-[var(--mo-text-primary)]">{session.displayName}</span>
      {session.dirty ? <span className="h-1.5 w-1.5 shrink-0 rounded-full bg-[var(--mo-warning)]" /> : null}
      {session.readonly ? (
        <span className="rounded bg-[var(--mo-surface-2)] px-1 py-0.5 text-[9px] text-[var(--mo-text-muted)]">
          只读
        </span>
      ) : null}
    </div>
  );
}

export function TabBar({
  onRequestClose,
  onHome,
  onSelectTab,
}: {
  onRequestClose: (id: string) => void;
  onHome: () => void;
  onSelectTab: (id: string) => void;
}) {
  const openFileDialog = useWorkspaceStore((state) => state.openFileDialog);

  return (
    <div
      className="mo-glass mo-no-print flex h-[38px] shrink-0 select-none items-stretch border-b border-[var(--mo-line)] bg-[var(--mo-bg)]"
      onMouseDown={(event) => {
        if (event.button !== 0) return;
        const target = event.target;
        if (
          target instanceof Element &&
          target.closest('button, [role="tab"], input, select, textarea, a, [role="button"]')
        ) {
          return;
        }
        event.preventDefault();
        void startWindowDrag();
      }}
    >
      <button
        type="button"
        onClick={onHome}
        aria-label="首页"
        className="flex w-9 shrink-0 items-center justify-center border-r border-[var(--mo-line)] text-[var(--mo-text-muted)] hover:bg-[var(--mo-surface-2)] hover:text-[var(--mo-text-primary)]"
      >
        <Home size={14} />
      </button>
      <TabStrip onRequestClose={onRequestClose} onSelectTab={onSelectTab} />
      <div className="flex shrink-0 items-center gap-1 border-l border-[var(--mo-line)] px-2">
        <button
          type="button"
          onClick={() => void openFileDialog()}
          className="hidden h-7 items-center gap-1.5 rounded-md px-2 text-[11px] text-[var(--mo-text-muted)] hover:bg-[var(--mo-surface-2)] sm:flex"
        >
          <FolderOpen size={13} /> 打开
        </button>
      </div>
    </div>
  );
}

/**
 * 功能：文档标签列表，标题栏合并与独立标签栏共用同一套拖拽排序和关闭逻辑；
 * 时间：2026-11-01；作者：lq。
 */
export function TabStrip({
  onRequestClose,
  onSelectTab,
  compact = false,
}: {
  onRequestClose: (id: string) => void;
  onSelectTab: (id: string) => void;
  /** 功能：并入标题栏时压缩标签高度，与 40px 顶栏对齐；时间：2026-11-01；作者：lq。 */
  compact?: boolean;
}) {
  const sessions = useWorkspaceStore((state) => state.sessions);
  const tabOrder = useWorkspaceStore((state) => state.tabOrder);
  const activeId = useWorkspaceStore((state) => state.activeId);
  const activateTab = useWorkspaceStore((state) => state.activateTab);
  const reorderTabs = useWorkspaceStore((state) => state.reorderTabs);
  const createDocument = useWorkspaceStore((state) => state.createDocument);
  const openFileDialog = useWorkspaceStore((state) => state.openFileDialog);

  const newItems: MenuItem[] = [
    {
      id: 'new-docx',
      label: '新建 Word 文档',
      icon: <FileText size={14} />,
      shortcut: `${MOD_KEY}N`,
      onSelect: () => void createDocument('docx'),
    },
    {
      id: 'new-xlsx',
      label: '新建 Excel 表格',
      icon: <Table2 size={14} />,
      onSelect: () => void createDocument('xlsx'),
    },
    { type: 'separator' },
    {
      id: 'open',
      label: '打开文件…',
      icon: <FolderOpen size={14} />,
      shortcut: `${MOD_KEY}O`,
      onSelect: () => void openFileDialog(),
    },
  ];

  return (
    <div className={`mo-scroll flex min-w-0 flex-1 items-stretch gap-1 overflow-x-auto ${compact ? 'px-1' : 'px-1.5 pt-1.5'}`}>
      {tabOrder.map((id) => {
        const session = sessions[id];
        if (!session) return null;
        const Icon = KIND_ICONS[session.kind];
        const active = id === activeId;
        return (
          <div
            key={id}
            role="tab"
            aria-selected={active}
            draggable
            onDragStart={(event) => {
              event.dataTransfer.setData('text/plain', id);
              event.dataTransfer.effectAllowed = 'move';
            }}
            onDragOver={(event) => {
              event.preventDefault();
              event.dataTransfer.dropEffect = 'move';
            }}
            onDrop={(event) => {
              event.preventDefault();
              const from = event.dataTransfer.getData('text/plain');
              if (from && from !== id) reorderTabs(from, id);
            }}
            onMouseDown={(event) => {
              if (event.button === 1) {
                event.preventDefault();
                onRequestClose(id);
              }
            }}
            onClick={() => {
              activateTab(id);
              onSelectTab(id);
            }}
            className={`group relative flex shrink-0 items-center gap-1.5 border px-2 text-[11px] transition-colors ${
              compact ? 'my-1 h-[26px] min-w-[104px] max-w-[190px] rounded-md' : 'h-[30px] min-w-[120px] max-w-[220px] rounded-t-md border-b-0'
            } ${
              active
                ? 'border-[var(--mo-line)] bg-[var(--mo-surface)] text-[var(--mo-text-primary)] shadow-[0_-1px_2px_rgba(16,24,40,0.03)]'
                : 'border-transparent bg-transparent text-[var(--mo-text-muted)] hover:bg-[var(--mo-surface-2)]/70 hover:text-[var(--mo-text-primary)]'
            }`}
          >
            <span
              className="flex h-4 w-4 shrink-0 items-center justify-center rounded-[3px]"
              style={{ background: `${KIND_TINTS[session.kind]}14`, color: KIND_TINTS[session.kind] }}
            >
              <Icon size={11} />
            </span>
            <span className="min-w-0 flex-1 truncate">{session.displayName}</span>
            {session.dirty ? <span className="h-1.5 w-1.5 shrink-0 rounded-full bg-[var(--mo-warning)]" /> : null}
            <button
              type="button"
              aria-label="关闭标签"
              onClick={(event) => {
                event.stopPropagation();
                onRequestClose(id);
              }}
              className={`flex h-4 w-4 shrink-0 items-center justify-center rounded hover:bg-[var(--mo-surface-3)] ${
                active ? 'opacity-100' : 'opacity-0 group-hover:opacity-100'
              }`}
            >
              <X size={11} />
            </button>
          </div>
        );
      })}
      <div className={`flex shrink-0 items-center gap-0.5 ${compact ? 'pl-0.5' : 'pb-1.5 pl-0.5 pt-1'}`}>
        <Dropdown
          minWidth={200}
          items={newItems}
          trigger={({ toggle }) => (
            <IconButton label="新建或打开" onClick={toggle}>
              <Plus size={14} />
            </IconButton>
          )}
        />
      </div>
    </div>
  );
}

export function StatusBar() {
  const sessions = useWorkspaceStore((state) => state.sessions);
  const activeId = useWorkspaceStore((state) => state.activeId);
  const settings = useWorkspaceStore((state) => state.settings);
  const updateView = useWorkspaceStore((state) => state.updateView);
  const session = activeId ? sessions[activeId] : undefined;
  const isWord = session?.kind === 'docx';
  // 功能：Excel 的缩放与 Word 统一放到右下角状态栏，工具栏只留输出类按钮；时间：2026-10-31；作者：lq。
  const isSheet = session?.kind === 'xlsx';
  const canZoom = isWord || isSheet;
  const viewZoom = session?.view.zoom ?? 1;
  const zoom = normalizeStatusZoom(viewZoom);
  const zoomPercent = Math.round(zoom * 100);
  const zoomOptions = [...new Set([40, 50, 75, 90, 100, 110, 125, 150, 175, 200, 250, zoomPercent])].sort((a, b) => a - b);
  const pageCount = Math.max(1, session?.view.pageCount ?? 1);
  const currentPage = Math.max(1, Math.min(pageCount, (session?.view.page ?? 0) + 1));
  const applyZoom = (value: number) => {
    if (canZoom && activeId) updateView(activeId, { zoom: normalizeStatusZoom(value) });
  };

  const saveLabel = !session
    ? '就绪'
    : session.state === 'SAVING'
      ? '正在保存…'
      : session.state === 'SAVE_FAILED'
        ? '保存失败'
        : session.dirty
          ? '未保存'
          : session.lastSavedAt
            ? `已保存 · ${formatDateTime(session.lastSavedAt)}`
            : '未修改';

  const stats = session ? documentStats(session) : null;

  return (
    <footer className="mo-glass mo-no-print flex h-[27px] min-w-0 shrink-0 items-center gap-2 border-t border-[var(--mo-line)] bg-[var(--mo-surface)] px-2 text-[10px] text-[var(--mo-text-muted)]">
      <div className="flex min-w-0 flex-1 items-center gap-3 overflow-hidden">
        {isWord ? <span className="shrink-0 whitespace-nowrap" aria-label="文档页数">第 {currentPage} 页 / 共 {pageCount} 页</span> : null}
        <div className="flex min-w-0 items-center gap-2">
          <span
            className={`h-1.5 w-1.5 shrink-0 rounded-full ${
              session?.state === 'SAVE_FAILED'
                ? 'bg-[var(--mo-danger)]'
                : session?.dirty
                  ? 'bg-[var(--mo-warning)]'
                  : 'bg-[var(--mo-success)]'
            }`}
          />
          <span className="whitespace-nowrap">{saveLabel}</span>
          {session?.readonly ? (
            <span className="rounded bg-[var(--mo-surface-2)] px-1 py-0.5 text-[9px]">只读保真模式</span>
          ) : session ? (
            <span className="rounded bg-[var(--mo-accent-soft)] px-1 py-0.5 text-[9px] text-[var(--mo-accent)]">可编辑</span>
          ) : null}
        </div>
        {session ? (
          <>
            {!isWord ? (
              <>
                <span className="h-3 w-px bg-[var(--mo-line)]" />
                <span className="min-w-0 truncate">{OFFICE_META[session.kind].label}</span>
              </>
            ) : null}
            {stats ? (
              <>
                <span className="h-3 w-px bg-[var(--mo-line)]" />
                <span className="whitespace-nowrap">{stats}</span>
              </>
            ) : null}
            <span className="h-3 w-px bg-[var(--mo-line)]" />
            <span className="whitespace-nowrap">{coverageLabel(session.capability)}</span>
            {session.capability.warnings.length > 0 ? (
              <span className="rounded bg-[var(--mo-warning-soft)] px-1 py-0.5 text-[9px] text-[var(--mo-warning)]">
                {session.capability.warnings.length} 条兼容提示
              </span>
            ) : null}
            {session.saveError ? (
              <span className="truncate text-[var(--mo-danger)]">{session.saveError}</span>
            ) : null}
          </>
        ) : (
          <span>Mini Office · 离线优先的办公套件</span>
        )}
      </div>
      {canZoom ? (
        <div role="group" aria-label={isWord ? 'Word 缩放' : '工作表缩放'} className="ml-auto flex shrink-0 items-center gap-0.5">
          {isWord ? (
            <button type="button" aria-label="适应宽度" title="适应文档宽度"
              className="mr-1 h-5 shrink-0 whitespace-nowrap rounded px-1.5 text-[10px] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]"
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => window.dispatchEvent(new CustomEvent('mini-office:word-fit-width', { detail: { sessionId: activeId } }))}>
              适应宽度
            </button>
          ) : null}
          <IconButton label="缩小" title={`缩小${isWord ? '文档' : '工作表'}（每次 5%）`} className="h-5 w-5" disabled={zoom <= 0.4}
            onMouseDown={(event) => event.preventDefault()} onClick={() => applyZoom(zoom - 0.05)}>
            <Minus size={12} />
          </IconButton>
          <input type="range" aria-label="缩放滑块" title={`${isWord ? '文档' : '工作表'}缩放（40% 到 250%）`} min={40} max={250} step={5}
            value={zoomPercent} aria-valuetext={`${zoomPercent}%`}
            className="h-4 w-20 min-w-0 cursor-pointer accent-[var(--mo-accent)] sm:w-24"
            onChange={(event) => applyZoom(Number(event.target.value) / 100)} />
          <IconButton label="放大" title={`放大${isWord ? '文档' : '工作表'}（每次 5%）`} className="h-5 w-5" disabled={zoom >= 2.5}
            onMouseDown={(event) => event.preventDefault()} onClick={() => applyZoom(zoom + 0.05)}>
            <Plus size={12} />
          </IconButton>
          <Select aria-label="缩放比例" title={`${isWord ? '文档' : '工作表'}缩放比例；放大 ${MOD_KEY} +；缩小 ${MOD_KEY} -；还原 ${MOD_KEY} 0`}
            className="h-5 w-[66px] border-transparent bg-transparent px-1 text-[10px] tabular-nums"
            value={String(zoomPercent)} onChange={(event) => applyZoom(Number(event.target.value) / 100)}>
            {zoomOptions.map((percent) => <option key={percent} value={percent}>{percent}%</option>)}
          </Select>
        </div>
      ) : (
        <div className="ml-auto flex shrink-0 items-center gap-3">
          <span className="hidden md:inline">自动恢复 {settings.autoSave ? '开启' : '关闭'}</span>
          <span className="hidden sm:inline">UTF-8 · OOXML</span>
          <span className="inline-flex items-center gap-1">
            <Kbd>{MOD_KEY}</Kbd>
            <Kbd>K</Kbd>
            <span className="hidden lg:inline">命令</span>
          </span>
          {/* 功能：桌面状态栏显示软件名称；时间：2026-10-24；作者：lq。 */}
          <span className="hidden lg:inline">{isTauriRuntime() ? 'Mini Office' : 'Web 模式'}</span>
        </div>
      )}
    </footer>
  );
}

function normalizeStatusZoom(value: number): number {
  return Math.max(0.4, Math.min(2.5, Math.round((Number.isFinite(value) ? value : 1) * 100) / 100));
}

function documentStats(session: DocumentSession): string {
  if (session.kind === 'docx') {
    const data = session.data as { stats?: { words?: number; characters?: number } };
    return `${data.stats?.words ?? 0} 字 · ${data.stats?.characters ?? 0} 字符`;
  }
  if (session.kind === 'xlsx') {
    const data = session.data as { sheets?: Array<{ cells: Record<string, unknown> }> };
    const cells = data.sheets?.reduce((sum, sheet) => sum + Object.keys(sheet.cells).length, 0) ?? 0;
    return `${data.sheets?.length ?? 0} 个工作表 · ${cells} 个单元格`;
  }
  const data = session.data as { slides?: unknown[]; aspect?: string };
  return `${data.slides?.length ?? 0} 页 · ${data.aspect ?? '16:9'}`;
}

export function BusyIndicator() {
  const busy = useWorkspaceStore((state) => state.busy);
  if (!busy) return null;
  return (
    <div className="mo-no-print pointer-events-none fixed left-0 right-0 top-0 z-[130] h-0.5 overflow-hidden bg-[var(--mo-accent-soft)]">
      <div className="h-full w-1/3 animate-[mo-busy_1.1s_ease-in-out_infinite] bg-[var(--mo-accent)]" />
      <style>{`@keyframes mo-busy { 0% { transform: translateX(-100%); } 100% { transform: translateX(400%); } }`}</style>
    </div>
  );
}

export function DropOverlay({ visible }: { visible: boolean }) {
  if (!visible) return null;
  return (
    <div className="mo-no-print pointer-events-none fixed inset-0 z-[125] flex items-center justify-center bg-[rgba(16,24,40,0.42)] backdrop-blur-[2px]">
      <div className="mo-pop-in flex flex-col items-center gap-3 rounded-2xl border-2 border-dashed border-white/50 bg-[rgba(255,255,255,0.12)] px-16 py-12 text-white">
        <FolderOpen size={36} />
        <div className="text-[15px] font-medium">松开以打开 Office 文件</div>
        <div className="text-[11px] text-white/70">支持 DOCX、XLSX、PPTX 与 CSV</div>
      </div>
    </div>
  );
}

export function EditorLoading({ title }: { title: string }) {
  return (
    <div className="flex h-full flex-col items-center justify-center gap-3 bg-[var(--mo-canvas)]">
      <div className="h-8 w-8 animate-[mo-spin_0.9s_linear_infinite] rounded-full border-2 border-[var(--mo-line-strong)] border-t-[var(--mo-accent)]" />
      <div className="text-[12px] text-[var(--mo-text-muted)]">正在解析 {title}…</div>
    </div>
  );
}

export function CloseConfirmDialog({
  session,
  onCancel,
  onDiscard,
  onSave,
  saving,
}: {
  session: DocumentSession;
  onCancel: () => void;
  onDiscard: () => void;
  onSave: () => void;
  saving: boolean;
}) {
  return (
    <Dialog
      open
      // 功能：保存过程中禁用关闭弹窗，避免重复处理退出选择；时间：2026-10-18；作者：lq。
      onClose={() => { if (!saving) onCancel(); }}
      title="保存更改？"
      icon={<AlertTriangle size={17} />}
      width="sm"
      closeOnBackdrop={false}
      footer={
        <>
          <Button variant="ghost" size="sm" disabled={saving} onClick={onCancel}>
            取消
          </Button>
          <Button
            variant="ghost"
            size="sm"
            className="text-[var(--mo-danger)] hover:bg-[var(--mo-danger-soft)]"
            onClick={onDiscard}
            disabled={saving}
          >
            不保存
          </Button>
          <Button variant="primary" size="sm" loading={saving} onClick={onSave}>
            {saving ? '正在保存…' : '保存并关闭'}
          </Button>
        </>
      }
    >
      <div className="space-y-3.5">
        <div className="flex items-center gap-2.5 rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/45 p-3">
          <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-[var(--mo-surface)] text-[var(--mo-text-muted)] shadow-sm">
            {session.kind === 'xlsx' ? <Table2 size={17} /> : session.kind === 'pptx' ? <Presentation size={17} /> : <FileText size={17} />}
          </span>
          <div className="min-w-0 flex-1">
            <div className="truncate text-[12px] font-medium" title={session.displayName}>
              {session.displayName}
            </div>
            <div className="mt-0.5 text-[10px] text-[var(--mo-text-muted)]">
              {session.readonly ? '只读文档' : '有未保存的修改'}
            </div>
          </div>
          {session.dirty ? <span className="h-2 w-2 shrink-0 rounded-full bg-[var(--mo-warning)]" /> : null}
        </div>
        <p className="text-[12px] leading-5 text-[var(--mo-text-primary)]">
          关闭前是否保存对此文档的更改？
        </p>
        <p className="text-[11px] leading-5 text-[var(--mo-text-muted)]">
          不保存的更改会保留在自动恢复快照中，可稍后从首页恢复。
        </p>
      </div>
    </Dialog>
  );
}

export function NotificationToasts() {
  const notifications = useWorkspaceStore((state) => state.notifications);
  const dismiss = useWorkspaceStore((state) => state.dismissNotification);
  return (
    <div className="pointer-events-none fixed bottom-9 right-4 z-[140] flex flex-col items-end gap-2">
      {notifications.map((notification) => (
        <NotificationToast key={notification.id} notification={notification} onDismiss={dismiss} />
      ))}
    </div>
  );
}

function NotificationToast({
  notification,
  onDismiss,
}: {
  notification: { id: string; level: 'info' | 'success' | 'warning' | 'error'; title: string; message?: string };
  onDismiss: (id: string) => void;
}) {
  const tones = {
    info: { border: 'var(--mo-accent)', icon: '●', text: 'text-[var(--mo-accent)]' },
    success: { border: 'var(--mo-success)', icon: '●', text: 'text-[var(--mo-success)]' },
    warning: { border: 'var(--mo-warning)', icon: '●', text: 'text-[var(--mo-warning)]' },
    error: { border: 'var(--mo-danger)', icon: '●', text: 'text-[var(--mo-danger)]' },
  };
  const tone = tones[notification.level];
  useEffect(() => {
    const timer = setTimeout(() => onDismiss(notification.id), 5200);
    return () => clearTimeout(timer);
  }, [notification.id, onDismiss]);
  return (
    <div className="mo-glass mo-pop-in pointer-events-auto flex w-[330px] items-start gap-2.5 rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface)] px-3 py-2.5 shadow-pop">
      <span className={`mt-0.5 text-[10px] ${tone.text}`}>{tone.icon}</span>
      <div className="min-w-0 flex-1">
        <div className="text-[12px] font-medium leading-5">{notification.title}</div>
        {notification.message ? (
          <div className="mt-0.5 break-words text-[11px] leading-4 text-[var(--mo-text-muted)]">{notification.message}</div>
        ) : null}
      </div>
      <button
        type="button"
        aria-label="关闭通知"
        onClick={() => onDismiss(notification.id)}
        className="mt-0.5 rounded p-0.5 text-[var(--mo-text-faint)] hover:bg-[var(--mo-surface-2)]"
      >
        <X size={13} />
      </button>
    </div>
  );
}
