import { BrandIcon } from '@mini-office/ui';
import { useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  AppWindow,
  Check,
  CheckCircle2,
  Download,
  FileCheck2,
  FileWarning,
  ExternalLink,
  Info,
  Keyboard,
  Monitor,
  Moon,
  Palette,
  Puzzle,
  RefreshCw,
  Settings as SettingsIcon,
  ShieldCheck,
  Sun,
  Trash2,
  X,
  XCircle,
} from 'lucide-react';
import { useWorkspaceStore } from '@mini-office/document-model';
import {
  Badge,
  Button,
  Dialog,
  Kbd,
  SegmentedControl,
  Select,
  Switch,
} from '@mini-office/ui';
import {
  APP_VERSION,
  DEFAULT_ACCENT_HEX,
  MOD_KEY,
  decodeDisplayName,
  type AccentId,
} from '@mini-office/shared';
import {
  ACCENT_PRESETS,
  applyAppearance,
  normalizeAccentHex,
  resolveAccentPalette,
  resolveTheme,
  type ThemeMode,
} from '@mini-office/design-system';
import { isTauriRuntime, openExternalUrl, supportsWindowVibrancy } from '@mini-office/platform-adapter';
import { collectThemePacks, usePluginRegistry } from '@mini-office/plugin-host';
import { useLatestRelease, useUpdateStore } from './update-check';
import { PluginManager } from './PluginManager';
import { ThemePreview } from './ThemePreview';
import { DefaultAppPanel, markDefaultAppPrompted } from './default-app';

// 功能：统一关于页的新版本发布地址；时间：2026-10-24；作者：lq。
const RELEASES_URL = 'https://github.com/lq0910/mini-office/releases';

export function WorkspaceDialogs() {
  const dialog = useWorkspaceStore((state) => state.dialog);
  const closeDialog = useWorkspaceStore((state) => state.closeDialog);
  const sessions = useWorkspaceStore((state) => state.sessions);

  if (!dialog) return null;
  if (dialog.type === 'settings') {
    // 功能：设置改为左侧分类 + 右侧面板的宽版窗口；时间：2026-10-29；作者：lq。
    return (
      <Dialog
        open
        onClose={closeDialog}
        title="设置"
        icon={<SettingsIcon size={16} />}
        width="2xl"
        bodyScroll={false}
        className="h-[min(640px,calc(100dvh-64px))]"
      >
        <SettingsContent onClose={closeDialog} />
      </Dialog>
    );
  }
  if (dialog.type === 'compatibility') {
    const session = dialog.sessionId ? sessions[dialog.sessionId] : undefined;
    return (
      <Dialog
        open
        onClose={closeDialog}
        title="兼容性报告"
        icon={<FileWarning size={16} />}
        width="xl"
        description="Mini Office 按 L0–L3 分级处理 Office 文件，优先保证原文件不被破坏。"
        footer={
          <div className="flex w-full items-center justify-between gap-4">
            <p className="hidden min-w-0 text-[10px] leading-4 text-[var(--mo-text-muted)] sm:block">
              L0 无法打开 · L1 仅预览 · L2 需另存 · L3 安全往返
            </p>
            <Button variant="secondary" size="sm" onClick={closeDialog} className="ml-auto shrink-0">
              关闭
            </Button>
          </div>
        }
      >
        {session ? <CompatibilityContent sessionId={session.id} onClose={closeDialog} /> : null}
      </Dialog>
    );
  }
  if (dialog.type === 'shortcuts') {
    return (
      <Dialog open onClose={closeDialog} title="键盘快捷键" icon={<Keyboard size={16} />} width="lg">
        <ShortcutsContent />
      </Dialog>
    );
  }
  if (dialog.type === 'about') {
    return (
      <Dialog open onClose={closeDialog} title="关于 Mini Office" icon={<Info size={16} />} width="md">
        <AboutContent />
      </Dialog>
    );
  }
  // 功能：发现新版本后的启动提示，同一版本只弹一次；时间：2026-10-29；作者：lq。
  if (dialog.type === 'update') {
    const close = () => {
      useUpdateStore.getState().dismiss();
      closeDialog();
    };
    return (
      <Dialog open onClose={close} title="发现新版本" icon={<Download size={16} />} width="md">
        <UpdateContent onClose={close} />
      </Dialog>
    );
  }
  if (dialog.type === 'plugins') {    return (
      <Dialog
        open
        onClose={closeDialog}
        title="插件"
        icon={<Puzzle size={16} />}
        width="lg"
        description="内核保持精简，主题、模板与行业能力通过插件扩展。插件在本地安装与运行，不上传文档。"
      >
        <PluginManager />
      </Dialog>
    );
  }
  // 功能：首次启动提示设置默认应用，关闭即记录已提示，不再重复打扰；时间：2026-10-29；作者：lq。
  if (dialog.type === 'default-app') {
    // 从菜单手动进入时不写入“已提示”标记，保留首次启动提醒。
    const isPrompt = dialog.payload?.prompt === true;
    const finish = () => {
      if (isPrompt) markDefaultAppPrompted();
      closeDialog();
    };
    return (
      <Dialog
        open
        onClose={finish}
        title={isPrompt ? '把 Mini Office 设为默认应用' : '默认应用'}
        icon={<FileCheck2 size={16} />}
        width="md"
      >
        <DefaultAppPanel variant={isPrompt ? 'prompt' : 'settings'} onClose={finish} />
      </Dialog>
    );
  }
  return null;
}

// 功能：设置窗口左侧分类导航，按用途分组避免长列表滚动；时间：2026-10-29；作者：lq。
type SettingsTabId = 'appearance' | 'editor' | 'default-app' | 'extensions' | 'privacy' | 'about';

const SETTINGS_NAV: {
  group: string;
  items: { id: SettingsTabId; label: string; icon: React.ReactNode }[];
}[] = [
  {
    group: '通用',
    items: [
      { id: 'appearance', label: '外观', icon: <Palette size={14} /> },
      { id: 'editor', label: '编辑与保存', icon: <RefreshCw size={14} /> },
    ],
  },
  {
    group: '应用',
    items: [
      { id: 'default-app', label: '默认应用', icon: <AppWindow size={14} /> },
      { id: 'extensions', label: '扩展', icon: <Puzzle size={14} /> },
    ],
  },
  {
    group: '其他',
    items: [
      { id: 'privacy', label: '安全与隐私', icon: <ShieldCheck size={14} /> },
      { id: 'about', label: '关于', icon: <Info size={14} /> },
    ],
  },
];

function SettingsContent({ onClose }: { onClose: () => void }) {
  const settings = useWorkspaceStore((state) => state.settings);
  const updateSettings = useWorkspaceStore((state) => state.updateSettings);
  const clearSnapshots = useWorkspaceStore((state) => state.clearSnapshots);
  const notify = useWorkspaceStore((state) => state.notify);
  const openDialog = useWorkspaceStore((state) => state.openDialog);

  // 功能：接入主题插件与插件管理入口；时间：2026-10-29；作者：lq。
  const installedPlugins = usePluginRegistry((state) => state.installed);
  const activeThemePack = usePluginRegistry((state) => state.activeThemePack);
  const setActiveThemePack = usePluginRegistry((state) => state.setActiveThemePack);
  const themePacks = useMemo(() => collectThemePacks(installedPlugins), [installedPlugins]);

  // 功能：设置页直接显示更新状态并可手动检查；时间：2026-10-29；作者：lq。
  const updateLatest = useUpdateStore((state) => state.latest);
  const updateStatus = useUpdateStore((state) => state.status);
  const checkUpdate = useUpdateStore((state) => state.check);

  const resolvedTheme = resolveTheme(settings.theme);
  const canVibrancy = supportsWindowVibrancy();
  const [customHex, setCustomHex] = useState(settings.customAccent);
  const [tab, setTab] = useState<SettingsTabId>('appearance');
  const customShades = resolveAccentPalette('custom', settings.customAccent, resolvedTheme);
  const currentShades = resolveAccentPalette(settings.accent, settings.customAccent, resolvedTheme);
  const activeAccentLabel =
    ACCENT_PRESETS.find((preset) => preset.id === settings.accent)?.label ?? '自定义';

  useEffect(() => {
    setCustomHex(settings.customAccent);
  }, [settings.customAccent]);

  const previewAppearance = (patch: {
    theme?: ThemeMode;
    accent?: AccentId;
    customAccent?: string;
    eyeCare?: boolean;
  }) => {
    applyAppearance({
      theme: patch.theme ?? settings.theme,
      accent: patch.accent ?? settings.accent,
      customAccent: patch.customAccent ?? settings.customAccent,
      eyeCare: patch.eyeCare ?? settings.eyeCare,
    });
  };

  // 功能：护眼模式即时预览，避免用户开关后看不出变化；时间：2026-10-30；作者：lq。
  const setEyeCare = (eyeCare: boolean) => {
    updateSettings({ eyeCare });
    previewAppearance({ eyeCare });
  };

  const setTheme = (theme: ThemeMode) => {
    updateSettings({ theme });
    previewAppearance({ theme });
  };

  const chooseAccent = (accent: AccentId) => {
    updateSettings({ accent });
    previewAppearance({ accent });
  };

  const commitCustomHex = (value: string) => {
    const normalized = normalizeAccentHex(value);
    if (!normalized) return false;
    updateSettings({ accent: 'custom', customAccent: normalized });
    previewAppearance({ accent: 'custom', customAccent: normalized });
    return true;
  };

  const resetDefaults = () => {
    updateSettings({
      theme: 'system',
      accent: 'mist',
      customAccent: DEFAULT_ACCENT_HEX,
      vibrancy: false,
      autoSave: true,
      autoSaveDelayMs: 1400,
      showProperties: false,
      showToolbar: true,
      restoreSession: true,
      eyeCare: false,
    });
    applyAppearance({
      theme: 'system',
      accent: 'mist',
      customAccent: DEFAULT_ACCENT_HEX,
      eyeCare: false,
    });
    setCustomHex(DEFAULT_ACCENT_HEX);
    notify('success', '已恢复默认设置');
  };

  return (
    <>
      <nav
        aria-label="设置分类"
        className="flex w-[188px] shrink-0 flex-col gap-3 overflow-y-auto border-r border-[var(--mo-line)] bg-[var(--mo-surface-2)]/40 px-2.5 py-3"
      >
        {SETTINGS_NAV.map((group) => (
          <div key={group.group} className="flex flex-col gap-0.5">
            <span className="px-2 pb-1 text-[10px] font-semibold uppercase tracking-[0.08em] text-[var(--mo-text-faint)]">
              {group.group}
            </span>
            {group.items.map((item) => {
              const active = tab === item.id;
              return (
                <button
                  key={item.id}
                  type="button"
                  aria-current={active ? 'page' : undefined}
                  onClick={() => setTab(item.id)}
                  className={`flex items-center gap-2 rounded-lg px-2 py-[7px] text-left text-[12px] transition-colors duration-fast ${
                    active
                      ? 'bg-[var(--mo-accent-soft)] font-medium text-[var(--mo-accent)]'
                      : 'text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]'
                  }`}
                >
                  <span className={active ? 'text-[var(--mo-accent)]' : 'text-[var(--mo-text-muted)]'}>{item.icon}</span>
                  <span className="min-w-0 flex-1 truncate">{item.label}</span>
                </button>
              );
            })}
          </div>
        ))}
      </nav>

      <div className="flex min-w-0 flex-1 flex-col">
        <div className="mo-scroll min-h-0 flex-1 overflow-y-auto px-6 py-5">
          {tab === 'appearance' ? (
            <SettingsPane
              title="外观"
              description="主题、配色与界面语言只影响应用外壳，不会改变文档排版与打印结果。"
            >
              <section className="space-y-3">
                <SectionTitle icon={<Sun size={13} />} title="主题与配色" />
                <div className="flex flex-col gap-1.5">
                  <span className="text-[12px] font-medium text-[var(--mo-text-primary)]">主题</span>
                  <SegmentedControl
                    value={settings.theme}
                    onChange={setTheme}
                    className="w-full max-w-[360px]"
                    options={[
                      { value: 'light', label: <span className="inline-flex items-center gap-1"><Sun size={12} /> 浅色</span> },
                      { value: 'dark', label: <span className="inline-flex items-center gap-1"><Moon size={12} /> 深色</span> },
                      { value: 'system', label: <span className="inline-flex items-center gap-1"><Monitor size={12} /> 跟随系统</span> },
                    ]}
                  />
                </div>

                <div className="flex flex-col gap-1.5">
                  <span className="text-[12px] font-medium text-[var(--mo-text-primary)]">主题色</span>
                  <div role="group" aria-label="主题色" className="grid grid-cols-2 gap-1.5 sm:grid-cols-3">
                    {ACCENT_PRESETS.map((preset) => {
                      const active = settings.accent === preset.id;
                      const shades = preset[resolvedTheme];
                      return (
                        <button
                          key={preset.id}
                          type="button"
                          aria-label={preset.label}
                          aria-pressed={active}
                          title={preset.description}
                          onClick={() => chooseAccent(preset.id)}
                          className={`flex items-center gap-2 rounded-lg border px-2 py-1.5 text-left transition-colors duration-fast ${
                            active
                              ? 'border-[var(--mo-accent)] bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]'
                              : 'border-[var(--mo-line)] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]'
                          }`}
                        >
                          <span
                            className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full ring-1 ring-black/10"
                            style={{ background: shades.accent }}
                          >
                            {active ? <Check size={11} style={{ color: shades.ink }} /> : null}
                          </span>
                          <span className="min-w-0 flex-1 truncate text-[11.5px] font-medium">
                            {preset.label}
                          </span>
                        </button>
                      );
                    })}
                    <button
                      type="button"
                      aria-label="自定义"
                      aria-pressed={settings.accent === 'custom'}
                      title="使用自定义颜色"
                      onClick={() => chooseAccent('custom')}
                      className={`flex items-center gap-2 rounded-lg border px-2 py-1.5 text-left transition-colors duration-fast ${
                        settings.accent === 'custom'
                          ? 'border-[var(--mo-accent)] bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]'
                          : 'border-[var(--mo-line)] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]'
                      }`}
                    >
                      <span
                        className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full ring-1 ring-black/10"
                        style={{
                          background:
                            settings.accent === 'custom'
                              ? customShades.accent
                              : 'conic-gradient(from 210deg, #2E7D72, #2F63E0, #6B5BD6, #B25F4A, #2E7D72)',
                        }}
                      >
                        {settings.accent === 'custom' ? (
                          <Check size={11} style={{ color: customShades.ink }} />
                        ) : null}
                      </span>
                      <span className="min-w-0 flex-1 truncate text-[11.5px] font-medium">自定义</span>
                    </button>
                  </div>
                  <span className="text-[11px] leading-4 text-[var(--mo-text-muted)]">
                    主题色用于按钮、选区、链接与高亮；浅色/深色会分别自动适配。
                  </span>
                </div>

                {themePacks.length > 0 ? (
                  <div className="flex flex-col gap-1.5">
                    <span className="text-[12px] font-medium text-[var(--mo-text-primary)]">界面风格</span>
                    <div role="group" aria-label="界面风格" className="grid grid-cols-2 gap-1.5 sm:grid-cols-3">
                      <button
                        type="button"
                        aria-pressed={activeThemePack === 'builtin'}
                        onClick={() => setActiveThemePack('builtin')}
                        className={`flex items-center gap-2 rounded-lg border px-2 py-1.5 text-left transition-colors duration-fast ${
                          activeThemePack === 'builtin'
                            ? 'border-[var(--mo-accent)] bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]'
                            : 'border-[var(--mo-line)] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]'
                        }`}
                      >
                        <span className="relative flex h-6 w-8 shrink-0 items-center justify-center overflow-hidden rounded bg-[var(--mo-surface-3)] ring-1 ring-black/10">
                          <span className="flex h-full w-full flex-col">
                            <span className="h-[34%] w-full bg-[var(--mo-accent)] opacity-80" />
                            <span className="flex flex-1 flex-col justify-center gap-[2px] px-[3px]">
                              <span className="h-[1.5px] w-[72%] bg-[var(--mo-text-primary)] opacity-70" />
                              <span className="h-[1.5px] w-[48%] bg-[var(--mo-text-primary)] opacity-45" />
                            </span>
                          </span>
                          {activeThemePack === 'builtin' ? (
                            <span className="absolute inset-0 flex items-center justify-center bg-black/25 text-white">
                              <Check size={11} />
                            </span>
                          ) : null}
                        </span>
                        <span className="min-w-0 flex-1 truncate text-[11.5px] font-medium">Mini 默认</span>
                      </button>
                      {themePacks.map((pack) => {
                        const active = activeThemePack === pack.key;
                        return (
                          <button
                            key={pack.key}
                            type="button"
                            aria-pressed={active}
                            title={pack.description ?? pack.pluginName}
                            onClick={() => setActiveThemePack(pack.key)}
                            className={`flex items-center gap-2 rounded-lg border px-2 py-1.5 text-left transition-colors duration-fast ${
                              active
                                ? 'border-[var(--mo-accent)] bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]'
                                : 'border-[var(--mo-line)] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)]'
                            }`}
                          >
                            <span className="relative flex h-6 w-8 shrink-0 items-center justify-center overflow-hidden rounded ring-1 ring-black/10">
                              <ThemePreview pack={pack} className="h-full w-full rounded" />
                              {active ? (
                                <span className="absolute inset-0 flex items-center justify-center bg-black/25 text-white">
                                  <Check size={11} />
                                </span>
                              ) : null}
                            </span>
                            <span className="min-w-0 flex-1 truncate text-[11.5px] font-medium">{pack.label}</span>
                          </button>
                        );
                      })}
                    </div>
                    <span className="text-[11px] leading-4 text-[var(--mo-text-muted)]">
                      界面风格由主题插件提供，只改变外壳，不影响文档排版与打印结果。
                    </span>
                  </div>
                ) : null}

                {settings.accent === 'custom' ? (
                  <div className="flex flex-col gap-1.5">
                    <span className="text-[12px] font-medium text-[var(--mo-text-primary)]">自定义颜色</span>
                    <div className="flex items-center gap-2">
                      <input
                        type="color"
                        aria-label="选择自定义主题色"
                        value={normalizeAccentHex(customHex) ?? DEFAULT_ACCENT_HEX}
                        onChange={(event) => {
                          setCustomHex(event.target.value);
                          commitCustomHex(event.target.value);
                        }}
                        className="h-8 w-10 shrink-0 cursor-pointer rounded-md border border-[var(--mo-line-strong)] bg-transparent p-0.5"
                      />
                      <input
                        value={customHex}
                        aria-label="主题色十六进制值"
                        spellCheck={false}
                        onChange={(event) => {
                          const value = event.target.value;
                          setCustomHex(value);
                          if (normalizeAccentHex(value)) commitCustomHex(value);
                        }}
                        onBlur={() => {
                          if (!normalizeAccentHex(customHex)) setCustomHex(settings.customAccent);
                        }}
                        className={`mo-field h-8 w-[132px] font-mono text-[11px] ${
                          normalizeAccentHex(customHex) ? '' : 'border-[var(--mo-danger)]'
                        }`}
                      />
                      <span className="text-[10px] text-[var(--mo-text-muted)]">
                        当前：{activeAccentLabel}
                      </span>
                    </div>
                    <span className="text-[11px] leading-4 text-[var(--mo-text-muted)]">
                      支持 #RGB 或 #RRGGBB，深色模式会自动提亮。
                    </span>
                  </div>
                ) : null}

                <div className="flex flex-wrap items-center gap-2 rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/45 px-3 py-2.5">
                  <span className="rounded-md bg-[var(--mo-accent)] px-2.5 py-1 text-[11px] font-medium text-[var(--mo-accent-ink)]">
                    主操作
                  </span>
                  <span className="rounded-md bg-[var(--mo-accent-soft)] px-2 py-1 text-[11px] font-medium text-[var(--mo-accent)]">
                    标签
                  </span>
                  <span className="text-[11px] font-medium text-[var(--mo-accent)] underline decoration-1 underline-offset-2">
                    链接
                  </span>
                  <span className="ml-auto text-[10px] text-[var(--mo-text-muted)]">
                    {activeAccentLabel} · {currentShades.accent}
                  </span>
                </div>

                {/* 功能：护眼模式开关，纸张与界面同步转米黄，导出仍为白纸；时间：2026-10-30；作者：lq。 */}
                <SettingRow
                  title="护眼模式"
                  description="界面与纸张改用低蓝光米黄色，长时间阅读更舒适；打印与导出 PDF 仍是标准白纸。"
                  control={
                    <Switch checked={settings.eyeCare} onChange={setEyeCare} label="护眼模式" />
                  }
                />
                {/* 功能：标尺默认隐藏（与 Word / WPS 一致），需要精确缩进与制表位时再开启；时间：2026-10-30；作者：lq。 */}
                <SettingRow
                  title="Word 标尺"
                  description="在正文上方显示标尺，可拖动首行与左右缩进、增删制表位；排版合同、公文填空行时再开启。"
                  control={
                    <Switch
                      checked={settings.wordRuler === true}
                      onChange={(checked) => updateSettings({ wordRuler: checked })}
                      label="Word 标尺"
                    />
                  }
                />
                {/* 功能：持久保存裁剪标记开关，显示状态不影响文档内容和输出；时间：2026-11-05；作者：lq。 */}
                <SettingRow
                  title="Word 裁剪标记"
                  description="在正文区域四角显示灰色直角标记，随页边距和页面缩放定位；打印与 PDF 不显示。"
                  control={<Switch checked={settings.wordCropMarks !== false} onChange={(checked) => updateSettings({ wordCropMarks: checked })} label="Word 裁剪标记" />}
                />
                {canVibrancy ? (
                  <SettingRow
                    title="毛玻璃效果"
                    description="使用 macOS 原生磨砂材质，让标题栏与面板透出桌面；深色模式自动适配。"
                    control={
                      <Switch
                        checked={settings.vibrancy}
                        onChange={(checked) => updateSettings({ vibrancy: checked })}
                        label="毛玻璃效果"
                      />
                    }
                  />
                ) : null}
                {canVibrancy && settings.vibrancy ? (
                  <p className="rounded-md bg-[var(--mo-surface-2)] px-3 py-2 text-[10px] leading-4 text-[var(--mo-text-muted)]">
                    原生磨砂材质在窗口创建时生效。如果刚更新过应用或效果不明显，请完全退出 Mini Office 后重新打开。
                  </p>
                ) : null}

                <SettingRow
                  title="界面语言"
                  description="当前仅提供简体中文，后续版本会按需增加。"
                  control={
                    <Select
                      aria-label="界面语言"
                      className="w-[168px] shrink-0"
                      value={settings.language}
                      onChange={(event) => updateSettings({ language: event.target.value as 'zh-CN' })}
                    >
                      <option value="zh-CN">简体中文</option>
                    </Select>
                  }
                />
              </section>
            </SettingsPane>
          ) : null}

          {tab === 'editor' ? (
            <SettingsPane
              title="编辑与保存"
              description="自动恢复、启动行为与界面密度；缩放比例请使用工具栏中的缩放控件。"
            >
              <section className="space-y-2">
                <SectionTitle icon={<RefreshCw size={13} />} title="保存与恢复" />
                <SettingRow
                  title="自动恢复快照"
                  description="编辑时在本地保存恢复快照，异常退出后可继续编辑。"
                  control={<Switch checked={settings.autoSave} onChange={(checked) => updateSettings({ autoSave: checked })} label="自动恢复快照" />}
                />
                <SettingRow
                  title="自动恢复间隔"
                  description="两次快照之间的等待时间，越短越安全，写入也越频繁。"
                  control={
                    <Select
                      aria-label="自动恢复间隔"
                      className="w-[168px] shrink-0"
                      value={String(settings.autoSaveDelayMs)}
                      onChange={(event) => updateSettings({ autoSaveDelayMs: Number(event.target.value) })}
                      disabled={!settings.autoSave}
                    >
                      <option value="800">0.8 秒</option>
                      <option value="1400">1.4 秒（推荐）</option>
                      <option value="2500">2.5 秒</option>
                      <option value="4000">4 秒</option>
                    </Select>
                  }
                />
                <SettingRow
                  title="启动时显示最近文件"
                  description="首页展示本机快照与可恢复文件。"
                  control={<Switch checked={settings.restoreSession} onChange={(checked) => updateSettings({ restoreSession: checked })} label="显示最近文件" />}
                />
                <SettingRow
                  title="显示工具栏"
                  description="隐藏后可用命令面板调用高频命令。"
                  control={<Switch checked={settings.showToolbar} onChange={(checked) => updateSettings({ showToolbar: checked })} label="显示工具栏" />}
                />
              </section>
            </SettingsPane>
          ) : null}

          {tab === 'default-app' ? (
            <SettingsPane
              title="默认应用"
              description="把常见 Office 扩展名指向 Mini Office，双击文件即可直接打开。"
            >
              <DefaultAppPanel variant="settings" />
            </SettingsPane>
          ) : null}

          {tab === 'extensions' ? (
            <SettingsPane
              title="扩展"
              description="内核保持精简，主题、模板与行业能力通过插件扩展；插件在本地安装与运行。"
            >
              <SettingRow
                title="插件管理"
                description={
                  installedPlugins.length > 0
                    ? `已安装 ${installedPlugins.length} 个插件，可管理主题、模板与扩展能力。`
                    : '尚未安装插件，可按需扩展主题、模板与行业能力。'
                }
                control={
                  <Button size="sm" variant="secondary" onClick={() => openDialog({ type: 'plugins' })}>
                    管理插件
                  </Button>
                }
              />
              <SettingRow
                title="界面风格"
                description={
                  themePacks.length > 0
                    ? `当前有 ${themePacks.length} 个主题插件提供可选风格，可在「外观」中切换。`
                    : '安装主题插件后，可在外观中切换界面风格。'
                }
                control={
                  <Button size="sm" variant="ghost" onClick={() => setTab('appearance')}>
                    前往外观
                  </Button>
                }
              />
            </SettingsPane>
          ) : null}

          {tab === 'privacy' ? (
            <SettingsPane
              title="安全与隐私"
              description="Mini Office 离线运行，不将文档上传到任何服务器。"
            >
              <p className="rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/40 px-3.5 py-3 text-[11px] leading-5 text-[var(--mo-text-muted)]">
                最近文件与自动恢复数据保存在本机浏览器 IndexedDB 或应用数据目录中，不会同步到云端。
              </p>
              <div className="flex items-center justify-between gap-4 rounded-xl border border-[var(--mo-line)] px-3.5 py-2.5">
                <div className="min-w-0">
                  <div className="text-[12px] font-medium">清除所有恢复快照</div>
                  <div className="mt-0.5 text-[10px] leading-4 text-[var(--mo-text-muted)]">
                    不会删除磁盘上的原始 Office 文件。
                  </div>
                </div>
                <Button
                  size="sm"
                  variant="ghost"
                  icon={<Trash2 size={13} />}
                  className="shrink-0 text-[var(--mo-danger)]"
                  onClick={() => {
                    if (window.confirm('确定清除所有自动恢复快照与最近文件记录吗？')) {
                      void clearSnapshots().then(() => notify('success', '已清除恢复快照'));
                    }
                  }}
                >
                  清除
                </Button>
              </div>
            </SettingsPane>
          ) : null}

          {tab === 'about' ? (
            <SettingsPane title="关于" description="版本信息与本地优先的运行方式。">
              <div className="flex items-center gap-3 rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/40 px-3.5 py-3">
                <BrandIcon size={38} />
                <div className="min-w-0">
                  <div className="text-[13px] font-semibold text-[var(--mo-text-primary)]">Mini Office</div>
                  <div className="mt-0.5 text-[10.5px] text-[var(--mo-text-muted)]">
                    版本 {APP_VERSION} · 本地优先 · 离线可用
                  </div>
                </div>
              </div>
              <SettingRow
                title="版本与更新"
                description={
                  updateLatest
                    ? `发现新版本 ${updateLatest.tagName}，当前 ${APP_VERSION}`
                    : '查看发布说明，或手动检查新版本。'
                }
                control={
                  <div className="flex items-center gap-1.5">
                    <Button
                      size="sm"
                      variant={updateLatest ? 'primary' : 'secondary'}
                      icon={updateStatus === 'checking' ? <RefreshCw size={12} className="animate-spin" /> : undefined}
                      disabled={updateStatus === 'checking'}
                      onClick={() => void checkUpdate({ force: true })}
                    >
                      {updateLatest ? '有更新' : '检查更新'}
                    </Button>
                    <Button size="sm" variant="ghost" onClick={() => openDialog({ type: 'about' })}>
                      查看详情
                    </Button>
                  </div>
                }
              />
              <SettingRow
                title="发布页面"
                description="在系统浏览器中打开 GitHub Releases。"
                control={
                  <Button
                    size="sm"
                    variant="ghost"
                    icon={<ExternalLink size={12} />}
                    onClick={() => void openExternalUrl(RELEASES_URL)}
                  >
                    打开
                  </Button>
                }
              />
            </SettingsPane>
          ) : null}
        </div>

        <div className="flex items-center justify-between gap-3 border-t border-[var(--mo-line)] px-6 py-3">
          <Button size="sm" variant="ghost" onClick={resetDefaults}>
            恢复默认设置
          </Button>
          <Button variant="primary" size="sm" onClick={onClose}>
            完成
          </Button>
        </div>
      </div>
    </>
  );
}

function SettingsPane({
  title,
  description,
  children,
}: {
  title: string;
  description?: string;
  children: React.ReactNode;
}) {
  return (
    <div className="flex max-w-[620px] flex-col gap-4">
      <header>
        <h3 className="text-[15px] font-semibold tracking-[-0.01em] text-[var(--mo-text-primary)]">{title}</h3>
        {description ? (
          <p className="mt-1 text-[11.5px] leading-5 text-[var(--mo-text-muted)]">{description}</p>
        ) : null}
      </header>
      {children}
    </div>
  );
}

function SectionTitle({ icon, title }: { icon: React.ReactNode; title: string }) {
  return (
    <div className="flex items-center gap-1.5 text-[11px] font-semibold text-[var(--mo-text-muted)]">
      {icon}
      {title}
    </div>
  );
}

function SettingRow({
  title,
  description,
  control,
}: {
  title: string;
  description: string;
  control: React.ReactNode;
}) {
  return (
    <div className="flex items-center justify-between gap-4 rounded-md border border-[var(--mo-line)] px-3 py-2">
      <div className="min-w-0">
        <div className="text-[12px] font-medium">{title}</div>
        <div className="mt-0.5 text-[10px] leading-4 text-[var(--mo-text-muted)]">{description}</div>
      </div>
      {control}
    </div>
  );
}

function CompatibilityContent({ sessionId, onClose }: { sessionId: string; onClose: () => void }) {
  const session = useWorkspaceStore((state) => state.sessions[sessionId]);
  const saveSession = useWorkspaceStore((state) => state.saveSession);
  const [saving, setSaving] = useState(false);
  // 功能：记录字体重扫结果，让用户确认新装字体已被识别；时间：2026-10-29；作者：lq。
  const [fontScan, setFontScan] = useState<{ busy: boolean; found: number; total: number } | null>(null);

  const grouped = useMemo(() => {
    if (!session) return { supported: [], unsupported: [] };
    return {
      supported: session.capability.features.filter((feature) => feature.supported),
      unsupported: session.capability.features.filter((feature) => !feature.supported),
    };
  }, [session]);

  if (!session) return <p className="text-[12px] text-[var(--mo-text-muted)]">文档已关闭。</p>;

  const level = session.capability.level;
  const displayName = decodeDisplayName(session.displayName);
  const tone =
    level >= 3
      ? {
          bg: 'var(--mo-success-soft)',
          color: 'var(--mo-success)',
          label: 'L3 安全往返保存',
          note: '原文件可安全保存；写入前会校验 OOXML 容器，并保留未知部件。',
        }
      : level === 2
        ? {
            bg: 'var(--mo-warning-soft)',
            color: 'var(--mo-warning)',
            label: 'L2 可编辑，建议另存副本',
            note: '文档包含未完全支持的特性，建议另存副本以保护原始文件。',
          }
        : level === 1
          ? {
              bg: 'var(--mo-surface-2)',
              color: 'var(--mo-text-muted)',
              label: 'L1 仅预览 / 只读',
              note: '当前为只读预览，原始文件不会被修改。',
            }
          : {
              bg: 'var(--mo-danger-soft)',
              color: 'var(--mo-danger)',
              label: 'L0 无法打开',
              note: '文件无法安全解析，Mini Office 不会写入原始文件。',
            };
  const warnings = session.capability.warnings.filter((warning) => warning.code !== 'fonts.missing');
  const missingFonts = session.capability.missingFonts;
  const warningTone = {
    blocking: {
      box: 'border-[color-mix(in_srgb,var(--mo-danger)_22%,var(--mo-line))] bg-[var(--mo-danger-soft)]/25',
      icon: 'bg-[var(--mo-danger-soft)] text-[var(--mo-danger)]',
      label: '阻断',
      labelClass: 'text-[var(--mo-danger)]',
    },
    warning: {
      box: 'border-[color-mix(in_srgb,var(--mo-warning)_24%,var(--mo-line))] bg-[var(--mo-warning-soft)]/22',
      icon: 'bg-[var(--mo-warning-soft)] text-[var(--mo-warning)]',
      label: '注意',
      labelClass: 'text-[var(--mo-warning)]',
    },
    info: {
      box: 'border-[var(--mo-line)] bg-[var(--mo-surface-2)]/45',
      icon: 'bg-[var(--mo-accent-soft)] text-[var(--mo-accent)]',
      label: '提示',
      labelClass: 'text-[var(--mo-accent)]',
    },
  } as const;

  return (
    <div className="space-y-4">
      <div className="overflow-hidden rounded-xl border border-[var(--mo-line)] bg-gradient-to-br from-[var(--mo-surface-2)]/70 via-[var(--mo-surface)] to-[var(--mo-surface)]">
        <div className="flex items-start gap-3 p-3.5">
          <div
            className="flex h-11 w-11 shrink-0 items-center justify-center rounded-xl"
            style={{ background: tone.bg, color: tone.color }}
          >
            {level >= 3 ? <CheckCircle2 size={22} /> : level === 1 ? <Info size={22} /> : <AlertTriangle size={22} />}
          </div>
          <div className="min-w-0 flex-1">
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-[14px] font-semibold" style={{ color: tone.color }}>
                {tone.label}
              </span>
              <span className="rounded-full border border-[var(--mo-line)] bg-[var(--mo-surface)]/80 px-1.5 py-0.5 font-mono text-[9px] font-semibold text-[var(--mo-text-muted)]">
                L{level}
              </span>
            </div>
            <div className="mt-1 flex min-w-0 items-center gap-1.5 text-[11px] text-[var(--mo-text-muted)]">
              <span className="truncate" title={displayName}>
                {displayName}
              </span>
              <span className="shrink-0 text-[var(--mo-text-faint)]">·</span>
              <span className="shrink-0">{session.readonly ? '只读模式' : '可编辑'}</span>
            </div>
            <p className="mt-1.5 text-[10.5px] leading-4 text-[var(--mo-text-muted)]">{tone.note}</p>
          </div>
          {level === 2 && !session.readonly ? (
            <Button
              size="sm"
              variant="soft"
              loading={saving}
              className="shrink-0"
              onClick={async () => {
                setSaving(true);
                const ok = await saveSession(session.id, { asNew: true });
                setSaving(false);
                if (ok) onClose();
              }}
            >
              另存副本
            </Button>
          ) : null}
        </div>
      </div>

      {warnings.length > 0 ? (
        <section className="space-y-2">
          <SectionHeading title="兼容性提示" count={warnings.length} />
          {warnings.map((warning, index) => {
            const visual = warningTone[warning.severity];
            return (
              <div
                key={`${warning.code}-${index}`}
                className={`flex gap-2.5 rounded-xl border p-2.5 ${visual.box}`}
              >
                <span className={`flex h-7 w-7 shrink-0 items-center justify-center rounded-lg ${visual.icon}`}>
                  {warning.severity === 'blocking' ? (
                    <XCircle size={14} />
                  ) : warning.severity === 'warning' ? (
                    <AlertTriangle size={14} />
                  ) : (
                    <Info size={14} />
                  )}
                </span>
                <div className="min-w-0 flex-1">
                  <div className="flex flex-wrap items-center gap-1.5">
                    <span className="rounded bg-[var(--mo-surface)]/70 px-1.5 py-0.5 font-mono text-[9px] text-[var(--mo-text-muted)]">
                      {warning.code}
                    </span>
                    <span className={`text-[9px] font-semibold ${visual.labelClass}`}>{visual.label}</span>
                  </div>
                  <p className="mt-1 break-words text-[11.5px] leading-5 text-[var(--mo-text-primary)]">
                    {warning.message}
                  </p>
                  {warning.detail ? (
                    <p className="mt-1 break-words text-[10px] leading-4 text-[var(--mo-text-faint)]">
                      {warning.detail}
                    </p>
                  ) : null}
                </div>
              </div>
            );
          })}
        </section>
      ) : null}

      {missingFonts.length > 0 ? (
        <section className="space-y-2">
          <SectionHeading title="缺失字体" count={missingFonts.length} />
          <div className="rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/35 p-2.5">
            <div className="flex flex-wrap gap-1.5">
              {missingFonts.map((font) => (
                <span
                  key={font}
                  title={font}
                  className="max-w-full truncate rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface)] px-2 py-1 text-[10px] text-[var(--mo-text-primary)]"
                >
                  {font}
                </span>
              ))}
            </div>
            <p className="mt-2 text-[10px] leading-4 text-[var(--mo-text-muted)]">
              缺失字体将使用替代字体显示，版面可能与原始文件不同。建议在目标系统安装相应字体。
            </p>
            <div className="mt-2 flex flex-wrap items-center gap-2">
              <button
                type="button"
                disabled={fontScan?.busy}
                onClick={() => {
                  const missing = missingFonts;
                  setFontScan({ busy: true, found: 0, total: missing.length });
                  void useWorkspaceStore
                    .getState()
                    .refreshFonts()
                    .then((fonts) => {
                      const normalize = (value: string) => value.toLowerCase().replace(/\s+/g, '');
                      const available = new Set(
                        fonts.flatMap((font) => [font.family, ...(font.aliases ?? [])]).map(normalize),
                      );
                      const found = missing.filter((font) => available.has(normalize(font))).length;
                      setFontScan({ busy: false, found, total: missing.length });
                    })
                    .catch(() => setFontScan({ busy: false, found: 0, total: missing.length }));
                }}
                className="rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface)] px-2 py-1 text-[10px] text-[var(--mo-text-primary)] hover:bg-[var(--mo-surface-2)] disabled:opacity-50"
              >
                {fontScan?.busy ? '正在重新扫描…' : '重新扫描系统字体'}
              </button>
              {fontScan && !fontScan.busy ? (
                <span className="text-[10px] text-[var(--mo-text-muted)]">
                  {fontScan.found > 0
                    ? `已识别 ${fontScan.found}/${fontScan.total} 个缺失字体，重新打开文档后报告将更新。`
                    : `仍未找到这 ${fontScan.total} 个字体，请确认字体已安装到系统字体目录后重试。`}
                </span>
              ) : null}
            </div>
          </div>
        </section>
      ) : null}

      <section className="space-y-2">
        <SectionHeading title="能力概览" />
        <div className="rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/30 p-2.5">
          {grouped.supported.length === 0 && grouped.unsupported.length === 0 ? (
            <p className="text-[10px] text-[var(--mo-text-faint)]">暂无能力记录</p>
          ) : (
            <div className="flex flex-wrap gap-1.5">
              {grouped.supported.map((feature) => (
                <span
                  key={feature.code}
                  className="inline-flex max-w-full items-center gap-1 rounded-md border border-[color-mix(in_srgb,var(--mo-success)_18%,var(--mo-line))] bg-[var(--mo-success-soft)]/25 px-1.5 py-1 text-[10px] text-[var(--mo-text-primary)]"
                >
                  <Check size={10} className="shrink-0 text-[var(--mo-success)]" />
                  <span className="min-w-0 truncate" title={feature.label}>{feature.label}</span>
                </span>
              ))}
              {grouped.unsupported.map((feature) => (
                <span
                  key={feature.code}
                  className="inline-flex max-w-full items-center gap-1 rounded-md border border-[color-mix(in_srgb,var(--mo-warning)_20%,var(--mo-line))] bg-[var(--mo-warning-soft)]/30 px-1.5 py-1 text-[10px] text-[var(--mo-text-primary)]"
                >
                  <X size={10} className="shrink-0 text-[var(--mo-warning)]" />
                  <span className="min-w-0 truncate" title={feature.label}>{feature.label}</span>
                </span>
              ))}
            </div>
          )}
        </div>
      </section>

    </div>
  );
}

function SectionHeading({ title, count }: { title: string; count?: number }) {
  return (
    <div className="flex items-center gap-2">
      <h3 className="text-[10.5px] font-semibold uppercase tracking-[0.04em] text-[var(--mo-text-muted)]">
        {title}
      </h3>
      {typeof count === 'number' ? (
        <span className="rounded-full bg-[var(--mo-surface-2)] px-1.5 py-0.5 text-[9px] font-medium tabular-nums text-[var(--mo-text-muted)]">
          {count}
        </span>
      ) : null}
      <span className="h-px flex-1 bg-[var(--mo-line)]" />
    </div>
  );
}

const SHORTCUTS: Array<{ group: string; items: Array<{ label: string; keys: string[] }> }> = [
  {
    group: '文件',
    items: [
      { label: '打开文件', keys: [MOD_KEY, 'O'] },
      { label: '新建 Word 文档', keys: [MOD_KEY, 'N'] },
      { label: '保存', keys: [MOD_KEY, 'S'] },
      { label: '另存为', keys: ['⇧', MOD_KEY, 'S'] },
      { label: '关闭标签', keys: [MOD_KEY, 'W'] },
      { label: '打印', keys: [MOD_KEY, 'P'] },
    ],
  },
  {
    group: '编辑',
    items: [
      { label: '撤销', keys: [MOD_KEY, 'Z'] },
      { label: '重做', keys: ['⇧', MOD_KEY, 'Z'] },
      { label: '查找替换', keys: [MOD_KEY, 'F'] },
      { label: '复制 / 粘贴', keys: [MOD_KEY, 'C / V'] },
      { label: '全选', keys: [MOD_KEY, 'A'] },
    ],
  },
  {
    group: '视图',
    items: [
      { label: '命令面板', keys: [MOD_KEY, 'K'] },
      { label: '设置', keys: [MOD_KEY, ','] },
      // 功能：将 Word 文档缩放快捷键加入内置帮助；时间：2026-10-24；作者：lq。
      { label: '放大 Word 文档', keys: [MOD_KEY, '+'] },
      { label: '缩小 Word 文档', keys: [MOD_KEY, '-'] },
      { label: 'Word 文档恢复 100%', keys: [MOD_KEY, '0'] },
      { label: 'Word 文档滚轮缩放', keys: [MOD_KEY, '滚轮'] },
      // 功能：将 Excel 工作表缩放快捷键加入内置帮助；时间：2026-10-24；作者：lq。
      { label: 'Excel 工作表放大', keys: [MOD_KEY, '+'] },
      { label: 'Excel 工作表缩小', keys: [MOD_KEY, '-'] },
      { label: 'Excel 工作表恢复 100%', keys: [MOD_KEY, '0'] },
      { label: 'Excel 工作表滚轮缩放', keys: [MOD_KEY, '滚轮'] },
      { label: '全屏演示（PPT）', keys: ['F5'] },
      { label: '退出演示 / 关闭弹窗', keys: ['Esc'] },
    ],
  },
];

function ShortcutsContent() {
  return (
    <div className="space-y-4">
      {SHORTCUTS.map((section) => (
        <section key={section.group}>
          <h3 className="mb-2 text-[10px] font-semibold uppercase tracking-wide text-[var(--mo-text-faint)]">
            {section.group}
          </h3>
          <div className="space-y-1">
            {section.items.map((item) => (
              <div
                key={item.label}
                className="flex items-center justify-between rounded-md px-2 py-1.5 text-[12px] hover:bg-[var(--mo-surface-2)]"
              >
                <span>{item.label}</span>
                <span className="flex items-center gap-1">
                  {item.keys.map((key, index) => (
                    <Kbd key={`${key}-${index}`}>{key}</Kbd>
                  ))}
                </span>
              </div>
            ))}
          </div>
        </section>
      ))}
      <p className="text-[10px] leading-4 text-[var(--mo-text-muted)]">
        在表格中，方向键移动选区，Shift + 方向键扩展选区，输入任意字符开始编辑，F2 编辑当前单元格。
      </p>
    </div>
  );
}

function UpdateContent({ onClose }: { onClose: () => void }) {
  const latest = useUpdateStore((state) => state.latest);
  const status = useUpdateStore((state) => state.status);
  const error = useUpdateStore((state) => state.error);
  const check = useUpdateStore((state) => state.check);

  if (!latest) {
    return (
      <div className="space-y-3">
        <p className="text-[12px] leading-6 text-[var(--mo-text-muted)]">
          {status === 'checking' ? '正在检查更新…' : error ? `检查更新失败：${error}` : '当前已是最新版本。'}
        </p>
        <div className="flex justify-end gap-2">
          <Button size="sm" variant="secondary" onClick={() => void check({ force: true })}>
            重新检查
          </Button>
          <Button size="sm" variant="primary" onClick={onClose}>
            关闭
          </Button>
        </div>
      </div>
    );
  }

  const openRelease = () => {
    // 桌面端交给系统浏览器，避免在 WebView 里打开 GitHub。
    if (isTauriRuntime()) void openExternalUrl(latest.url);
    else window.open(latest.url, '_blank', 'noreferrer');
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center gap-3 rounded-lg bg-[var(--mo-surface-2)] p-3">
        <BrandIcon size={40} />
        <div className="min-w-0">
          <div className="flex flex-wrap items-center gap-1.5 text-[13px] font-semibold">
            {latest.tagName}
            {latest.prerelease ? <Badge tone="accent">预览版</Badge> : null}
          </div>
          <div className="mt-0.5 text-[11px] text-[var(--mo-text-muted)]">
            {APP_VERSION} → {latest.tagName}
            {latest.publishedAt ? ` · ${new Date(latest.publishedAt).toLocaleDateString('zh-CN')}` : ''}
          </div>
        </div>
      </div>

      <p className="text-[12px] leading-6 text-[var(--mo-text-muted)]">
        已发布新版本「{latest.name}」。桌面版需要下载安装包覆盖安装，本地文档、快照与已装插件都会保留。
      </p>

      <div className="flex items-center justify-between gap-2 border-t border-[var(--mo-line)] pt-3">
        <button
          type="button"
          onClick={onClose}
          className="text-[11px] text-[var(--mo-text-muted)] hover:underline"
        >
          稍后提醒（本版本不再弹出）
        </button>
        <div className="flex items-center gap-2">
          <Button size="sm" variant="secondary" onClick={onClose}>
            关闭
          </Button>
          <Button size="sm" variant="primary" icon={<ExternalLink size={13} />} onClick={openRelease}>
            查看更新内容
          </Button>
        </div>
      </div>
    </div>
  );
}

/** 关于页的手动检查入口：自动检查失败时用户还有一条明确的重试路径。 */
function UpdateCheckRow() {
  const status = useUpdateStore((state) => state.status);
  const checkedAt = useUpdateStore((state) => state.checkedAt);
  const latest = useUpdateStore((state) => state.latest);
  const error = useUpdateStore((state) => state.error);
  const check = useUpdateStore((state) => state.check);
  const openDialog = useWorkspaceStore((state) => state.openDialog);

  const detail =
    status === 'checking'
      ? '正在检查…'
      : latest
        ? `发现新版本 ${latest.tagName}`
        : status === 'error'
          ? `检查失败：${error ?? '网络不可用'}`
          : status === 'none'
            ? '已是最新版本'
            : checkedAt
              ? `上次检查 ${new Date(checkedAt).toLocaleString('zh-CN', { hour12: false })}`
              : '尚未检查';

  return (
    <div className="flex items-center justify-between gap-2 rounded-md border border-[var(--mo-line)] px-3 py-2">
      <div className="min-w-0">
        <div className="text-[11px] font-medium text-[var(--mo-text-primary)]">版本更新</div>
        <div className="truncate text-[10px] text-[var(--mo-text-muted)]">{detail}</div>
      </div>
      <div className="flex shrink-0 items-center gap-1.5">
        {latest ? (
          <Button size="sm" variant="primary" onClick={() => openDialog({ type: 'update' })}>
            查看
          </Button>
        ) : null}
        <Button
          size="sm"
          variant="secondary"
          icon={status === 'checking' ? <RefreshCw size={12} className="animate-spin" /> : undefined}
          disabled={status === 'checking'}
          onClick={() => void check({ force: true })}
        >
          检查更新
        </Button>
      </div>
    </div>
  );
}

function AboutContent() {
  const latestRelease = useLatestRelease();  return (
    <div className="space-y-4">
      <div className="flex items-center gap-3 rounded-lg bg-[var(--mo-surface-2)] p-3">
        {/* 功能：关于页使用正式图标并标明出品方；时间：2026-10-18；作者：lq。 */}
        <BrandIcon size={52} />
        <div>
          <div className="text-[15px] font-semibold">Mini Office</div>
          {/* 功能：更正出品方名称；时间：2026-10-24；作者：lq。 */}
          <div className="mt-0.5 text-[12px] font-medium text-[var(--mo-accent)]">非人类程序员出品</div>
          {/* 功能：关于页展示统一版本；时间：2026-10-24；作者：lq。 */}
          <div className="flex flex-wrap items-center gap-x-2 gap-y-1 text-[11px] text-[var(--mo-text-muted)]">
            <span>版本 {APP_VERSION} · 离线优先的轻量办公套件</span>
            <a
              href={latestRelease?.url ?? RELEASES_URL}
              target="_blank"
              rel="noreferrer"
              // 功能：从关于页跳转到新版本发布页，桌面端交由系统浏览器打开；时间：2026-10-24；作者：lq。
              onClick={(event) => {
                if (!isTauriRuntime()) return;
                event.preventDefault();
                void openExternalUrl(RELEASES_URL);
              }}
              className="inline-flex items-center gap-1 rounded px-1 py-0.5 text-[var(--mo-accent)] hover:bg-[var(--mo-surface-2)] hover:underline"
            >
              {latestRelease ? `发现新版本 ${latestRelease.tagName}${latestRelease.prerelease ? '（预览版）' : ''}` : '新版本更新'}
              <ExternalLink size={11} />
            </a>
          </div>
        </div>
      </div>
      <UpdateCheckRow />
      <p className="text-[12px] leading-6 text-[var(--mo-text-muted)]">
        使用 Rust + Tauri 2、React、TypeScript 构建。支持 DOCX 合同编辑、XLSX 表格与公式、
        PPTX 基础内容编辑与高保真演示，并通过能力扫描与安全另存策略保护原始文件。
      </p>
      <div className="grid grid-cols-2 gap-2 text-[11px]">
        <div className="rounded-md border border-[var(--mo-line)] p-2.5">
          <div className="text-[var(--mo-text-muted)]">文档模型</div>
          <div className="mt-0.5 font-medium">OOXML 双模式 · L0–L3</div>
        </div>
        <div className="rounded-md border border-[var(--mo-line)] p-2.5">
          <div className="text-[var(--mo-text-muted)]">数据存储</div>
          <div className="mt-0.5 font-medium">本地快照 · 原子写入</div>
        </div>
        <div className="rounded-md border border-[var(--mo-line)] p-2.5">
          <div className="text-[var(--mo-text-muted)]">公式引擎</div>
          <div className="mt-0.5 font-medium">60+ 常用函数</div>
        </div>
        <div className="rounded-md border border-[var(--mo-line)] p-2.5">
          <div className="text-[var(--mo-text-muted)]">导出</div>
          <div className="mt-0.5 font-medium">PDF · 图片 ZIP</div>
        </div>
      </div>
      {/* 功能：关于页并排展示公众号与反馈群二维码，控制尺寸以兼顾扫码识别和弹窗空间；时间：2026-10-28；作者：lq。 */}
      <div className="grid grid-cols-2 gap-2 rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface-2)] p-2.5">
        <div className="flex min-w-0 flex-col items-center gap-1.5 text-center">
          <img src="/brand/non-human-programmer-wechat.jpg" alt="非人类程序员微信公众号二维码" width={116} height={116} draggable={false} className="size-[116px] bg-white object-contain" style={{ imageRendering: 'pixelated' }} />
          <div className="text-[11px] font-semibold">关注公众号</div>
          <div className="text-[10px] leading-4 text-[var(--mo-text-muted)]">获取更新与技巧</div>
        </div>
        <div className="flex min-w-0 flex-col items-center gap-1.5 text-center">
          <img src="/brand/community-feedback-group.png" alt="Mini Office 反馈交流共建群二维码" width={116} height={116} draggable={false} className="size-[116px] bg-white object-contain" style={{ imageRendering: 'pixelated' }} />
          <div className="text-[11px] font-semibold">加入反馈群</div>
          <div className="text-[10px] leading-4 text-[var(--mo-text-muted)]">交流共建，反馈问题</div>
        </div>
      </div>
      <div className="flex items-center gap-2 text-[10px] text-[var(--mo-text-faint)]">
        <ShieldCheck size={12} /> 不收集文档内容 · 不上传文件 · 本地优先
      </div>
    </div>
  );
}
