/** 功能：默认应用设置面板与首次启动提示，统一 Windows 系统设置与 macOS 直接申请两种路径；时间：2026-10-29；作者：lq。 */
import { useCallback, useEffect, useRef, useState } from 'react';
import {
  Check,
  CircleSlash,
  ExternalLink,
  FileCode2,
  FileSpreadsheet,
  FileText,
  FileType2,
  Presentation,
  Table2,
  type LucideIcon,
} from 'lucide-react';
import { useWorkspaceStore } from '@mini-office/document-model';
import { Button } from '@mini-office/ui';
import {
  getDefaultAppStatus,
  isTauriRuntime,
  openDefaultAppSettings,
  repairFileAssociations,
  requestDefaultApp,
  type DefaultAppStatus,
} from '@mini-office/platform-adapter';

/** 首次启动提示的去重标记，写入 localStorage，卸载重装前不再打扰。 */
const PROMPT_KEY = 'mini-office.default-app-prompt';

const EXTENSION_VISUALS: Record<string, { icon: LucideIcon; tint: string; soft: string }> = {
  docx: { icon: FileText, tint: '#2B579A', soft: '#EEF3FF' },
  xlsx: { icon: Table2, tint: '#217346', soft: '#E9F7EF' },
  pptx: { icon: Presentation, tint: '#C43E1C', soft: '#FCEFE9' },
  csv: { icon: FileSpreadsheet, tint: '#6B5BD6', soft: '#EFEDFC' },
  // 功能：只读格式沿用各自品牌色，避免新类型落进灰色占位样式；时间：2026-11-01；作者：lq。
  pdf: { icon: FileType2, tint: '#C0392B', soft: '#FBECEA' },
  md: { icon: FileCode2, tint: '#2F6FEB', soft: '#ECF2FE' },
  html: { icon: FileCode2, tint: '#D97706', soft: '#FDF3E3' },
};

const FALLBACK_VISUAL = { icon: FileText, tint: 'var(--mo-accent)', soft: 'var(--mo-accent-soft)' };

/** 功能：同族扩展名提示，让用户知道 .markdown / .htm 等写法也一并关联；时间：2026-11-01；作者：lq。 */
const EXTENSION_FAMILY: Record<string, string> = {
  docx: 'docm、dotx',
  xlsx: 'xlsm、xltx',
  pptx: 'pptm、potx',
  md: 'markdown',
  html: 'htm',
};

export function hasDefaultAppPrompted(): boolean {
  try {
    return !!localStorage.getItem(PROMPT_KEY);
  } catch {
    // 存储不可用时不再重复弹窗，避免每次启动打扰。
    return true;
  }
}

export function markDefaultAppPrompted(): void {
  try {
    localStorage.setItem(PROMPT_KEY, String(Date.now()));
  } catch {
    // 忽略存储失败。
  }
}

/** 功能：首次启动且尚未设为默认时弹窗提示，已设置或已提示过则静默跳过；时间：2026-10-29；作者：lq。 */
export async function maybePromptDefaultApp(): Promise<void> {
  if (!isTauriRuntime() || hasDefaultAppPrompted()) return;
  if (useWorkspaceStore.getState().dialog) return;
  const status = await getDefaultAppStatus();
  if (status.allDefault) {
    markDefaultAppPrompted();
    return;
  }
  const store = useWorkspaceStore.getState();
  // 状态查询期间用户可能已打开其他弹窗，此时不抢占焦点。
  if (store.dialog) return;
  store.openDialog({ type: 'default-app', payload: { prompt: true } });
}

export function DefaultAppPanel({
  variant,
  onClose,
}: {
  variant: 'prompt' | 'settings';
  onClose?: () => void;
}) {
  const notify = useWorkspaceStore((state) => state.notify);
  const [status, setStatus] = useState<DefaultAppStatus | null>(null);
  const [pending, setPending] = useState(false);
  const alive = useRef(true);

  const refresh = useCallback(async () => {
    const next = await getDefaultAppStatus();
    if (alive.current) setStatus(next);
  }, []);

  useEffect(() => {
    alive.current = true;
    void refresh();
    // 功能：用户从系统设置返回窗口时自动刷新，避免状态长期停留在“未设置”；时间：2026-10-29；作者：lq。
    const onFocus = () => {
      void refresh();
    };
    window.addEventListener('focus', onFocus);
    return () => {
      alive.current = false;
      window.removeEventListener('focus', onFocus);
    };
  }, [refresh]);

  const apply = async () => {
    if (pending) return;
    setPending(true);
    try {
      const result = await requestDefaultApp();
      if (alive.current) setStatus(result.status);
      if (result.outcome === 'applied') {
        notify('success', '已设为默认应用', result.message);
        onClose?.();
      } else if (result.outcome === 'settings') {
        notify('info', '请在系统设置中确认', result.message);
      } else if (result.outcome === 'error') {
        notify('error', '设置默认应用失败', result.message);
      } else {
        notify('warning', '需要手动确认', result.message);
      }
    } finally {
      if (alive.current) setPending(false);
    }
  };

  const openSettings = async () => {
    try {
      await openDefaultAppSettings();
    } catch (error) {
      notify('error', '无法打开系统设置', error instanceof Error ? error.message : String(error));
    }
  };

  /** 功能：修复文件关联，解决“右键打开方式选了我们却只启动软件”的反馈；时间：2026-10-29；作者：lq。 */
  const repair = async () => {
    if (pending) return;
    setPending(true);
    try {
      const message = await repairFileAssociations();
      notify('success', '文件关联已修复', message);
      await refresh();
    } catch (error) {
      notify('error', '修复文件关联失败', error instanceof Error ? error.message : String(error));
    } finally {
      if (alive.current) setPending(false);
    }
  };

  if (!isTauriRuntime()) {
    return (
      <p className="rounded-md bg-[var(--mo-surface-2)] px-3 py-2 text-[11px] leading-5 text-[var(--mo-text-muted)]">
        默认应用设置仅在桌面客户端可用。安装版会在首次启动时提示，也可随时在设置中调整。
      </p>
    );
  }

  const entries = status?.entries ?? [];
  const isWindows = status?.platform === 'windows';
  const allDefault = status?.allDefault ?? false;

  return (
    <div className="space-y-3">
      {variant === 'prompt' ? (
        <p className="text-[12px] leading-5 text-[var(--mo-text-muted)]">
          设为默认应用后，双击文档、把文件拖到 Mini Office 图标，都会直接用 Mini Office 打开，无需先启动软件再选择文件。
        </p>
      ) : null}

      {/* 功能：类型增加到 6 项后压缩行高，保持单列避免窄弹窗内文字截断；时间：2026-11-01；作者：lq。 */}
      <div className="space-y-1">
        {entries.map((entry) => {
          const visual = EXTENSION_VISUALS[entry.extension] ?? FALLBACK_VISUAL;
          const Icon = visual.icon;
          return (
            <div
              key={entry.extension}
              className="flex items-center gap-2.5 rounded-lg border border-[var(--mo-line)] px-2.5 py-1.5"
            >
              <span
                className="flex h-7 w-7 shrink-0 items-center justify-center rounded-md"
                style={{ background: visual.soft, color: visual.tint }}
              >
                <Icon size={14} />
              </span>
              <div className="min-w-0 flex-1">
                <div className="truncate text-[12px] font-medium text-[var(--mo-text-primary)]">
                  {entry.label}
                </div>
                <div className="truncate text-[10px] text-[var(--mo-text-muted)]">
                  .{entry.extension} 文件
                  {EXTENSION_FAMILY[entry.extension] ? `（含 ${EXTENSION_FAMILY[entry.extension]}）` : ''}
                </div>
              </div>
              {entry.isDefault ? (
                <span className="inline-flex shrink-0 items-center gap-1 rounded-full bg-[var(--mo-success-soft)] px-2 py-0.5 text-[10px] font-medium text-[var(--mo-success)]">
                  <Check size={10} /> 默认
                </span>
              ) : (
                <span className="inline-flex shrink-0 items-center gap-1 rounded-full bg-[var(--mo-surface-2)] px-2 py-0.5 text-[10px] font-medium text-[var(--mo-text-muted)]">
                  <CircleSlash size={10} /> 未设置
                </span>
              )}
            </div>
          );
        })}
        {entries.length === 0 ? (
          <div className="rounded-lg border border-[var(--mo-line)] px-3 py-2 text-[11px] text-[var(--mo-text-muted)]">
            正在读取系统关联状态…
          </div>
        ) : null}
      </div>

      <p className="rounded-md bg-[var(--mo-surface-2)] px-3 py-2 text-[10px] leading-4 text-[var(--mo-text-muted)]">
        {allDefault
          ? `当前已是 ${entries.map((entry) => entry.extension.toUpperCase()).join(' / ')} 的默认打开方式，无需重复设置。`
          : status?.detail ?? '正在读取系统默认应用状态…'}
      </p>

      <div className="flex flex-wrap items-center justify-end gap-2">
        {variant === 'prompt' ? (
          <Button variant="ghost" size="sm" onClick={() => onClose?.()}>
            以后再说
          </Button>
        ) : null}
        {/* 功能：提供一键修复入口，解决“右键打开方式选了我们却只启动软件”的问题；时间：2026-10-29；作者：lq。 */}
        <Button variant="ghost" size="sm" onClick={() => void repair()}>
          修复文件关联
        </Button>
        {isWindows ? (
          <Button variant="secondary" size="sm" icon={<ExternalLink size={12} />} onClick={() => void openSettings()}>
            打开系统设置
          </Button>
        ) : null}
        <Button variant="primary" size="sm" loading={pending} onClick={() => void apply()}>
          {allDefault ? '重新申请默认' : isWindows ? '去系统设置设为默认' : '设为默认应用'}
        </Button>
      </div>
    </div>
  );
}
