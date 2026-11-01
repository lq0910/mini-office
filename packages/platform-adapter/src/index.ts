import { getCurrentWindow } from '@tauri-apps/api/window';
import { isTauriRuntime, supportsWindowVibrancy } from './types';
import { WebPlatformAdapter } from './web';
import { TauriPlatformAdapter } from './tauri';
import type { PlatformAdapter } from './types';

export * from './types';
export * from './fonts';
export * from './snapshots';
export { WebPlatformAdapter } from './web';
export { TauriPlatformAdapter } from './tauri';

let adapter: PlatformAdapter | null = null;

export function getPlatformAdapter(): PlatformAdapter {
  if (adapter) return adapter;
  adapter = isTauriRuntime() ? new TauriPlatformAdapter() : new WebPlatformAdapter();
  return adapter;
}

/** Test/embedded hosts can inject a fake adapter. */
export function setPlatformAdapter(next: PlatformAdapter | null): void {
  adapter = next;
}

export function platformName(): 'web' | 'tauri' {
  return getPlatformAdapter().kind;
}

// 功能：在桌面端使用系统默认浏览器打开外部链接；时间：2026-10-24；作者：lq。
export async function openExternalUrl(url: string): Promise<void> {
  if (isTauriRuntime()) {
    try {
      const { openUrl } = await import('@tauri-apps/plugin-opener');
      await openUrl(url);
      return;
    } catch (error) {
      console.warn('打开外部链接失败', error);
    }
  }
  window.open(url, '_blank', 'noopener,noreferrer');
}

export async function startWindowDrag(): Promise<void> {
  if (!isTauriRuntime()) return;
  try {
    await getCurrentWindow().startDragging();
  } catch (error) {
    console.warn('窗口拖动失败', error);
  }
}

// 功能：Windows / Linux 桌面端改用自绘标题栏，需要自行提供最小化、最大化与关闭；时间：2026-11-01；作者：lq。
export async function minimizeWindow(): Promise<void> {
  if (!isTauriRuntime()) return;
  try {
    await getCurrentWindow().minimize();
  } catch (error) {
    console.warn('窗口最小化失败', error);
  }
}

export async function closeWindow(): Promise<void> {
  if (!isTauriRuntime()) return;
  try {
    await getCurrentWindow().close();
  } catch (error) {
    console.warn('窗口关闭失败', error);
  }
}

export async function isWindowMaximized(): Promise<boolean> {
  if (!isTauriRuntime()) return false;
  try {
    return await getCurrentWindow().isMaximized();
  } catch {
    return false;
  }
}

/**
 * 功能：去掉系统标题栏，让文档标签直接并入窗口顶部一行（Windows 顶栏合并）；
 * 时间：2026-11-01；作者：lq。
 *
 * 只在非 macOS 桌面端调用：macOS 使用 Overlay 标题栏保留红绿灯按钮。
 */
export async function applyCustomTitleBar(enabled: boolean): Promise<boolean> {
  if (!isTauriRuntime()) return false;
  try {
    await getCurrentWindow().setDecorations(!enabled);
    return true;
  } catch (error) {
    console.warn('切换自绘标题栏失败', error);
    return false;
  }
}

export async function setPresentationFullscreen(enabled: boolean): Promise<boolean> {
  if (isTauriRuntime()) {
    try {
      await getCurrentWindow().setFullscreen(enabled);
      return true;
    } catch (error) {
      console.warn('切换窗口全屏失败', error);
      return false;
    }
  }
  try {
    if (enabled) {
      if (!document.fullscreenElement && document.documentElement.requestFullscreen) {
        await document.documentElement.requestFullscreen({ navigationUI: 'hide' });
      }
      return true;
    }
    if (document.fullscreenElement) await document.exitFullscreen();
    return true;
  } catch (error) {
    console.warn('浏览器全屏请求失败', error);
    return false;
  }
}

let windowVibrancyApplied = false;

/**
 * Apply (or hide) the native macOS frosted-glass material behind the webview.
 *
 * The window itself is created transparent on macOS, so turning the effect off
 * only needs to restore opaque CSS surfaces. The native view is intentionally
 * kept alive for an instant re-enable and because older Tauri runtimes cannot
 * reliably remove it at runtime.
 */
export async function applyWindowVibrancy(enabled: boolean): Promise<boolean> {
  if (!supportsWindowVibrancy()) return false;
  if (!enabled) return true;
  if (windowVibrancyApplied) return true;
  try {
    const { Effect, EffectState } = await import('@tauri-apps/api/window');
    await getCurrentWindow().setEffects({
      effects: [Effect.UnderWindowBackground],
      state: EffectState.Active,
    });
    windowVibrancyApplied = true;
    return true;
  } catch (error) {
    console.warn('启用毛玻璃效果失败', error);
    return false;
  }
}

export async function toggleWindowMaximize(): Promise<void> {
  if (!isTauriRuntime()) return;
  try {
    await getCurrentWindow().toggleMaximize();
  } catch (error) {
    console.warn('窗口最大化失败', error);
  }
}

// 功能：打印能力统一收敛到 print.ts，Word 与 Excel 共用同一套分页与纸张规则；时间：2026-10-29；作者：lq。
export * from './print';

// 功能：导出桌面文档关联和退出保护桥接；时间：2026-10-18；作者：lq。
export * from './desktop-events';
// 功能：导出默认应用状态查询与申请桥接；时间：2026-10-29；作者：lq。
export * from './default-app';
