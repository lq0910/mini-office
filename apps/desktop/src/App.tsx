import { useEffect } from 'react';
import { useWorkspaceStore } from '@mini-office/document-model';
import {
  applyCustomTitleBar,
  applyWindowVibrancy,
  isTauriRuntime,
  isWindows,
  supportsWindowVibrancy,
  watchDesktopFiles,
} from '@mini-office/platform-adapter';
import { AppShell, maybePromptDefaultApp } from '@mini-office/editor-shell';
import { applyAppearance } from '@mini-office/design-system';
import { usePluginRegistry } from '@mini-office/plugin-host';
// 功能：使用静态启动面板，待工作区首页完成首帧绘制后移除；时间：2026-10-24；作者：lq。
import { finishStartup } from './startup';

export default function App() {
  const hydrated = useWorkspaceStore((state) => state.hydrated);
  const initialize = useWorkspaceStore((state) => state.initialize);
  const settings = useWorkspaceStore((state) => state.settings);
  const openPath = useWorkspaceStore((state) => state.openPath);
  const initializePlugins = usePluginRegistry((state) => state.initialize);

  useEffect(() => {
    void initialize();
    // 功能：恢复已安装插件并应用当前主题包，需在首帧后完成；时间：2026-10-29；作者：lq。
    void initializePlugins();
    document.documentElement.dataset.platform = isTauriRuntime() ? 'tauri' : 'web';
  }, [initialize, initializePlugins]);

  // 功能：外观变化不再重复初始化工作区；时间：2026-10-24；作者：lq。
  useEffect(() => {
    applyAppearance(settings);
  }, [settings.theme, settings.accent, settings.customAccent]);

  // 功能：双帧确认首页已布局绘制，不用固定延迟延长用户等待；时间：2026-10-24；作者：lq。
  useEffect(() => {
    if (hydrated) return finishStartup();
  }, [hydrated]);

  // 功能：首次启动且未设为默认应用时提示一次，等启动面板移除后再弹窗避免遮挡；时间：2026-10-29；作者：lq。
  useEffect(() => {
    if (!hydrated) return;
    const timer = setTimeout(() => {
      void maybePromptDefaultApp();
    }, 800);
    return () => clearTimeout(timer);
  }, [hydrated]);

  useEffect(() => {
    if (!supportsWindowVibrancy()) {
      document.documentElement.dataset.vibrancy = 'off';
      return;
    }
    if (!settings.vibrancy) {
      document.documentElement.dataset.vibrancy = 'off';
      void applyWindowVibrancy(false);
      return;
    }
    let cancelled = false;
    void applyWindowVibrancy(true).then((applied) => {
      if (cancelled) return;
      document.documentElement.dataset.vibrancy = applied ? 'on' : 'off';
      if (!applied) {
        useWorkspaceStore
          .getState()
          .notify('warning', '毛玻璃效果不可用', '当前窗口无法启用原生磨砂材质，已保持不透明背景。');
      }
    });
    return () => {
      cancelled = true;
    };
  }, [settings.vibrancy]);

  useEffect(() => {
    if (!isTauriRuntime()) return;
    let unlisten: (() => void) | undefined;
    let cancelled = false;
    void (async () => {
      try {
        const { getCurrentWebview } = await import('@tauri-apps/api/webview');
        const stop = await getCurrentWebview().onDragDropEvent((event) => {
          const payload = event.payload;
          if (payload.type === 'drop' && payload.paths.length > 0) {
            void openPath(payload.paths[0]);
          }
        });
        if (cancelled) stop();
        else unlisten = stop;
      } catch {
        // Web fallback handles drag and drop when the Tauri API is unavailable.
      }
    })();
    return () => {
      cancelled = true;
      unlisten?.();
    };
  }, [openPath]);

  // 功能：Windows 端去掉系统标题栏，改由应用顶栏承载标签与窗口按钮；时间：2026-11-01；作者：lq。
  useEffect(() => {
    if (!isTauriRuntime() || !isWindows()) return;
    void applyCustomTitleBar(true).then((applied) => {
      document.documentElement.dataset.customTitlebar = applied ? 'on' : 'off';
    });
  }, []);

  // 功能：工作区初始化后领取访达双击文件，监听注册失败时明确提示；时间：2026-10-18；作者：lq。
  useEffect(() => {
    if (!hydrated || !isTauriRuntime()) return;
    let disposed = false;
    let stop: (() => void) | undefined;
    void watchDesktopFiles(openPath).then((unlisten) => {
      if (disposed) unlisten();
      else stop = unlisten;
    }).catch((error) => useWorkspaceStore.getState().notify('error', '文件关联监听失败', String(error)));
    return () => { disposed = true; stop?.(); };
  }, [hydrated, openPath]);

  if (!hydrated) {
    // 功能：保留 HTML 启动面板，覆盖 React 下载、执行和工作区初始化全过程；时间：2026-10-24；作者：lq。
    return null;
  }

  return <AppShell />;
}
