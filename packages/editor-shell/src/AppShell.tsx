import { lazy, Suspense, useCallback, useEffect, useRef, useState } from 'react';
import { useWorkspaceStore, type WordSession, type SheetSession, type SlidesSession } from '@mini-office/document-model';
import { applyAppearance, watchSystemTheme } from '@mini-office/design-system';
import { ToastViewport, TooltipHost } from '@mini-office/ui';
import { extensionOf } from '@mini-office/shared';
import {
  BusyIndicator,
  CloseConfirmDialog,
  DropOverlay,
  EditorLoading,
  StatusBar,
  TabBar,
  TitleBar,
} from './chrome';
import { CommandPalette } from './CommandPalette';
import { MenuBar } from './MenuBar';
import { dispatchPluginShortcut } from './plugin-macro';
import { checkForUpdates, useUpdateStore } from './update-check';
import { useChromeVariant } from './chrome-variant';
import { WorkspaceDialogs } from './dialogs';
import { HomePage } from './HomePage';
import { watchDesktopClose, finishDesktopClose } from '@mini-office/platform-adapter';
import { confirmWorkspaceClose, type CloseChoice } from './close-workspace';

// 功能：仅在打开对应文档时加载编辑界面，首页不再执行三种编辑器；时间：2026-10-24；作者：lq。
const WordEditor = lazy(() => import('../../editor-word/src/WordEditor').then((module) => ({ default: module.WordEditor })));
const SheetEditor = lazy(() => import('../../editor-sheet/src/SheetEditor').then((module) => ({ default: module.SheetEditor })));
const SlidesEditor = lazy(() => import('../../editor-slides/src/SlidesEditor').then((module) => ({ default: module.SlidesEditor })));

export function AppShell() {
  const sessions = useWorkspaceStore((state) => state.sessions);
  const tabOrder = useWorkspaceStore((state) => state.tabOrder);
  const activeId = useWorkspaceStore((state) => state.activeId);
  const settings = useWorkspaceStore((state) => state.settings);
  const dialog = useWorkspaceStore((state) => state.dialog);
  const closeTab = useWorkspaceStore((state) => state.closeTab);
  const saveSession = useWorkspaceStore((state) => state.saveSession);
  const saveActive = useWorkspaceStore((state) => state.saveActive);
  const openFileDialog = useWorkspaceStore((state) => state.openFileDialog);
  const openDroppedFile = useWorkspaceStore((state) => state.openDroppedFile);
  const createDocument = useWorkspaceStore((state) => state.createDocument);
  const setCommandPalette = useWorkspaceStore((state) => state.setCommandPalette);
  const commandPaletteOpen = useWorkspaceStore((state) => state.commandPaletteOpen);
  const closeDialog = useWorkspaceStore((state) => state.closeDialog);
  const openDialog = useWorkspaceStore((state) => state.openDialog);
  const notifications = useWorkspaceStore((state) => state.notifications);
  const dismissNotification = useWorkspaceStore((state) => state.dismissNotification);
  const notify = useWorkspaceStore((state) => state.notify);

  const [showHome, setShowHome] = useState(!activeId);
  // 功能：主题插件可切换外壳变体，经典风格需在标题栏下多渲染一行菜单栏；时间：2026-10-29；作者：lq。
  const chrome = useChromeVariant();
  const [closeRequestId, setCloseRequestId] = useState<string | null>(null);
  const [closing, setClosing] = useState(false);
  const [dragging, setDragging] = useState(false);
  const previousActive = useRef<string | null>(activeId);
  const dragDepth = useRef(0);
  // 功能：复用保存弹窗处理整个应用退出，阻止重复退出请求；时间：2026-10-18；作者：lq。
  const nativeCloseRunning = useRef(false);
  const nativeCloseChoice = useRef<((choice: CloseChoice) => void) | null>(null);

  const cancelClose = useCallback(() => {
    if (useWorkspaceStore.getState().sessions[closeRequestId ?? '']?.state === 'SAVING') return;
    nativeCloseChoice.current?.('cancel');
    nativeCloseChoice.current = null;
    setCloseRequestId(null);
  }, [closeRequestId]);

  // 功能：关窗、程序坞退出和 Cmd+Q 逐份确认，保存失败和取消均保留窗口；时间：2026-10-18；作者：lq。
  useEffect(() => {
    let disposed = false;
    let unlisten: (() => void) | undefined;
    const request = async () => {
      if (nativeCloseRunning.current) return;
      nativeCloseRunning.current = true;
      try {
        const approved = await confirmWorkspaceClose(
          () => Object.values(useWorkspaceStore.getState().sessions),
          (session) => new Promise<CloseChoice>((resolve) => {
            nativeCloseChoice.current = resolve;
            setCloseRequestId(session.id);
          }),
          async (id) => {
            setClosing(true);
            try { return await useWorkspaceStore.getState().saveSession(id); }
            finally { setClosing(false); }
          },
        );
        if (approved && !disposed) await finishDesktopClose();
      } catch (error) {
        useWorkspaceStore.getState().notify('error', '暂未退出应用', String(error));
      } finally {
        nativeCloseRunning.current = false;
        nativeCloseChoice.current = null;
        setCloseRequestId(null);
      }
    };
    void watchDesktopClose(() => { void request(); }).then((stop) => {
      if (disposed) stop();
      else unlisten = stop;
    }).catch((error) => useWorkspaceStore.getState().notify('error', '退出保护监听失败', String(error)));
    return () => {
      disposed = true;
      unlisten?.();
      nativeCloseChoice.current?.('cancel');
    };
  }, []);

  const activeSession = activeId ? sessions[activeId] : undefined;
  const closeRequestSession = closeRequestId ? sessions[closeRequestId] : undefined;

  const requestClose = useCallback(
    (id: string) => {
      const session = useWorkspaceStore.getState().sessions[id];
      if (!session) return;
      if (session.dirty) setCloseRequestId(id);
      else void useWorkspaceStore.getState().closeTab(id);
    },
    [],
  );

  useEffect(() => {
    if (activeId && activeId !== previousActive.current) setShowHome(false);
    if (!activeId) setShowHome(true);
    previousActive.current = activeId;
  }, [activeId]);

  useEffect(() => {
    applyAppearance(settings);
    return watchSystemTheme(settings.theme);
  }, [settings.theme, settings.accent, settings.customAccent]);

  useEffect(() => {
    document.documentElement.dataset.fontSize = settings.fontSize;
  }, [settings.fontSize]);

  // 功能：护眼模式——编辑区表面切换为护眼色；时间：2026-10-30；作者：hou0。
  useEffect(() => {
    document.documentElement.dataset.eyeProtection = settings.eyeProtection ? 'on' : 'off';
  }, [settings.eyeProtection]);

  useEffect(() => {
    document.title = activeSession ? `${activeSession.displayName} — Mini Office` : 'Mini Office';
  }, [activeSession]);

  useEffect(() => {
    const hasDirty = Object.values(sessions).some((session) => session.dirty);
    const handler = (event: BeforeUnloadEvent) => {
      if (!hasDirty) return;
      event.preventDefault();
      event.returnValue = '';
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [sessions]);

  // 功能：启动后异步检查新版本，命中时弹一次提示；时间：2026-10-29；作者：lq。
  // 检查本身不阻塞首屏，网络失败也不打扰用户——小红点不亮，关于页仍可手动重试。
  const latestRelease = useUpdateStore((state) => state.latest);
  const dismissedTag = useUpdateStore((state) => state.dismissedTag);
  const promptedUpdate = useRef<string | null>(null);

  useEffect(() => {
    void checkForUpdates();
  }, []);

  useEffect(() => {
    if (!latestRelease) return;
    if (dismissedTag === latestRelease.tagName) return;
    if (promptedUpdate.current === latestRelease.tagName) return;
    // 已有弹窗时先让用户处理完，避免叠窗；关闭后本 effect 会再次触发。
    if (dialog) return;
    promptedUpdate.current = latestRelease.tagName;
    openDialog({ type: 'update' });
  }, [dialog, dismissedTag, latestRelease, openDialog]);

  useEffect(() => {
    const handler = (event: KeyboardEvent) => {
      const mod = event.metaKey || event.ctrlKey;
      const target = event.target as HTMLElement | null;
      const inEditable =
        !!target &&
        (target.tagName === 'INPUT' ||
          target.tagName === 'TEXTAREA' ||
          target.tagName === 'SELECT' ||
          target.isContentEditable);
      if (mod && event.key.toLowerCase() === 'k') {
        event.preventDefault();
        setCommandPalette(!commandPaletteOpen);
        return;
      }
      if (mod && event.key.toLowerCase() === 's') {
        event.preventDefault();
        if (activeId) {
          if (event.shiftKey) void saveActive({ asNew: true });
          else void saveActive();
        }
        return;
      }
      if (mod && event.key.toLowerCase() === 'o') {
        event.preventDefault();
        void openFileDialog();
        return;
      }
      if (mod && event.key.toLowerCase() === 'n' && !event.shiftKey) {
        event.preventDefault();
        void createDocument('docx');
        return;
      }
      if (mod && event.key === ',') {
        event.preventDefault();
        openDialog({ type: 'settings' });
        return;
      }
      if (mod && event.key.toLowerCase() === 'w') {
        event.preventDefault();
        if (activeId) requestClose(activeId);
        return;
      }
      if (mod && event.key.toLowerCase() === 'p') {
        event.preventDefault();
        if (activeId) window.dispatchEvent(new CustomEvent('mo:print'));
        return;
      }
      if (mod && event.key.toLowerCase() === 'f') {
        event.preventDefault();
        if (activeId) window.dispatchEvent(new CustomEvent('mo:find'));
        return;
      }
      if (event.key === 'F5') {
        event.preventDefault();
        if (activeSession?.kind === 'pptx') window.dispatchEvent(new CustomEvent('mo:present'));
        return;
      }
      if (event.key === 'Escape') {
        if (commandPaletteOpen) {
          setCommandPalette(false);
          return;
        }
        if (dialog) {
          closeDialog();
          return;
        }
        if (closeRequestId) {
          cancelClose();
          return;
        }
      }
      if (mod && !inEditable && /^[1-9]$/.test(event.key)) {
        const index = Number(event.key) - 1;
        const id = tabOrder[index];
        if (id) {
          event.preventDefault();
          useWorkspaceStore.getState().activateTab(id);
          setShowHome(false);
        }
        return;
      }
      // 插件宏快捷键放在最后：内置命令都在上方提前 return，因此内置始终优先。
      if (!event.defaultPrevented) dispatchPluginShortcut(event, activeSession);
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [
    activeId,
    activeSession,
    activeSession?.kind,
    closeRequestId,
    closeDialog,
    commandPaletteOpen,
    createDocument,
    dialog,
    openDialog,
    openFileDialog,
    saveActive,
    setCommandPalette,
    tabOrder,
    requestClose,
    cancelClose,
  ]);

  const discardAndClose = async () => {
    if (!closeRequestId) return;
    // 功能：应用退出时暂存逐份选择，全部确认前不移除任何标签；时间：2026-10-18；作者：lq。
    if (closing) return;
    if (nativeCloseChoice.current) {
      nativeCloseChoice.current('discard');
      nativeCloseChoice.current = null;
      setCloseRequestId(null);
      return;
    }
    await closeTab(closeRequestId, true);
    setCloseRequestId(null);
  };

  const saveAndClose = async () => {
    if (!closeRequestId) return;
    if (closing) return;
    if (nativeCloseChoice.current) {
      nativeCloseChoice.current('save');
      nativeCloseChoice.current = null;
      setCloseRequestId(null);
      return;
    }
    setClosing(true);
    const saved = await saveSession(closeRequestId);
    setClosing(false);
    if (saved && !useWorkspaceStore.getState().sessions[closeRequestId]?.dirty) {
      await closeTab(closeRequestId, true);
      setCloseRequestId(null);
    }
  };

  const handleDrop = async (event: React.DragEvent) => {
    event.preventDefault();
    dragDepth.current = 0;
    setDragging(false);
    const files = Array.from(event.dataTransfer.files);
    if (files.length === 0) return;
    const supported = files.find((file) =>
      ['.docx', '.docm', '.xlsx', '.xlsm', '.pptx', '.pptm', '.csv', '.doc', '.xls', '.ppt'].includes(
        extensionOf(file.name),
      ),
    );
    if (!supported) {
      notify('warning', '不支持的文件类型', '请选择 DOCX、XLSX、PPTX 或 CSV 文件。');
      return;
    }
    if (files.length > 1) notify('info', '一次打开一个文件', `已选择 ${supported.name}`);
    await openDroppedFile(supported);
  };

  const content = (() => {
    if (showHome || !activeSession) return <HomePage />;
    if (activeSession.state === 'VALIDATING' || activeSession.state === 'LOADING') {
      return <EditorLoading title={activeSession.displayName} />;
    }
    if (activeSession.kind === 'docx') {
      return <WordEditor key={activeSession.id} session={activeSession as WordSession} />;
    }
    if (activeSession.kind === 'xlsx') {
      return <SheetEditor key={activeSession.id} session={activeSession as SheetSession} />;
    }
    return <SlidesEditor key={activeSession.id} session={activeSession as SlidesSession} />;
  })();

  return (
    <div
      className="mo-app flex h-full flex-col overflow-hidden bg-[var(--mo-bg)]"
      onDragEnter={(event) => {
        if (!event.dataTransfer.types.includes('Files')) return;
        dragDepth.current += 1;
        setDragging(true);
      }}
      onDragOver={(event) => {
        if (event.dataTransfer.types.includes('Files')) event.preventDefault();
      }}
      onDragLeave={() => {
        dragDepth.current = Math.max(0, dragDepth.current - 1);
        if (dragDepth.current === 0) setDragging(false);
      }}
      onDrop={(event) => void handleDrop(event)}
    >
      <TitleBar
        onHome={() => setShowHome(true)}
        onOpenPalette={() => setCommandPalette(true)}
      />
      {chrome.toolbar === 'classic-two-row' ? <MenuBar onHome={() => setShowHome(true)} /> : null}
      {tabOrder.length > 0 ? (
        <TabBar
          onRequestClose={requestClose}
          onHome={() => setShowHome(true)}
          onSelectTab={() => setShowHome(false)}
        />
      ) : null}
      {/* 功能：首次加载编辑器时保留标题栏、标签栏和清晰的加载反馈；时间：2026-10-24；作者：lq。 */}
      <main className="min-h-0 flex-1">
        <Suspense fallback={<EditorLoading title={activeSession?.displayName ?? '文档'} />}>{content}</Suspense>
      </main>
      <StatusBar />
      <BusyIndicator />
      <DropOverlay visible={dragging} />
      <CommandPalette
        open={commandPaletteOpen}
        onClose={() => setCommandPalette(false)}
        onHome={() => {
          setShowHome(true);
          useWorkspaceStore.getState().closeDialog();
        }}
        onRequestClose={requestClose}
      />
      <WorkspaceDialogs />
      {closeRequestSession ? (
        <CloseConfirmDialog
          session={closeRequestSession}
          saving={closing}
          onCancel={cancelClose}
          onDiscard={() => void discardAndClose()}
          onSave={() => void saveAndClose()}
        />
      ) : null}
      <TooltipHost />
      <ToastViewport
        toasts={notifications.map((notification) => ({
          id: notification.id,
          level: notification.level,
          title: notification.title,
          message: notification.message,
        }))}
        onDismiss={dismissNotification}
      />
    </div>
  );
}
