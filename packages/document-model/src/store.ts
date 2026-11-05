import { create } from 'zustand';
import { getPlatformAdapter } from '@mini-office/platform-adapter';
import {
  DEFAULT_SETTINGS,
  ErrorCodes,
  decodeDisplayName,
  OFFICE_META,
  OfficeError,
  ensureExtension,
  extensionOf,
  safeJsonParse,
  stripExtension,
  uid,
  type AppSettings,
  type CapabilityReport,
  type DocumentSnapshot,
  type DocumentSnapshotMeta,
  type FontInfo,
  type OfficeKind,
  type RecentEntry,
  type WarningSeverity,
} from '@mini-office/shared';
import type {
  AnyDocumentData,
  AnyFormatHandler,
  DocumentSession,
  FormatOpenResult,
  FormatRegistry,
  LifecycleState,
  OpenDocumentInput,
  SessionView,
} from './session';
import { initialViewFor } from './session';
import type { WorkbookModel } from './sheet';
import type { DocModel } from './word';
import type { SlideDeck } from './slides';

const SETTINGS_KEY = 'mini-office.settings';
/** 功能：设置项默认值调整的一次性迁移标记，只影响升级后的第一次读取；时间：2026-10-30；作者：lq。 */
const SETTINGS_MIGRATION_KEY = 'mini-office.settings.migrated.v1';
const THEME_KEY = 'mini-office.theme';
// 功能：让系统文件选择器与拖拽入口识别三种只读预览格式；时间：2026-10-31；作者：lq。
const ALL_EXTENSIONS = ['.docx', '.docm', '.dotx', '.xlsx', '.xlsm', '.xltx', '.pptx', '.pptm', '.csv', '.doc', '.xls', '.ppt', '.pdf', '.md', '.markdown', '.html', '.htm'];
const KIND_BY_EXTENSION: Record<string, OfficeKind> = {
  '.docx': 'docx',
  '.docm': 'docx',
  '.dotx': 'docx',
  '.doc': 'docx',
  '.xlsx': 'xlsx',
  '.xlsm': 'xlsx',
  '.xltx': 'xlsx',
  '.xls': 'xlsx',
  '.csv': 'xlsx',
  '.pptx': 'pptx',
  '.pptm': 'pptx',
  '.potx': 'pptx',
  '.ppt': 'pptx',
  '.pdf': 'pdf',
  '.md': 'md',
  '.markdown': 'md',
  '.html': 'html',
  '.htm': 'html',
};

const handlers: FormatRegistry = {};
const sourceBuffers = new Map<string, Uint8Array>();
const fileHandles = new Map<string, FileSystemFileHandle>();
const snapshotTimers = new Map<string, ReturnType<typeof setTimeout>>();
const snapshotInFlight = new Set<string>();
// 功能：独立记录写盘任务，避免编辑状态改变后重复发起保存；时间：2026-10-18；作者：lq。
const saveInFlight = new Set<string>();

export function registerFormatHandler(handler: AnyFormatHandler): void {
  handlers[handler.kind] = handler;
}

export function getFormatHandler(kind: OfficeKind): AnyFormatHandler | undefined {
  return handlers[kind];
}

export interface OpenBytesInput {
  name: string;
  bytes: Uint8Array;
  path?: string;
  handle?: FileSystemFileHandle;
  restored?: { dirty: boolean };
  /**
   * 功能：旧版格式（DOC/XLS/PPT）转换后的原始文件名；存在时按只读兼容会话打开，保存默认另存为；时间：2026-11-02；作者：lq。
   */
  sourceName?: string;
  /** 功能：携带旧格式转换器报告的兼容警告；时间：2026-11-03；作者：lq。 */
  conversionWarnings?: string[];
}

export type NotificationLevel = 'info' | 'success' | 'warning' | 'error';

export interface AppNotification {
  id: string;
  level: NotificationLevel;
  title: string;
  message?: string;
  createdAt: number;
}

export type DialogType =
  | 'settings'
  | 'export'
  | 'compatibility'
  | 'about'
  | 'shortcuts'
  | 'plugins'
  // 功能：新增默认应用设置弹窗，首次启动提示与设置页共用；时间：2026-10-29；作者：lq。
  | 'default-app'
  // 功能：发现新版本后的启动提示弹窗；时间：2026-10-29；作者：lq。
  | 'update';

export interface DialogState {
  type: DialogType;
  sessionId?: string;
  payload?: Record<string, unknown>;
}

export interface WorkspaceState {
  sessions: Record<string, DocumentSession>;
  tabOrder: string[];
  activeId: string | null;
  settings: AppSettings;
  recent: RecentEntry[];
  snapshots: DocumentSnapshotMeta[];
  fonts: FontInfo[];
  hydrated: boolean;
  busy: string | null;
  dialog: DialogState | null;
  commandPaletteOpen: boolean;
  notifications: AppNotification[];

  initialize(): Promise<void>;
  /** 功能：后台及文档打开共用同一次字体扫描；时间：2026-10-24；作者：lq。 */
  loadFonts(): Promise<FontInfo[]>;
  /** 功能：重新扫描系统字体，安装新字体后无需重启应用；时间：2026-10-29；作者：lq。 */
  refreshFonts(): Promise<FontInfo[]>;
  setBusy(message: string | null): void;
  openFileDialog(): Promise<string | null>;
  openBytes(input: OpenBytesInput): Promise<string | null>;
  openPath(path: string): Promise<string | null>;
  openDroppedFile(file: File): Promise<string | null>;
  openSnapshot(id: string): Promise<string | null>;
  createDocument(kind: OfficeKind): Promise<string | null>;
  activateTab(id: string): void;
  reorderTabs(fromId: string, toId: string): void;
  closeTab(id: string, force?: boolean): Promise<boolean>;
  closeAll(force?: boolean): Promise<void>;
  updateData(
    id: string,
    data: AnyDocumentData | ((current: AnyDocumentData) => AnyDocumentData),
    options?: { dirty?: boolean },
  ): void;
  updateDocMeta(id: string, metaPatch: Partial<DocModel['meta']> | ((current: DocModel['meta']) => DocModel['meta']), options?: { dirty?: boolean }): void;
  updateView(id: string, patch: SessionView): void;
  saveSession(id: string, options?: { asNew?: boolean }): Promise<boolean>;
  saveActive(options?: { asNew?: boolean }): Promise<boolean>;
  exportBytes(bytes: Uint8Array, fileName: string, mime: string): Promise<void>;
  updateSettings(patch: Partial<AppSettings>): void;
  refreshSnapshots(): Promise<void>;
  deleteSnapshot(id: string): Promise<void>;
  clearSnapshots(): Promise<void>;
  openDialog(dialog: DialogState): void;
  closeDialog(): void;
  setCommandPalette(open: boolean): void;
  notify(level: NotificationLevel, title: string, message?: string): void;
  dismissNotification(id: string): void;
  getSession(id: string | null | undefined): DocumentSession | undefined;
}

function emptyCapability(): CapabilityReport {
  return { level: 3, warnings: [], features: [], missingFonts: [], scannedAt: Date.now() };
}

function loadSettings(): AppSettings {
  if (typeof localStorage === 'undefined') return { ...DEFAULT_SETTINGS };
  const stored = safeJsonParse<Partial<AppSettings>>(localStorage.getItem(SETTINGS_KEY), {});
  const merged = { ...DEFAULT_SETTINGS, ...stored };
  // 功能：Word 标尺改为默认隐藏（与 Word / WPS 一致），旧存档里被写成 true 的一次性改回，
  // 否则改默认值对已存过设置的用户无效，看起来像没生效；用户之后手动开启仍会被正常保存。
  // 时间：2026-10-30；作者：lq。
  try {
    if (!localStorage.getItem(SETTINGS_MIGRATION_KEY)) {
      merged.wordRuler = DEFAULT_SETTINGS.wordRuler;
      localStorage.setItem(SETTINGS_MIGRATION_KEY, '1');
    }
  } catch {
    // 迁移失败不影响启动，最多是这一次仍按旧值显示。
  }
  return merged;
}

function persistSettings(settings: AppSettings): void {
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
    // The inline bootstrap script reads this key before React mounts.
    localStorage.setItem(THEME_KEY, JSON.stringify(settings.theme));
  } catch {
    // Ignore quota failures.
  }
}

function defaultDocumentBase(kind: OfficeKind): string {
  if (kind === 'docx') return '未命名文档';
  if (kind === 'xlsx') return '未命名工作簿';
  return '未命名演示文稿';
}

function uniqueDisplayName(kind: OfficeKind, sessions: Record<string, DocumentSession>): string {
  const base = defaultDocumentBase(kind);
  const extension = OFFICE_META[kind].extension;
  const used = new Set(
    Object.values(sessions)
      .filter((session) => session.kind === kind)
      .map((session) => stripExtension(session.displayName)),
  );
  if (!used.has(base)) return `${base}${extension}`;
  let index = 2;
  while (used.has(`${base} ${index}`)) index += 1;
  return `${base} ${index}${extension}`;
}

function isCsvSession(session: DocumentSession): boolean {
  return session.kind === 'xlsx' && (session.data as WorkbookModel).csv === true;
}

function saveNameFor(session: DocumentSession): string {
  const base = stripExtension(session.displayName || defaultDocumentBase(session.kind));
  const extension = isCsvSession(session) ? '.csv' : OFFICE_META[session.kind].extension;
  return ensureExtension(base, extension);
}

function mimeForName(kind: OfficeKind, name: string): string {
  if (extensionOf(name) === '.csv') return 'text/csv;charset=utf-8';
  return OFFICE_META[kind].mime;
}

async function findHandler(name: string, bytes: Uint8Array): Promise<AnyFormatHandler | null> {
  const extension = extensionOf(name);
  const direct = KIND_BY_EXTENSION[extension];
  if (direct && handlers[direct]) return handlers[direct];
  let pkg;
  try {
    // 功能：仅对扩展名未知的文件加载容器探测代码；时间：2026-10-24；作者：lq。
    const { detectKind, loadOoxmlPackage } = await import('@mini-office/ooxml');
    pkg = await loadOoxmlPackage(bytes);
    const detection = await detectKind(pkg, name);
    return handlers[detection.kind] ?? null;
  } catch {
    return null;
  } finally {
    pkg?.dispose();
  }
}

function severityRank(severity: WarningSeverity): number {
  return severity === 'blocking' ? 2 : severity === 'warning' ? 1 : 0;
}

export const useWorkspaceStore = create<WorkspaceState>()((set, get) => {
  // 功能：合并 StrictMode 重挂载和并发打开触发的初始化、字体扫描；时间：2026-10-24；作者：lq。
  let initializing: Promise<void> | undefined;
  let fontsLoading: Promise<FontInfo[]> | undefined;
  const patchSession = (
    id: string,
    patch: Partial<DocumentSession> | ((session: DocumentSession) => Partial<DocumentSession>),
  ) => {
    set((state) => {
      const session = state.sessions[id];
      if (!session) return state;
      const resolved = typeof patch === 'function' ? patch(session) : patch;
      return {
        sessions: {
          ...state.sessions,
          [id]: { ...session, ...resolved, updatedAt: Date.now() },
        },
      };
    });
  };

  const refreshSnapshots = async () => {
    try {
      const metas = await getPlatformAdapter().listSnapshots();
      const recent: RecentEntry[] = metas.map((meta) => ({
        id: meta.id,
        kind: meta.kind,
        displayName: meta.displayName,
        path: meta.path,
        size: meta.size,
        openedAt: meta.updatedAt,
        updatedAt: meta.updatedAt,
        dirty: meta.dirty,
        level: meta.level,
      }));
      set({ snapshots: metas, recent });
    } catch (error) {
      console.warn('[mini-office] snapshot list failed', error);
    }
  };

  const persistSnapshot = async (
    id: string,
    options: { dirty?: boolean; bytes?: Uint8Array } = {},
  ): Promise<void> => {
    if (snapshotInFlight.has(id)) return;
    const session = get().sessions[id];
    if (!session) return;
    const handler = handlers[session.kind];
    snapshotInFlight.add(id);
    try {
      const dirty = options.dirty ?? session.dirty;
      let bytes = options.bytes ?? sourceBuffers.get(id);
      if (!bytes || dirty) {
        if (!handler) return;
        bytes = await handler.serialize(session.data, {
          sourceBytes: bytes,
          name: session.displayName,
        });
      }
      if (!bytes) return;
      const snapshot: DocumentSnapshot = {
        id,
        kind: session.kind,
        displayName: session.displayName,
        path: session.path,
        size: bytes.byteLength,
        updatedAt: Date.now(),
        dirty,
        level: session.capability.level,
        bytes,
      };
      await getPlatformAdapter().saveSnapshot(snapshot);
      await refreshSnapshots();
    } catch (error) {
      console.warn('[mini-office] snapshot failed', error);
    } finally {
      snapshotInFlight.delete(id);
    }
  };

  const scheduleSnapshot = (id: string) => {
    const { settings } = get();
    if (!settings.autoSave) return;
    const existing = snapshotTimers.get(id);
    if (existing) clearTimeout(existing);
    const delay = Math.max(500, settings.autoSaveDelayMs);
    snapshotTimers.set(
      id,
      setTimeout(() => {
        snapshotTimers.delete(id);
        void persistSnapshot(id, { dirty: true });
      }, delay),
    );
  };

  const removeSnapshotFor = (id: string) => {
    void getPlatformAdapter()
      .deleteSnapshot(id)
      .then(() => refreshSnapshots());
  };

  return {
    sessions: {},
    tabOrder: [],
    activeId: null,
    settings: loadSettings(),
    recent: [],
    snapshots: [],
    fonts: [],
    hydrated: false,
    busy: null,
    dialog: null,
    commandPaletteOpen: false,
    notifications: [],

    getSession(id) {
      return id ? get().sessions[id] : undefined;
    },

    async initialize() {
      if (get().hydrated) return;
      // 功能：首页仅等待最近文件元数据，字体扫描不再阻塞工作区就绪；时间：2026-10-24；作者：lq。
      initializing ??= refreshSnapshots().then(() => { set({ hydrated: true }); });
      await initializing;
    },

    loadFonts() {
      // 功能：首次打开文档可等待真实字体结果，失败后允许重试，避免误报缺失字体；时间：2026-10-24；作者：lq。
      fontsLoading ??= Promise.resolve().then(() => getPlatformAdapter().listFonts()).then((fonts) => {
        const available = fonts ?? [];
        set({ fonts: available });
        return available;
      }).catch(() => {
        fontsLoading = undefined;
        return [] as FontInfo[];
      });
      return fontsLoading;
    },

    /**
     * 功能：强制重新扫描系统字体，解决“用户新装字体仍提示缺失”的反馈；时间：2026-10-29；作者：lq。
     * 桌面端走原生刷新命令，Web 端退化为重新枚举一次。
     */
    async refreshFonts() {
      fontsLoading = undefined;
      const adapter = getPlatformAdapter();
      try {
        const fonts = adapter.refreshFonts ? await adapter.refreshFonts() : await adapter.listFonts();
        set({ fonts: fonts ?? [] });
        fontsLoading = Promise.resolve(fonts ?? []);
        return fonts ?? [];
      } catch {
        return get().loadFonts();
      }
    },

    setBusy(message) {
      set({ busy: message });
    },

    async openFileDialog() {
      try {
        const adapter = getPlatformAdapter();
        const opened = await adapter.openFile(ALL_EXTENSIONS);
        if (!opened) return null;
        return get().openBytes({
          name: opened.name,
          bytes: opened.bytes,
          path: opened.path,
          handle: opened.handle,
          sourceName: opened.sourceName,
          conversionWarnings: opened.conversionWarnings,
        });
      } catch (error) {
        get().notify(
          'error',
          '打开文件失败',
          error instanceof Error ? error.message : String(error),
        );
        return null;
      }
    },

    async openBytes(input) {
      const adapter = getPlatformAdapter();
      // 功能：路径、拖拽、Web 字节入口统一使用保留字体/样式的 DOC 转换器，失败不静默改走有损通道；时间：2026-11-03；作者：lq。
      if (/\.doc$/i.test(input.name)) {
        try {
          const { convertLegacyDoc } = await import('../../platform-adapter/src/legacy-doc');
          return get().openBytes(await convertLegacyDoc(input.bytes, input.name));
        } catch (error) {
          get().notify('error', '旧版 DOC 转换失败', error instanceof Error ? error.message : String(error));
          return null;
        }
      }
      // 功能：旧版 DOC/XLS/PPT 在内存中转换为 OOXML，源目录零写入，并以只读兼容会话打开；时间：2026-11-02；作者：lq。
      if (input.path && /\.(doc|xls|ppt)$/i.test(input.name) && adapter.convertLegacy) {
        try {
          const converted = await adapter.convertLegacy(input.path);
          if (converted) {
            return get().openBytes({
              name: converted.name,
              bytes: converted.bytes,
              sourceName: converted.sourceName,
              conversionWarnings: converted.conversionWarnings,
            });
          }
        } catch (error) {
          get().notify('error', '旧版文件转换失败', error instanceof Error ? error.message : String(error));
          return null;
        }
      }
      const handler = await findHandler(input.name, input.bytes);
      if (!handler) {
        get().notify(
          'error',
          '不支持的格式',
          '无法识别该文件。V1 支持 DOCX、XLSX、PPTX，以及 CSV。',
        );
        return null;
      }
      const id = uid('doc');
      const now = Date.now();
      const displayName = decodeDisplayName(input.name);
      // 功能：标签从占位阶段就显示原始旧格式文件名，避免先闪出 .docx 再改名；时间：2026-11-02；作者：lq。
      const legacyDisplayName = input.sourceName ? decodeDisplayName(input.sourceName) : undefined;
      const placeholder: DocumentSession = {
        id,
        kind: handler.kind,
        displayName: legacyDisplayName ?? displayName,
        path: input.path,
        readonly: false,
        capability: emptyCapability(),
        state: 'VALIDATING',
        data: handler.createBlank(),
        dirty: false,
        isNew: false,
        restoredFromSnapshot: !!input.restored,
        canSaveInPlace: false,
        saveAsCopyRequired: false,
        sourceSize: input.bytes.byteLength,
        createdAt: now,
        updatedAt: now,
        view: initialViewFor(handler.kind),
      };
      set((state) => ({
        sessions: { ...state.sessions, [id]: placeholder },
        tabOrder: [...state.tabOrder, id],
        activeId: id,
        busy: `正在解析 ${displayName}…`,
        dialog: null,
      }));

      try {
        // 功能：只在实际打开文档时等待字体扫描，保持兼容性检查使用完整字体清单；时间：2026-10-24；作者：lq。
        const fonts = await get().loadFonts();
        const openInput: OpenDocumentInput = {
          name: displayName,
          bytes: input.bytes,
          path: input.path,
          // 功能：把字体族的中英文别名一并纳入已知字体，避免已安装字体被误报缺失；时间：2026-10-29；作者：lq。
          knownFonts: fonts.flatMap((font) => [font.family, ...(font.aliases ?? [])]),
        };
        const result: FormatOpenResult = await handler.open(openInput);
        // 功能：旧版格式转换产物只用于解析，标签仍显示用户原始文件名，且不绑定原路径，避免把 DOCX 写回 .doc；时间：2026-11-02；作者：lq。
        const legacySource = input.sourceName && /\.(doc|xls|ppt)$/i.test(input.sourceName) ? input.sourceName : undefined;
        const readonly = result.readonly === true || result.capability.level <= 1;
        const capability = legacySource
          ? {
              ...result.capability,
              warnings: [
                {
                  code: 'legacy.compat',
                  severity: 'info' as const,
                  message: `${legacySource} 以兼容模式打开`,
                  detail: '旧版格式经兼容通道转换，复杂版式可能与原文件不同；保存时会另存为 OOXML，不会覆盖原文件。',
                },
                ...(input.conversionWarnings ?? []).map((detail) => ({ code: 'legacy.loss', severity: 'warning' as const, message: '旧格式转换存在兼容限制', detail })),
                ...result.capability.warnings,
              ],
            }
          : result.capability;
        const restoredDirty = input.restored?.dirty === true;
        const nextState: LifecycleState = readonly
          ? 'READ_ONLY'
          : restoredDirty
            ? 'DIRTY'
            : 'EDIT_READY';
        const session: DocumentSession = {
          ...placeholder,
          data: result.data,
          capability,
          readonly,
          state: nextState,
          dirty: restoredDirty,
          displayName: legacySource
            ? decodeDisplayName(legacySource)
            : decodeDisplayName(result.displayName ?? displayName),
          // 功能：兼容会话不绑定源路径，保存只能另存为 OOXML，杜绝把 DOCX 字节写回 .doc；时间：2026-11-02；作者：lq。
          path: restoredDirty || legacySource ? undefined : input.path,
          canSaveInPlace:
            !readonly &&
            !legacySource &&
            !restoredDirty &&
            result.capability.level >= 3 &&
            !!(input.handle || input.path),
          saveAsCopyRequired:
            !readonly && (!!legacySource || restoredDirty || result.capability.level === 2),
          isNew: false,
          updatedAt: Date.now(),
        };
        sourceBuffers.set(id, input.bytes.slice());
        if (input.handle) fileHandles.set(id, input.handle);
        set((state) => ({
          sessions: { ...state.sessions, [id]: session },
          busy: null,
        }));
        void persistSnapshot(id, {
          dirty: restoredDirty,
          bytes: input.bytes,
        });
        // 功能：一次性告知兼容模式与保存去向，避免用户以为原 .doc 会被直接改写；时间：2026-11-02；作者：lq。
        if (legacySource) {
          get().notify(
            'info',
            '以兼容模式打开',
            `${legacySource} 已转换打开，复杂版式可能与原文件不同；保存时将另存为 ${saveNameFor(session)}。`,
          );
        }

        const needsAttention = result.capability.warnings.some(
          (warning) =>
            warning.severity === 'blocking' ||
            (warning.severity === 'warning' &&
              warning.code !== 'fonts.missing' &&
              warning.code !== 'pptx.readonly'),
        );
        if (needsAttention) {
          set({ dialog: { type: 'compatibility', sessionId: id } });
        }
        return id;
      } catch (error) {
        set((state) => {
          const sessions = { ...state.sessions };
          delete sessions[id];
          const tabOrder = state.tabOrder.filter((tabId) => tabId !== id);
          return {
            sessions,
            tabOrder,
            activeId: state.activeId === id ? tabOrder.at(-1) ?? null : state.activeId,
            busy: null,
          };
        });
        const officeError = error instanceof OfficeError ? error : null;
        get().notify(
          'error',
          '无法打开文档',
          officeError?.message ?? (error instanceof Error ? error.message : String(error)),
        );
        if (officeError?.detail) console.error('[mini-office]', officeError.detail);
        return null;
      }
    },

    async openPath(path) {
      // 功能：访达重复打开已打开文件时激活标签，避免覆盖未保存编辑；时间：2026-10-18；作者：lq。
      const existing = Object.values(get().sessions).find((session) => session.path === path);
      if (existing) {
        get().activateTab(existing.id);
        return existing.id;
      }
      try {
        const opened = await getPlatformAdapter().openPath(path);
        if (!opened) {
          get().notify('error', '无法打开文件', '当前平台不支持通过路径打开文件。');
          return null;
        }
        return get().openBytes({ name: opened.name, bytes: opened.bytes, path: opened.path, sourceName: opened.sourceName, conversionWarnings: opened.conversionWarnings });
      } catch (error) {
        get().notify('error', '打开文件失败', error instanceof Error ? error.message : String(error));
        return null;
      }
    },

    async openDroppedFile(file) {
      try {
        const adapter = getPlatformAdapter();
        const opened = adapter.readDroppedFile
          ? await adapter.readDroppedFile(file)
          : {
              name: file.name,
              bytes: new Uint8Array(await file.arrayBuffer()),
            };
        return get().openBytes({ name: opened.name, bytes: opened.bytes, path: opened.path, sourceName: opened.sourceName, conversionWarnings: opened.conversionWarnings });
      } catch (error) {
        get().notify('error', '打开文件失败', error instanceof Error ? error.message : String(error));
        return null;
      }
    },

    async openSnapshot(id) {
      try {
        const adapter = getPlatformAdapter();
        const snapshot = await adapter.loadSnapshot(id);
        if (!snapshot) {
          get().notify('warning', '恢复快照不存在', '该快照可能已被清理。');
          await refreshSnapshots();
          return null;
        }
        const opened = await get().openBytes({
          name: snapshot.displayName,
          bytes: snapshot.bytes,
          path: snapshot.dirty ? undefined : snapshot.path,
          restored: { dirty: snapshot.dirty },
        });
        if (opened) {
          await adapter.deleteSnapshot(id);
          await refreshSnapshots();
        }
        return opened;
      } catch (error) {
        get().notify('error', '恢复失败', error instanceof Error ? error.message : String(error));
        return null;
      }
    },

    async createDocument(kind) {
      const handler = handlers[kind];
      if (!handler) {
        get().notify('error', '模块未加载', `尚未注册 ${OFFICE_META[kind].label} 编辑器。`);
        return null;
      }
      const id = uid('doc');
      const now = Date.now();
      const session: DocumentSession = {
        id,
        kind,
        displayName: uniqueDisplayName(kind, get().sessions),
        readonly: false,
        capability: emptyCapability(),
        state: 'EDIT_READY',
        data: handler.createBlank(),
        dirty: false,
        isNew: true,
        restoredFromSnapshot: false,
        canSaveInPlace: false,
        saveAsCopyRequired: false,
        sourceSize: 0,
        createdAt: now,
        updatedAt: now,
        view: initialViewFor(kind),
      };
      set((state) => ({
        sessions: { ...state.sessions, [id]: session },
        tabOrder: [...state.tabOrder, id],
        activeId: id,
        dialog: null,
      }));
      void persistSnapshot(id, { dirty: false });
      return id;
    },

    activateTab(id) {
      if (!get().sessions[id]) return;
      set({ activeId: id, dialog: null });
    },

    reorderTabs(fromId, toId) {
      set((state) => {
        const order = [...state.tabOrder];
        const from = order.indexOf(fromId);
        const to = order.indexOf(toId);
        if (from < 0 || to < 0 || from === to) return state;
        order.splice(from, 1);
        order.splice(to, 0, fromId);
        return { tabOrder: order };
      });
    },

    async closeTab(id, force = false) {
      const session = get().sessions[id];
      if (!session) return true;
      if (session.dirty && !force) return false;
      const timer = snapshotTimers.get(id);
      if (timer) clearTimeout(timer);
      snapshotTimers.delete(id);
      const cleanNew = session.isNew && !session.dirty;
      set((state) => {
        const sessions = { ...state.sessions };
        delete sessions[id];
        const tabOrder = state.tabOrder.filter((tabId) => tabId !== id);
        let activeId = state.activeId;
        if (activeId === id) {
          const oldIndex = state.tabOrder.indexOf(id);
          activeId = tabOrder[Math.min(oldIndex, tabOrder.length - 1)] ?? null;
        }
        return { sessions, tabOrder, activeId };
      });
      sourceBuffers.delete(id);
      fileHandles.delete(id);
      if (cleanNew) removeSnapshotFor(id);
      return true;
    },

    async closeAll(force = false) {
      const ids = [...get().tabOrder];
      for (const id of ids) {
        const ok = await get().closeTab(id, force);
        if (!ok) break;
      }
    },

    updateData(id, data, options) {
      const session = get().sessions[id];
      if (!session || session.readonly) return;
      const next = typeof data === 'function' ? data(session.data) : data;
      // 功能：默认不置脏，只有用户编辑动作才显式传入 dirty；避免统计刷新、重排等内部写入把文档标成已修改；时间：2026-11-02；作者：lq。
      const dirty = session.dirty || options?.dirty === true;
      patchSession(id, {
        data: next,
        dirty,
        // 功能：写盘期间允许编辑，但保持忙状态供退出保护判断；时间：2026-10-18；作者：lq。
        state: saveInFlight.has(id) ? 'SAVING' : dirty ? 'DIRTY' : 'EDIT_READY',
      });
      if (dirty) scheduleSnapshot(id);
    },

    // 功能：只更新文档 meta，不触碰正文；邮件合并面板状态/Word 兼容信息写入这里；时间：2026-10-29；作者：lq。
    updateDocMeta(id, metaPatch, options) {
      const session = get().sessions[id];
      if (!session || session.readonly) return;
      const wordData = session.data as DocModel;
      const next = typeof metaPatch === 'function' ? metaPatch(wordData.meta) : { ...wordData.meta, ...metaPatch };
      // 功能：面板状态等非正文写入默认不置脏，只有正文语义变化才由调用方显式声明；时间：2026-11-02；作者：lq。
      const dirty = session.dirty || options?.dirty === true;
      patchSession(id, { data: { ...session.data, meta: next }, dirty, state: saveInFlight.has(id) ? 'SAVING' : dirty ? 'DIRTY' : 'EDIT_READY' });
      if (dirty) scheduleSnapshot(id);
    },

    updateView(id, patch) {
      const session = get().sessions[id];
      if (!session) return;
      patchSession(id, { view: { ...session.view, ...patch } });
    },

    async saveSession(id, options) {
      const session = get().sessions[id];
      if (!session) return false;
      // 功能：阻止并发保存，确保关闭确认等待真实保存结果；时间：2026-10-18；作者：lq。
      if (saveInFlight.has(id) || session.state === 'SAVING') return false;
      if (session.readonly) {
        get().notify(
          'warning',
          '文档为只读',
          '该文档以保真只读模式打开，无法直接保存。请先在原应用中处理不受支持的内容。',
        );
        return false;
      }
      const handler = handlers[session.kind];
      if (!handler) {
        get().notify('error', '保存失败', '对应格式的处理模块尚未加载。');
        return false;
      }
      const previousState = session.state;
      saveInFlight.add(id);
      patchSession(id, { state: 'SAVING', saveError: undefined, updatedAt: Date.now() });
      try {
        const bytes = await handler.serialize(session.data, {
          sourceBytes: sourceBuffers.get(id),
          name: session.displayName,
        });
        const suggestedName = saveNameFor(session);
        const handle = fileHandles.get(id);
        const forcePicker =
          options?.asNew === true ||
          session.saveAsCopyRequired ||
          (!session.path && !handle);
        const outcome = await getPlatformAdapter().saveBytes(
          bytes,
          {
            suggestedName,
            path: session.path,
            handle,
            forcePicker,
          },
          mimeForName(session.kind, suggestedName),
        );
        const nextHandle = outcome.handle ?? handle;
        if (nextHandle) fileHandles.set(id, nextHandle);
        sourceBuffers.set(id, bytes.slice());
        const displayName = outcome.displayName || suggestedName;
        // 功能：保存期间继续编辑时保留脏标记，防止关闭窗口丢失新内容；时间：2026-10-18；作者：lq。
        const changedDuringSave = get().sessions[id]?.data !== session.data;
        patchSession(id, {
          displayName,
          path: outcome.path ?? (outcome.location === 'in-place' ? session.path : undefined),
          dirty: changedDuringSave,
          state: changedDuringSave ? 'DIRTY' : 'SAVED',
          lastSavedAt: Date.now(),
          saveError: undefined,
          saveAsCopyRequired: false,
          canSaveInPlace:
            outcome.location !== 'download' && !!(nextHandle || outcome.path || session.path),
        });
        await persistSnapshot(id, { dirty: changedDuringSave, bytes });
        get().notify(
          'success',
          outcome.location === 'download' ? '已下载副本' : '已保存',
          `${displayName}`,
        );
        return true;
      } catch (error) {
        const officeError = error instanceof OfficeError ? error : null;
        if (officeError?.code === ErrorCodes.CANCELLED) {
          patchSession(id, { state: get().sessions[id]?.dirty ? 'DIRTY' : previousState });
          return false;
        }
        patchSession(id, {
          state: 'SAVE_FAILED',
          saveError: officeError?.message ?? (error instanceof Error ? error.message : String(error)),
        });
        get().notify(
          'error',
          '保存失败',
          officeError?.message ?? (error instanceof Error ? error.message : String(error)),
        );
        return false;
      } finally {
        saveInFlight.delete(id);
      }
    },

    async saveActive(options) {
      const id = get().activeId;
      if (!id) return false;
      return get().saveSession(id, options);
    },

    async exportBytes(bytes, fileName, mime) {
      try {
        await getPlatformAdapter().saveBytes(
          bytes,
          { suggestedName: fileName, forcePicker: true },
          mime,
        );
        get().notify('success', '导出完成', fileName);
      } catch (error) {
        if ((error as OfficeError)?.code === ErrorCodes.CANCELLED) return;
        get().notify('error', '导出失败', error instanceof Error ? error.message : String(error));
      }
    },

    updateSettings(patch) {
      const next = { ...get().settings, ...patch };
      persistSettings(next);
      set({ settings: next });
    },

    async refreshSnapshots() {
      await refreshSnapshots();
    },

    async deleteSnapshot(id) {
      await getPlatformAdapter().deleteSnapshot(id);
      await refreshSnapshots();
    },

    async clearSnapshots() {
      const adapter = getPlatformAdapter();
      const metas = await adapter.listSnapshots();
      await Promise.all(metas.map((meta) => adapter.deleteSnapshot(meta.id)));
      await refreshSnapshots();
    },

    openDialog(dialog) {
      set({ dialog, commandPaletteOpen: false });
    },

    closeDialog() {
      set({ dialog: null });
    },

    setCommandPalette(open) {
      set({ commandPaletteOpen: open });
    },

    notify(level, title, message) {
      const notification: AppNotification = {
        id: uid('toast'),
        level,
        title,
        message,
        createdAt: Date.now(),
      };
      set((state) => ({
        notifications: [...state.notifications.slice(-4), notification],
      }));
    },

    dismissNotification(id) {
      set((state) => ({
        notifications: state.notifications.filter((notification) => notification.id !== id),
      }));
    },
  };
});

export function selectActiveSession(state: WorkspaceState): DocumentSession | undefined {
  return state.activeId ? state.sessions[state.activeId] : undefined;
}

export function activeDocument<T extends AnyDocumentData>(state: WorkspaceState): DocumentSession<T> | undefined {
  return selectActiveSession(state) as DocumentSession<T> | undefined;
}

export function sessionDocument<T extends AnyDocumentData>(session: DocumentSession): T {
  return session.data as T;
}

/** Convenience types for editor packages. */
export type WordSession = DocumentSession<DocModel>;
export type SheetSession = DocumentSession<WorkbookModel>;
export type SlidesSession = DocumentSession<SlideDeck>;

export function getSourceBytes(id: string): Uint8Array | undefined {
  return sourceBuffers.get(id);
}

export function getFileHandle(id: string): FileSystemFileHandle | undefined {
  return fileHandles.get(id);
}

export function capabilityNeedsAttention(report: CapabilityReport): boolean {
  return report.warnings.some((warning) => severityRank(warning.severity) > 0);
}

export function coverageLabel(report: CapabilityReport): string {
  if (report.level === 0) return 'L0 无法打开';
  if (report.level === 1) return 'L1 仅预览';
  if (report.level === 2) return 'L2 需另存';
  return 'L3 安全往返';
}
