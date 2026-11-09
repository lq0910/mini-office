import type {
  CapabilityReport,
  DocumentStats,
  OfficeKind,
} from '@mini-office/shared';
import type { DocModel } from './word';
import type { WorkbookModel } from './sheet';
import type { SlideDeck } from './slides';
import type { WordCoreSession } from '../../word-core/src/session/word-session';

/** 功能：工作区只保留新 Word 文档标识，正文与 DOM 均归独立会话管理；时间：2026-11-09；作者：lq。 */
export interface CoreWordDocument { readonly engine: 'word-core'; }

// 功能：保存 PDF、Markdown、HTML 的只读预览内容，避免引入重量级编辑模型；时间：2026-10-31；作者：lq。
export interface PreviewDocumentData {
  format: 'pdf' | 'md' | 'html';
  bytes?: Uint8Array;
  text?: string;
}

export type AnyDocumentData = DocModel | WorkbookModel | SlideDeck | PreviewDocumentData | CoreWordDocument;
export type AnyDocumentDataWithPreview = AnyDocumentData;

export type LifecycleState =
  | 'UNOPENED'
  | 'VALIDATING'
  | 'LOADING'
  | 'VIEW_READY'
  | 'READ_ONLY'
  | 'EDIT_READY'
  | 'DIRTY'
  | 'SAVING'
  | 'SAVED'
  | 'SAVE_FAILED'
  | 'EXPORTING'
  | 'ERROR';

export type LifecycleEvent =
  | 'VALIDATE'
  | 'LOAD'
  | 'VIEW'
  | 'READONLY'
  | 'EDIT'
  | 'CHANGE'
  | 'SAVE'
  | 'SAVED'
  | 'SAVE_FAILED'
  | 'EXPORT'
  | 'EXPORT_DONE'
  | 'ERROR'
  | 'RESET';

const TRANSITIONS: Record<LifecycleState, Partial<Record<LifecycleEvent, LifecycleState>>> = {
  UNOPENED: { VALIDATE: 'VALIDATING' },
  VALIDATING: { LOAD: 'LOADING', ERROR: 'ERROR' },
  LOADING: { VIEW: 'VIEW_READY', READONLY: 'READ_ONLY', EDIT: 'EDIT_READY', ERROR: 'ERROR' },
  VIEW_READY: {
    EDIT: 'EDIT_READY',
    READONLY: 'READ_ONLY',
    CHANGE: 'DIRTY',
    EXPORT: 'EXPORTING',
    ERROR: 'ERROR',
  },
  READ_ONLY: { CHANGE: 'READ_ONLY', EXPORT: 'EXPORTING', ERROR: 'ERROR' },
  EDIT_READY: { CHANGE: 'DIRTY', SAVE: 'SAVING', EXPORT: 'EXPORTING', READONLY: 'READ_ONLY', ERROR: 'ERROR' },
  DIRTY: { SAVE: 'SAVING', CHANGE: 'DIRTY', EXPORT: 'EXPORTING', ERROR: 'ERROR' },
  SAVING: { SAVED: 'SAVED', SAVE_FAILED: 'SAVE_FAILED', ERROR: 'ERROR' },
  SAVED: { CHANGE: 'DIRTY', EDIT: 'EDIT_READY', SAVE: 'SAVING', EXPORT: 'EXPORTING', ERROR: 'ERROR' },
  SAVE_FAILED: { SAVE: 'SAVING', CHANGE: 'DIRTY', EDIT: 'EDIT_READY', ERROR: 'ERROR' },
  EXPORTING: { EXPORT_DONE: 'VIEW_READY', ERROR: 'ERROR' },
  ERROR: { RESET: 'VIEW_READY', VALIDATE: 'VALIDATING' },
};

export function transition(from: LifecycleState, event: LifecycleEvent): LifecycleState {
  return TRANSITIONS[from]?.[event] ?? from;
}

export function canEdit(state: LifecycleState): boolean {
  return ['VIEW_READY', 'EDIT_READY', 'DIRTY', 'SAVED', 'SAVE_FAILED', 'SAVING'].includes(state);
}

export function isBusyState(state: LifecycleState): boolean {
  return state === 'VALIDATING' || state === 'LOADING' || state === 'SAVING' || state === 'EXPORTING';
}

export interface SessionView {
  zoom?: number;
  activeSheetId?: string;
  slideIndex?: number;
  showNotes?: boolean;
  presenting?: boolean;
  mode?: 'pages' | 'continuous';
  page?: number;
  pageCount?: number;
}

export interface DocumentSession<T extends AnyDocumentDataWithPreview = AnyDocumentDataWithPreview> {
  id: string;
  kind: OfficeKind;
  displayName: string;
  path?: string;
  readonly: boolean;
  capability: CapabilityReport;
  state: LifecycleState;
  data: T;
  /** 功能：新 Word 会话只在内存运行，恢复快照始终保存 DOCX 字节；时间：2026-11-09；作者：lq。 */
  wordCore?: WordCoreSession;
  dirty: boolean;
  isNew: boolean;
  restoredFromSnapshot: boolean;
  canSaveInPlace: boolean;
  /** True when the file contains features that require saving a copy first. */
  saveAsCopyRequired: boolean;
  sourceSize: number;
  createdAt: number;
  updatedAt: number;
  lastSavedAt?: number;
  saveError?: string;
  view: SessionView;
}

export interface OpenDocumentInput {
  name: string;
  bytes: Uint8Array;
  path?: string;
  knownFonts?: string[];
}

export interface FormatOpenResult<T extends AnyDocumentDataWithPreview = AnyDocumentDataWithPreview> {
  data: T;
  /** 功能：完整新核心会话直接交给工作区，不降级转换成旧 DocModel；时间：2026-11-09；作者：lq。 */
  wordCore?: WordCoreSession;
  capability: CapabilityReport;
  readonly?: boolean;
  displayName?: string;
}

export interface SerializeContext {
  sourceBytes?: Uint8Array;
  name: string;
}

export interface FormatHandler<T extends AnyDocumentDataWithPreview = AnyDocumentDataWithPreview> {
  readonly kind: OfficeKind;
  readonly extensions: string[];
  createBlank(): T;
  /** 功能：允许新核心异步创建原始 DOCX 包和真实编辑会话；时间：2026-11-09；作者：lq。 */
  createNew?(): Promise<FormatOpenResult<T>>;
  open(input: OpenDocumentInput): Promise<FormatOpenResult<T>>;
  serialize(data: T, context: SerializeContext): Promise<Uint8Array>;
  getStats?(data: T): DocumentStats;
}

export type AnyFormatHandler = FormatHandler<AnyDocumentDataWithPreview>;

export type FormatRegistry = Partial<Record<OfficeKind, AnyFormatHandler>>;

export function initialViewFor(kind: OfficeKind): SessionView {
  if (kind === 'xlsx') return { zoom: 1, activeSheetId: undefined };
  if (kind === 'pptx') return { slideIndex: 0, showNotes: false, presenting: false };
  return { zoom: 1, mode: 'pages', page: 0 };
}

export function stateAfterOpen(readonly: boolean, hasEdits = false): LifecycleState {
  if (readonly) return 'READ_ONLY';
  return hasEdits ? 'DIRTY' : 'EDIT_READY';
}

export function fileBaseName(name: string): string {
  const slash = Math.max(name.lastIndexOf('/'), name.lastIndexOf('\\'));
  return slash >= 0 ? name.slice(slash + 1) : name;
}
