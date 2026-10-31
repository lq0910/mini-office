export type OfficeKind = 'docx' | 'xlsx' | 'pptx';

export const OFFICE_KINDS: OfficeKind[] = ['docx', 'xlsx', 'pptx'];

export const OFFICE_META: Record<
  OfficeKind,
  { label: string; shortLabel: string; extension: string; mime: string; accept: string[]; color: string }
> = {
  docx: {
    label: 'Word 文档',
    shortLabel: '文档',
    extension: '.docx',
    mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    accept: ['.docx', '.doc'],
    color: '#2b579a',
  },
  xlsx: {
    label: 'Excel 表格',
    shortLabel: '表格',
    extension: '.xlsx',
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    accept: ['.xlsx', '.xlsm', '.csv'],
    color: '#217346',
  },
  pptx: {
    label: 'PowerPoint 演示文稿',
    shortLabel: '演示',
    extension: '.pptx',
    mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    accept: ['.pptx', '.ppt'],
    color: '#c43e1c',
  },
};

export type WarningSeverity = 'info' | 'warning' | 'blocking';

export interface CompatibilityWarning {
  code: string;
  message: string;
  severity: WarningSeverity;
  detail?: string;
}

/** L0 无法打开 · L1 可预览 · L2 可编辑但需另存 · L3 可验证安全往返保存 */
export type CapabilityLevel = 0 | 1 | 2 | 3;

export interface CapabilityFeature {
  code: string;
  label: string;
  supported: boolean;
  severity: WarningSeverity;
  note?: string;
}

export interface CapabilityReport {
  level: CapabilityLevel;
  warnings: CompatibilityWarning[];
  features: CapabilityFeature[];
  missingFonts: string[];
  scannedAt: number;
}

export interface FontInfo {
  family: string;
  style?: string;
  source: string;
  hasChinese: boolean;
  /** 功能：同一字体族的其他名称（如中英文名），用于缺失字体判断；时间：2026-10-29；作者：lq。 */
  aliases?: string[];
}

export interface RecentEntry {
  id: string;
  kind: OfficeKind;
  displayName: string;
  path?: string;
  size: number;
  openedAt: number;
  updatedAt: number;
  dirty: boolean;
  level: CapabilityLevel;
}

export interface DocumentSnapshotMeta {
  id: string;
  kind: OfficeKind;
  displayName: string;
  path?: string;
  size: number;
  updatedAt: number;
  dirty: boolean;
  level: CapabilityLevel;
}

export interface DocumentSnapshot extends DocumentSnapshotMeta {
  bytes: Uint8Array;
}

export type ThemeMode = 'light' | 'dark' | 'system';

/** Built-in accent palettes. `custom` uses `AppSettings.customAccent`. */
export type AccentId = 'mist' | 'ocean' | 'iris' | 'pine' | 'clay' | 'custom';

export interface AppSettings {
  theme: ThemeMode;
  accent: AccentId;
  /** Seed color for the `custom` accent, stored as #RRGGBB. */
  customAccent: string;
  /** Native macOS frosted-glass material behind the window. */
  vibrancy: boolean;
  language: 'zh-CN' | 'en-US';
  autoSave: boolean;
  autoSaveDelayMs: number;
  showToolbar: boolean;
  showProperties: boolean;
  defaultZoom: number;
  restoreSession: boolean;
  fontSize: 'small' | 'default' | 'large';
  /** 护眼模式：编辑区纸张/表面切换为护眼色。 */
  eyeProtection?: boolean;
  /** 高亮当前行/列（表格编辑时）。 */
  highlightRowCol?: boolean;
}

export const DEFAULT_ACCENT_HEX = '#2E7D72';

export const DEFAULT_SETTINGS: AppSettings = {
  theme: 'system',
  accent: 'mist',
  customAccent: DEFAULT_ACCENT_HEX,
  vibrancy: false,
  language: 'zh-CN',
  autoSave: true,
  autoSaveDelayMs: 1400,
  showToolbar: true,
  showProperties: false,
  defaultZoom: 1,
  restoreSession: true,
  fontSize: 'default',
  eyeProtection: false,
  highlightRowCol: false,
};

export interface DocumentStats {
  pages?: number;
  words?: number;
  characters?: number;
  cells?: number;
  sheets?: number;
  slides?: number;
}

export interface ExportOptions {
  fileName: string;
  pageRange?: string;
  scale?: number;
  includeNotes?: boolean;
}

export interface ExportProgress {
  phase: 'preparing' | 'rendering' | 'assembling' | 'saving' | 'done' | 'error';
  current: number;
  total: number;
  message: string;
}
