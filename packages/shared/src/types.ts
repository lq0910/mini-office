// 功能：扩展只读预览格式，统一纳入工作区标签与文件关联；时间：2026-10-31；作者：lq。
export type OfficeKind = 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'md' | 'html';

export const OFFICE_KINDS: OfficeKind[] = ['docx', 'xlsx', 'pptx', 'pdf', 'md', 'html'];

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
  pdf: { label: 'PDF 文档', shortLabel: 'PDF', extension: '.pdf', mime: 'application/pdf', accept: ['.pdf'], color: '#e34f4f' },
  md: { label: 'Markdown 文档', shortLabel: 'Markdown', extension: '.md', mime: 'text/markdown;charset=utf-8', accept: ['.md', '.markdown'], color: '#6b5bd2' },
  html: { label: 'HTML 页面', shortLabel: 'HTML', extension: '.html', mime: 'text/html;charset=utf-8', accept: ['.html', '.htm'], color: '#e67e22' },
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
  /** 功能：护眼模式——界面与纸张改用低蓝光米黄色，长时间阅读更舒适；时间：2026-10-30；作者：lq。 */
  eyeCare: boolean;
  /** 功能：Word 编辑区顶部标尺（含缩进与制表位标记），默认隐藏，与 Word / WPS 一致；时间：2026-10-30；作者：lq。 */
  wordRuler: boolean;
  /** 功能：Word 编辑页显示正文区域四角裁剪标记，仅作屏幕辅助；时间：2026-11-05；作者：lq。 */
  wordCropMarks: boolean;
  /** 功能：表格阅读模式——十字高亮当前单元格所在行列；时间：2026-10-30；作者：lq。 */
  sheetCrosshair: boolean;
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
  eyeCare: false,
  wordRuler: false,
  wordCropMarks: true,
  sheetCrosshair: false,
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
