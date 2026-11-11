import type { DocumentSnapshot, DocumentSnapshotMeta, FontInfo } from '@mini-office/shared';

export interface OpenedFile {
  name: string;
  bytes: Uint8Array;
  path?: string;
  handle?: FileSystemFileHandle;
  lastModified?: number;
  /**
   * 功能：旧版格式转换后的原始文件名（如 合同.doc）；存在即表示这是只读兼容会话；时间：2026-11-02；作者：lq。
   */
  sourceName?: string;
  /** 功能：旧格式转换检测到的兼容缺口，打开时向用户展示；时间：2026-11-03；作者：lq。 */
  conversionWarnings?: string[];
  /** 功能：false 表示仅识别到扩展名不符的 OOXML，未执行旧格式转换；时间：2026-11-11；作者：lq。 */
  conversionPerformed?: boolean;
}

/** 功能：旧版 DOC/XLS/PPT 的内存转换结果，源目录不产生任何文件；时间：2026-11-02；作者：lq。 */
export interface ConvertedLegacy {
  name: string;
  bytes: Uint8Array;
  sourceName: string;
  /** 功能：记录转换器报告的非致命格式损失；时间：2026-11-03；作者：lq。 */
  conversionWarnings?: string[];
  /** 功能：区分真实转换和原字节打开，避免把容器识别误报为损失；时间：2026-11-11；作者：lq。 */
  conversionPerformed?: boolean;
}

export interface SaveTarget {
  suggestedName: string;
  handle?: FileSystemFileHandle;
  path?: string;
  /** Always show the OS picker instead of writing back to the known location. */
  forcePicker?: boolean;
}

export interface SaveOutcome {
  displayName: string;
  path?: string;
  handle?: FileSystemFileHandle;
  location: 'in-place' | 'disk' | 'download';
}

export interface PlatformAdapter {
  readonly kind: 'web' | 'tauri';
  openFile(extensions: string[]): Promise<OpenedFile | null>;
  openPath(path: string): Promise<OpenedFile | null>;
  /** 功能：将旧版 Office 文件在内存中转换为 OOXML，不写临时文件、不污染源目录；时间：2026-11-02；作者：lq。 */
  convertLegacy?(path: string): Promise<ConvertedLegacy | null>;
  /** 功能：在原目录安全重命名文件，目标存在时拒绝覆盖；时间：2026-11-07；作者：lq。 */
  renameFile?(path: string, name: string): Promise<string>;
  saveBytes(bytes: Uint8Array, target: SaveTarget, mime: string): Promise<SaveOutcome>;
  listFonts(): Promise<FontInfo[]>;
  /** 功能：强制重新扫描字体，供用户安装字体后即时生效；时间：2026-10-29；作者：lq。 */
  refreshFonts?(): Promise<FontInfo[]>;
  saveSnapshot(snapshot: DocumentSnapshot): Promise<void>;
  listSnapshots(): Promise<DocumentSnapshotMeta[]>;
  loadSnapshot(id: string): Promise<DocumentSnapshot | null>;
  deleteSnapshot(id: string): Promise<void>;
  revealInFolder(path: string): Promise<void>;
  readDroppedFile?(file: File): Promise<OpenedFile>;
}

declare global {
  interface Window {
    showOpenFilePicker?: (options?: {
      multiple?: boolean;
      types?: Array<{ description: string; accept: Record<string, string[]> }>;
      excludeAcceptAllOption?: boolean;
    }) => Promise<FileSystemFileHandle[]>;
    showSaveFilePicker?: (options?: {
      suggestedName?: string;
      types?: Array<{ description: string; accept: Record<string, string[]> }>;
    }) => Promise<FileSystemFileHandle>;
  }
}

export function isTauriRuntime(): boolean {
  if (typeof window === 'undefined') return false;
  return '__TAURI_INTERNALS__' in window || '__TAURI__' in window;
}

export function isMacOS(): boolean {
  if (typeof navigator === 'undefined') return false;
  const platform = navigator.platform || '';
  return /mac/i.test(platform) || /Macintosh|Mac OS X/i.test(navigator.userAgent || '');
}

export function isWindows(): boolean {
  if (typeof navigator === 'undefined') return false;
  const platform = navigator.platform || '';
  return /win/i.test(platform) || /Windows/i.test(navigator.userAgent || '');
}

/**
 * 功能：Windows 桌面端使用自绘标题栏，文档标签并入顶栏一行；macOS 保留 Overlay 红绿灯；
 * 时间：2026-11-01；作者：lq。
 */
export function usesMergedTitleBar(): boolean {
  return isTauriRuntime() && isWindows();
}

/** Native vibrancy is only meaningful in the macOS desktop shell. */
export function supportsWindowVibrancy(): boolean {
  return isTauriRuntime() && isMacOS();
}

export function toUint8Array(value: unknown): Uint8Array {
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  if (Array.isArray(value)) return new Uint8Array(value);
  if (typeof value === 'string') {
    const binary = atob(value);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  }
  throw new Error('无法读取文件数据');
}

export function fileTypeFilters(extensions: string[]): Array<{ name: string; extensions: string[] }> {
  const cleaned = extensions.map((ext) => ext.replace(/^\./, '').toLowerCase());
  const labels: Record<string, string> = {
    docx: 'Word 文档',
    xlsx: 'Excel 工作簿',
    xlsm: 'Excel 启用宏的工作簿',
    pptx: 'PowerPoint 演示文稿',
    csv: 'CSV 文本',
    pdf: 'PDF 文档',
    png: 'PNG 图片',
  };
  if (cleaned.length === 1) {
    return [{ name: labels[cleaned[0]] ?? 'Office 文件', extensions: cleaned }];
  }
  return [{ name: 'Office 文件', extensions: cleaned }];
}

export function acceptFilters(extensions: string[]): Record<string, string[]> {
  const map: Record<string, string[]> = {};
  for (const ext of extensions) {
    const clean = ext.replace(/^\./, '').toLowerCase();
    if (clean === 'docx') map['application/vnd.openxmlformats-officedocument.wordprocessingml.document'] = ['.docx'];
    else if (clean === 'xlsx' || clean === 'xlsm')
      map['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'] = [`.${clean}`];
    else if (clean === 'pptx')
      map['application/vnd.openxmlformats-officedocument.presentationml.presentation'] = ['.pptx'];
    else if (clean === 'csv') map['text/csv'] = ['.csv'];
    else if (clean === 'pdf') map['application/pdf'] = ['.pdf'];
    else map[`application/${clean}`] = [`.${clean}`];
  }
  return map;
}
