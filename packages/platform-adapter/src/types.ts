import type { DocumentSnapshot, DocumentSnapshotMeta, FontInfo } from '@mini-office/shared';

export interface OpenedFile {
  name: string;
  bytes: Uint8Array;
  path?: string;
  handle?: FileSystemFileHandle;
  lastModified?: number;
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
  /** 功能：将旧版 Office 文件转换为 OOXML 后再进入统一编辑流程；时间：2026-10-27；作者：lq。 */
  convertLegacy?(path: string): Promise<OpenedFile | null>;
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
