import {
  ErrorCodes,
  OfficeError,
  baseName,
  type DocumentSnapshot,
  type DocumentSnapshotMeta,
  type FontInfo,
} from '@mini-office/shared';
import { fileTypeFilters, toUint8Array, type ConvertedLegacy, type OpenedFile, type PlatformAdapter, type SaveOutcome, type SaveTarget } from './types';
import { getSnapshot, listSnapshots, putSnapshot, removeSnapshot } from './snapshots';

async function tauriInvoke<T>(command: string, args?: Record<string, unknown>): Promise<T> {
  const { invoke } = await import('@tauri-apps/api/core');
  return invoke<T>(command, args);
}

async function pickOpenPath(extensions: string[]): Promise<string | null> {
  const { open } = await import('@tauri-apps/plugin-dialog');
  const selected = await open({
    multiple: false,
    directory: false,
    filters: fileTypeFilters(extensions),
  });
  return typeof selected === 'string' ? selected : null;
}

async function pickSavePath(suggestedName: string, extensions: string[]): Promise<string | null> {
  const { save } = await import('@tauri-apps/plugin-dialog');
  const selected = await save({
    defaultPath: suggestedName,
    filters: fileTypeFilters(extensions),
  });
  return typeof selected === 'string' ? selected : null;
}

function extensionOfName(name: string): string[] {
  const ext = name.split('.').pop();
  return ext ? [ext] : [];
}

export class TauriPlatformAdapter implements PlatformAdapter {
  readonly kind = 'tauri' as const;

  async openPath(path: string): Promise<OpenedFile | null> {
    try {
      const result = await tauriInvoke<ArrayBuffer | number[]>('read_file', { path });
      return {
        name: baseName(path) || '未命名文档',
        bytes: toUint8Array(result),
        path,
      };
    } catch (error) {
      throw new OfficeError(
        ErrorCodes.PARSE_FAILED,
        `无法读取文件：${baseName(path)}`,
        error instanceof Error ? error.message : String(error),
      );
    }
  }

  async openFile(extensions: string[]): Promise<OpenedFile | null> {
    const path = await pickOpenPath(extensions);
    if (!path) return null;
    return this.openPath(path);
  }

  /** 功能：调用 Rust 兼容桥在内存中转换 DOC/XLS/PPT，源目录不新增文件；时间：2026-11-02；作者：lq。 */
  async convertLegacy(path: string): Promise<ConvertedLegacy | null> {
    try {
      // 功能：DOC 改走保留字体、样式、表格和节信息的浏览器转换器；XLS/PPT 沿用现有 Rust 通道；时间：2026-11-03；作者：lq。
      if (/\.doc$/i.test(path)) {
        const opened = await this.openPath(path);
        if (!opened) return null;
        const { convertLegacyDoc } = await import('./legacy-doc');
        return convertLegacyDoc(opened.bytes, opened.name);
      }
      const converted = await tauriInvoke<{ name: string; sourceName: string; bytes: ArrayBuffer | number[] }>(
        'convert_legacy_document',
        { input: path },
      );
      return {
        name: converted.name,
        sourceName: converted.sourceName,
        bytes: toUint8Array(converted.bytes),
      };
    } catch (error) {
      throw new OfficeError(ErrorCodes.PARSE_FAILED, `无法转换旧版文件：${baseName(path)}`, error instanceof Error ? error.message : String(error));
    }
  }

  async saveBytes(bytes: Uint8Array, target: SaveTarget, mime: string): Promise<SaveOutcome> {
    let path = target.path;
    if (!path || target.forcePicker) {
      const picked = await pickSavePath(target.suggestedName, extensionOfName(target.suggestedName));
      if (!picked) throw new OfficeError(ErrorCodes.CANCELLED, '保存已取消');
      path = picked;
    }
    try {
      const result = await tauriInvoke<{ path: string; backupPath?: string | null }>('write_file', {
        path,
        bytes: Array.from(bytes),
      });
      return {
        displayName: baseName(result.path || path),
        path: result.path || path,
        location: target.forcePicker ? 'disk' : 'in-place',
      };
    } catch (error) {
      throw new OfficeError(
        ErrorCodes.SAVE_FAILED,
        `保存失败：${baseName(path)}`,
        error instanceof Error ? error.message : String(error),
      );
    }
  }

  async listFonts(): Promise<FontInfo[]> {
    try {
      const fonts = await tauriInvoke<Array<{ family: string; style?: string; path: string; aliases?: string[] }>>('list_fonts');
      return fonts.map((font) => ({
        family: font.family,
        style: font.style,
        source: font.path,
        hasChinese:
          /[\u3400-\u9fff]/.test(font.family) ||
          /CJK|Hei|Song|Kai|Fang|YaHei|PingFang/.test(font.family) ||
          (font.aliases ?? []).some((alias) => /[\u3400-\u9fff]/.test(alias)),
        aliases: font.aliases ?? [],
      }));
    } catch {
      return [];
    }
  }

  /** 功能：强制重新扫描系统字体，供用户安装字体后即时生效；时间：2026-10-29；作者：lq。 */
  async refreshFonts(): Promise<FontInfo[]> {
    try {
      const fonts = await tauriInvoke<Array<{ family: string; style?: string; path: string; aliases?: string[] }>>('refresh_fonts');
      return fonts.map((font) => ({
        family: font.family,
        style: font.style,
        source: font.path,
        hasChinese:
          /[\u3400-\u9fff]/.test(font.family) ||
          /CJK|Hei|Song|Kai|Fang|YaHei|PingFang/.test(font.family) ||
          (font.aliases ?? []).some((alias) => /[\u3400-\u9fff]/.test(alias)),
        aliases: font.aliases ?? [],
      }));
    } catch {
      return [];
    }
  }

  async saveSnapshot(snapshot: DocumentSnapshot): Promise<void> {
    await putSnapshot(snapshot);
  }

  async listSnapshots(): Promise<DocumentSnapshotMeta[]> {
    return listSnapshots();
  }

  async loadSnapshot(id: string): Promise<DocumentSnapshot | null> {
    return getSnapshot(id);
  }

  async deleteSnapshot(id: string): Promise<void> {
    await removeSnapshot(id);
  }

  async revealInFolder(path: string): Promise<void> {
    try {
      const { revealItemInDir } = await import('@tauri-apps/plugin-opener');
      await revealItemInDir(path);
    } catch {
      // ignore
    }
  }

  async readDroppedFile(file: File): Promise<OpenedFile> {
    const bytes = new Uint8Array(await file.arrayBuffer());
    if (/\.(doc|xls|ppt)$/i.test(file.name)) {
      try {
        // 功能：拖拽 DOC 复用路径打开的同一转换器，避免两个入口保真度不同；时间：2026-11-03；作者：lq。
        if (/\.doc$/i.test(file.name)) {
          const { convertLegacyDoc } = await import('./legacy-doc');
          return convertLegacyDoc(bytes, file.name);
        }
        // 功能：拖拽旧版 Office 文件时直接走内存转换，不创建用户可见临时文件；时间：2026-10-28；作者：lq。
        const converted = await tauriInvoke<{ name: string; sourceName: string; bytes: ArrayBuffer | number[] }>('convert_legacy_bytes', { name: file.name, bytes: Array.from(bytes) });
        return { name: converted.name, sourceName: converted.sourceName, bytes: toUint8Array(converted.bytes) };
      } catch (error) {
        throw new OfficeError(ErrorCodes.PARSE_FAILED, `无法转换旧版文件：${file.name}`, error instanceof Error ? error.message : String(error));
      }
    }
    return { name: file.name, bytes };
  }
}
