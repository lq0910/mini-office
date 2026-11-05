/** 功能：使用固定版本 GenOffice 转换旧 DOC，全部处理在内存完成且不写入源目录；时间：2026-11-03；作者：lq。 */
import { baseName } from '@mini-office/shared';
import type { ConvertedLegacy } from './types';

/** 功能：按需加载 DOC 转换器，保留源文件名并返回损失警告，避免无提示降级；时间：2026-11-03；作者：lq。 */
export async function convertLegacyDoc(bytes: Uint8Array, name: string): Promise<ConvertedLegacy> {
  const { convertLegacyDocToDocx } = await import('./vendor/genoffice/doc-converter.js');
  const result = await convertLegacyDocToDocx(bytes);
  const sourceName = baseName(name);
  return { name: sourceName.replace(/\.(doc|dot)$/i, '.docx'), sourceName, bytes: result.docx, conversionWarnings: result.warnings };
}
