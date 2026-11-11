/** 功能：使用固定版本 GenOffice 转换旧 DOC，全部处理在内存完成且不写入源目录；时间：2026-11-03；作者：lq。 */
import { baseName } from '@mini-office/shared';
import type { ConvertedLegacy } from './types';

/** 功能：按需加载 DOC 转换器，保留源文件名并返回损失警告，避免无提示降级；时间：2026-11-03；作者：lq。 */
export async function convertLegacyDoc(bytes: Uint8Array, name: string): Promise<ConvertedLegacy> {
  const sourceName = baseName(name);
  // 功能：按实际容器区分改名的 OOXML 与二进制 DOC，已是 Word ZIP 时不加载转换器或制造损失警告；时间：2026-11-11；作者：lq。
  if (bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 3 && bytes[3] === 4) {
    const { loadOoxmlPackage, detectKind } = await import('@mini-office/ooxml');
    const pkg = await loadOoxmlPackage(bytes);
    try {
      if ((await detectKind(pkg)).kind !== 'docx') throw new Error('该 .doc 文件的实际内容不是 Word 文档');
      return { name: sourceName.replace(/\.(doc|dot)$/i, '.docx'), sourceName, bytes, conversionPerformed: false, conversionWarnings: [] };
    } finally { pkg.dispose(); }
  }
  const { convertLegacyDocToDocx } = await import('./vendor/genoffice/doc-converter.js');
  const result = await convertLegacyDocToDocx(bytes);
  return { name: sourceName.replace(/\.(doc|dot)$/i, '.docx'), sourceName, bytes: result.docx, conversionPerformed: true, conversionWarnings: result.warnings };
}
