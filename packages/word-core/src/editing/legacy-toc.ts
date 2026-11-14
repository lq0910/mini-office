/** 功能：为旧版 DOC 转换快照补充 GenOffice 目录显示模型，不重写来源 XML、原字节或正文；时间：2026-11-14；作者：lq。 */
import JSZip from 'jszip';
import type { ParsedDocFull } from '../upstream/packages/docx-engine/src/index';
import { fieldDisplayOf } from '../upstream/packages/docx-engine/src/parse-fields';
import { resolveCtxOf } from '../upstream/packages/docx-engine/src/resolve-props';

/** 功能：只识别连续的 _Toc 内链、右对齐点引导制表位和页码组合，避免把普通超链接正文当目录；时间：2026-11-14；作者：lq。 */
export async function adoptLegacyTocDisplay(parsed: ParsedDocFull): Promise<void> {
  const candidates = parsed.blocks.map((block, index) => ({ block, index })).filter(({ block }) => {
    const xml = block.originalXml ?? '';
    return (block.type === 'paragraph' || (block.fieldDisplay?.kind === 'tocLine' && block.fieldDisplay.runStyleId === 'Hyperlink'))
      && /<w:hyperlink\b[^>]*w:anchor="_Toc[^"\s]+"/.test(xml)
      && /<w:tabs>[\s\S]*?<w:tab\b[^>]*w:val="right"[^>]*w:leader="dot"/.test(xml)
      && !/<w:(?:drawing|pict|sdt|sectPr)\b/.test(xml);
  });
  if (!candidates.length) return;
  const zip = await JSZip.loadAsync(parsed.internal.originalBytes);
  if (!/<Application>GenOffice<\/Application>/.test(await zip.file('docProps/app.xml')?.async('string') ?? '')) return;
  const resolver = resolveCtxOf(parsed);
  const indices = new Set(candidates.map(item => item.index));
  for (const { block, index } of candidates) {
    if (block.fieldDisplay?.kind !== 'tocLine' && !indices.has(index - 1) && !indices.has(index + 1)) continue;
    const xml = block.originalXml!;
    // 功能：只向同源字段读取器提供临时 TOC 样式提示，实际保存仍持有原始 XML，旧缓存中的显式下划线也不删除；时间：2026-11-14；作者：lq。
    const hint = /<w:pStyle\b/.test(xml) ? xml.replace(/<w:pStyle\b[^>]*\/>/, '<w:pStyle w:val="TOC1"/>')
      : xml.includes('<w:pPr>') ? xml.replace('<w:pPr>', '<w:pPr><w:pStyle w:val="TOC1"/>')
        : xml.replace(/(<w:p(?:\s[^>]*)?>)/, '$1<w:pPr><w:pStyle w:val="TOC1"/></w:pPr>');
    const field = block.fieldDisplay?.kind === 'tocLine' ? { ...block.fieldDisplay } : fieldDisplayOf(hint, parsed.styles, resolver);
    if (field?.kind !== 'tocLine' || !field.left || !/^(?:\d+|[ivxlcdm]+)$/i.test(field.right ?? '')) continue;
    // 功能：目录链接沿用同源 plain 显示语义，旧转换器附加的 Hyperlink 字符样式不再污染目录展示；时间：2026-11-14；作者：lq。
    if (field.runStyleId === 'Hyperlink') delete field.runStyleId;
    block.type = 'passthrough'; block.label = 'TOC'; block.fieldDisplay = field;
  }
}
