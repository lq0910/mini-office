/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */
/**
 * Reusable Range: Blink fixes up every live Range on each DOM removal, so
 * per-word probe Ranges make later ProseMirror re-renders crawl. Do not nest.
 */
export function rangeSlot(): () => Range {
  let range: Range | undefined
  return () => (range ??= document.createRange())
}
