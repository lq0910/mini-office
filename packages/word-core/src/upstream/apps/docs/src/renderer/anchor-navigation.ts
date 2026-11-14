/** 功能：同步 GenOffice 目录导航；时间：2026-11-14；作者：lq；上游版本及 Apache-2.0 许可见 UPSTREAM.md。 */
import type { ParsedDocFull } from '../../../../packages/docx-engine/src/index';
/** 功能：保留上游函数体，仅注入当前文档和真实编辑根，避免跨标签或快照误命中；时间：2026-11-14；作者：lq。 */
export function jumpToWordAnchor(root: HTMLElement, parsed: ParsedDocFull, anchor: string, tocTitle = ''): Element | null {
  const doc = { parsed };
  const document = root;

      let target: Element | null = null
      if (anchor && doc) {
        const block = doc.parsed.blocks.find((b) => b.originalXml?.includes(`w:name="${anchor}"`))
        if (block && block.docxIndex !== null) {
          target = document.querySelector(`.ProseMirror [data-idx="${block.docxIndex}"]`)
        }
      }
      const title = tocTitle.replace(/\s+/g, '')
      if (!target && title) {
        target =
          [
            ...document.querySelectorAll(
              '.ProseMirror h1, .ProseMirror h2, .ProseMirror h3, .ProseMirror h4, .ProseMirror h5, .ProseMirror h6',
            ),
          ].find((h) => (h.textContent ?? '').replace(/\s+/g, '') === title) ?? null
      }
      target?.scrollIntoView({ behavior: 'smooth', block: 'start' })

  return target;
}
