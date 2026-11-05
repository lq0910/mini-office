import type { DocBlock, DocParagraph } from '@mini-office/document-model';

export interface EffectiveParagraphSpacing {
  beforePx: number;
  afterPx: number;
}

/** 功能：计算上下文段落的有效间距，邻接同样式段落且启用 contextualSpacing 时消除两侧空隙；时间：2026-11-05；作者：lq。 */
export function effectiveParagraphSpacing(blocks: readonly DocBlock[], index: number): EffectiveParagraphSpacing | undefined {
  const paragraph = blocks[index];
  if (!paragraph || paragraph.type !== 'paragraph') return undefined;
  const previous = blocks[index - 1];
  const next = blocks[index + 1];
  const sameStyle = (other: DocBlock | undefined): other is DocParagraph =>
    other?.type === 'paragraph' && other.styleId === paragraph.styleId;
  const beforeSharesStyle = sameStyle(previous);
  const afterSharesStyle = sameStyle(next);
  // 功能：显式 contextualSpacing 仅控制本段；false 覆盖样式继承，不受相邻段设置反向覆盖；时间：2026-11-05；作者：lq。
  const beforePx = beforeSharesStyle && paragraph.style.contextualSpacing
    ? 0 : paragraph.style.spaceBeforePx ?? 0;
  const afterPx = afterSharesStyle && paragraph.style.contextualSpacing
    ? 0 : paragraph.style.spaceAfterPx ?? 0;
  return { beforePx, afterPx };
}
