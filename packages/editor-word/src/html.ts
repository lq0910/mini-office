import type { DocImage, DocParagraph, DocRun, ParagraphStyle, RunStyle } from '@mini-office/document-model';
import { DOC_FONT_STACK, fontStackFor, isDarkInk, isLightShade, shadeColor } from '@mini-office/design-system';
import { escapeHtml } from '@mini-office/shared';
import type { TocEntry } from './toc';
// 功能：段落行高使用各 run 实际字体度量，而非文档默认字体；时间：2026-11-05；作者：lq。
import { fallbackFontFamily, paragraphLineMetrics } from './metrics';

// 功能：颜色换算逻辑已抽到 design-system，与 Excel 共用同一套判定；这里原样转出，保持既有导入路径可用；时间：2026-10-31；作者：lq。
export { isDarkInk, isLightShade, shadeColor };


export interface RunRenderContext {
  pageNumber?: number;
  totalPages?: number;
  /** 功能：SECTIONPAGES 使用当前报告节的总页数；时间：2026-10-24；作者：lq。 */
  sectionPages?: number;
  /** 功能：目录段落的条目（含页码），由编辑器按当前分页实时计算；时间：2026-10-30；作者：lq。 */
  tocEntries?: TocEntry[];
  /** 功能：编辑态显示 Word 风格的换行符与制表符格式标记；时间：2026-10-31；作者：lq。 */
  showFormattingMarks?: boolean;
  /** 功能：当前段落的左缩进，用于把 positionH（相对栏）的偏移换算成相对段落盒的位移；时间：2026-10-31；作者：lq。 */
  indentLeftPx?: number;
}

/**
 * 功能：判断是否为「装饰性细条」——下划线、分隔线这类又扁又长的浮动图；时间：2026-10-31；作者：lq。
 *
 * 这类图片在 Word 里虽然声明了四周型环绕，但高度只有几个像素，
 * 让它们参与行内流只会把后续段落的行盒挤短（正文莫名换行），语义上也毫无绕排意义。
 */
function isDecorativeStrip(image: DocImage): boolean {
  return image.heightPx <= 24 && image.widthPx / Math.max(1, image.heightPx) >= 8;
}

/**
 * 功能：判断浮动图是否应完全脱离行内流（绝对定位）；时间：2026-10-31；作者：lq。
 *
 * 依据 Word 语义分两类：
 * 1. `wp:wrapNone`（浮于文字上方/下方，常见于公章、签名图）本身就不参与绕排；
 * 2. 装饰性细条（下划线）虽声明四周型，但高度只有几像素，参与流布局弊大于利。
 * 其余四周型/上下型图片仍走 float/block，保留文字绕排能力。
 */
function isOutOfFlowImage(image: DocImage): boolean {
  const anchor = image.anchor;
  if (!anchor) return false;
  if ((image.wrap ?? 'inline') === 'inline') return false;
  if (anchor.wrapType === 'none') return true;
  return isDecorativeStrip(image);
}

/** 功能：深色文字打上 data-ink 标记，深色纸张提亮、打印与导出时还原；时间：2026-10-31；作者：lq。 */
function inkAttr(style: RunStyle): string {
  return style.color && isDarkInk(style.color) ? ' data-ink="1"' : '';
}

/** 功能：把 wp:anchor 的 positionH/positionV 换算成绝对定位样式；时间：2026-10-31；作者：lq。 */
function absoluteAnchorCss(image: DocImage, context: RunRenderContext): string {
  const anchor = image.anchor ?? {};
  const indent = context.indentLeftPx ?? 0;
  const offset = anchor.hOffsetPx ?? 0;
  const horizontal =
    anchor.hAlign === 'right' || anchor.hAlign === 'outside'
      ? 'right:0;'
      : anchor.hAlign === 'center'
        ? 'left:50%;transform:translateX(-50%);'
        : `left:${Math.max(0, offset - (anchor.hRelative === 'column' ? indent : 0))}px;`;
  const vertical = anchor.vAlign === 'bottom'
    ? 'bottom:0;'
    : anchor.vAlign === 'center'
      ? 'top:50%;transform:translateY(-50%);'
      : `top:${Math.max(0, anchor.vOffsetPx ?? 0)}px;`;
  // 功能：衬于文字下方的图片（behindDoc）放到正文之下，浮于文字上方的放到正文之上；时间：2026-10-31；作者：lq。
  const layer = anchor.behindDoc ? 0 : 3;
  return `position:absolute;${horizontal}${vertical}z-index:${layer};`;
}

function imageHtml(run: DocRun, context: RunRenderContext): string {
  const image = run.image;
  if (!image) return '';
  const alt = escapeHtml(image.alt ?? image.name ?? '');
  const imageId = escapeHtml(image.id ?? image.relId ?? '');
  const wrap = image.wrap ?? 'inline';
  const anchor = image.anchor;
  // 功能：环绕距离直接取 wp:anchor 的 distT/distB/distL/distR，不再硬编码 8px 下边距；时间：2026-10-31；作者：lq。
  // 旧实现给浮动图片加了 8px 下边距，使 8px 高的下划线图片实际占据 16px 行盒，
  // 正好压到下一段首行，把「乙方：xxx」这类整行文字挤成两行。
  const distTop = anchor?.distTopPx ?? 0;
  const distBottom = anchor?.distBottomPx ?? 0;
  const distLeft = anchor?.distLeftPx ?? (anchor ? 0 : 12);
  const distRight = anchor?.distRightPx ?? (anchor ? 0 : 12);
  // 功能：positionH 的偏移量换算成图片相对栏左边缘的水平位移；时间：2026-10-31；作者：lq。
  const offsetLeft = Math.max(0, anchor?.hOffsetPx ?? 0);
  const outOfFlow = isOutOfFlowImage(image);
  const layout = outOfFlow
    ? absoluteAnchorCss(image, context)
    : wrap === 'square-left'
      ? `position:relative;float:left;margin:${distTop}px ${distRight}px ${distBottom}px ${offsetLeft}px;`
      : wrap === 'square-right'
        ? `position:relative;float:right;margin:${distTop}px ${offsetLeft}px ${distBottom}px ${distLeft}px;`
        : wrap === 'top-bottom'
          ? `position:relative;display:block;margin:${Math.max(distTop, 8)}px auto ${Math.max(distBottom, 8)}px;`
          : 'position:relative;display:inline-block;vertical-align:baseline;';
  // 功能：内嵌图片效果外扩参与行盒占位，浮动图继续由锚定布局负责；时间：2026-11-04；作者：lq。
  const effect = !anchor && wrap === 'inline' ? image.effectExtent : undefined;
  const effectMargin = effect ? `margin:${effect.top}px ${effect.right}px ${effect.bottom}px ${effect.left}px;` : '';
  // 功能：自动倍数图片行的额外间距随每张内嵌图参与行盒，换行后每行均生效，并与原有效果下外扩相加；时间：2026-11-05；作者：lq。
  const autoLineMargin = !anchor && wrap === 'inline'
    ? `margin-bottom:calc(${effect?.bottom ?? 0}px + var(--mo-inline-image-line-extra, 0px));` : '';
  // 功能：图片帧显式锁定为图片尺寸，避免帧被内部块级元素或样式丢失撑成大块；时间：2026-10-31；作者：lq。
  const frameStyle = `line-height:0;box-sizing:content-box;width:${image.widthPx}px;height:${image.heightPx}px;${layout}${effectMargin}${autoLineMargin}`;
  // 功能：按文档声明的显示尺寸拉伸图片，避免非等比图片被 contain 留白，导致图案与尺寸手柄不一致；时间：2026-11-04；作者：lq。
  const imageStyle = `display:block;width:${image.widthPx}px;height:${image.heightPx}px;max-width:100%;object-fit:fill;`;
  // 功能：为 Word 图片输出可直接拖动的八方向尺寸手柄；时间：2026-10-24；作者：lq。
  const handles = [
    ['nw', 'nwse-resize', 'left:-5px;top:-5px;'],
    ['n', 'ns-resize', 'left:calc(50% - 4px);top:-5px;'],
    ['ne', 'nesw-resize', 'right:-5px;top:-5px;'],
    ['e', 'ew-resize', 'right:-5px;top:calc(50% - 4px);'],
    ['se', 'nwse-resize', 'right:-5px;bottom:-5px;'],
    ['s', 'ns-resize', 'left:calc(50% - 4px);bottom:-5px;'],
    ['sw', 'nesw-resize', 'left:-5px;bottom:-5px;'],
    ['w', 'ew-resize', 'left:-5px;top:calc(50% - 4px);'],
  ].map(([name, cursor, position]) => `<span data-image-handle="${name}" contenteditable="false" aria-hidden="true" style="position:absolute;z-index:3;width:8px;height:8px;box-sizing:border-box;border:1px solid #3478f6;background:#fff;border-radius:2px;opacity:0;pointer-events:none;cursor:${cursor};${position}"></span>`).join('');
  return `<span data-image-frame data-image-id="${imageId}" contenteditable="false" style="${frameStyle}"><img src="${escapeHtml(image.src)}" alt="${alt}" data-image-id="${imageId}" contenteditable="false" draggable="false" style="${imageStyle}">${handles}</span>`;
}

function runStyleCss(style: RunStyle): string {
  const css: string[] = [];
  if (style.fontFamily) {
    // 功能：渲染时遵循文档字体表的备用名，保持主字体语义；时间：2026-11-05；作者：lq。
    css.push(`font-family:${fontStackFor(style.fontFamily, fallbackFontFamily(style)).replace(/"/g, "'")}`);
  }
  if (style.fontSizePt) css.push(`font-size:${style.fontSizePt}pt`);
  if (style.bold) css.push('font-weight:700');
  if (style.italic) css.push('font-style:italic');
  const decorations: string[] = [];
  if (style.underline) decorations.push('underline');
  if (style.strike) decorations.push('line-through');
  if (decorations.length) css.push(`text-decoration:${decorations.join(' ')}`);
  // 功能：保留文档原始颜色；深色纸张下由样式表按 data-ink 标记统一提亮，打印与导出自动还原；时间：2026-10-31；作者：lq。
  if (style.color) css.push(`color:${style.color}`);
  // 功能：文字底纹同样走纸张换算——纯白底纹在深色纸上换成深色底，否则会出现白块压住文字；时间：2026-10-31；作者：lq。
  if (style.highlight) css.push(`background-color:${shadeColor(style.highlight)}`);
  if (style.letterSpacingPx) css.push(`letter-spacing:${style.letterSpacingPx}px`);
  if (style.allCaps) css.push('text-transform:uppercase');
  if (style.smallCaps) css.push('font-variant:small-caps');
  return css.join(';');
}

/** 功能：显式 Tab 与文字内原始 Tab 共用占位结构，确保前导线、编辑和制表位测量一致；时间：2026-11-03；作者：lq。 */
function tabPlaceholderHtml(context: RunRenderContext): string {
  return `<span data-tab="1" style="display:inline-block;width:0;height:0;line-height:0;vertical-align:baseline;user-select:none;pointer-events:none;">${context.showFormattingMarks ? '<span aria-hidden="true" style="color:#4b9cff;font-size:.72em;line-height:0;user-select:none;">→</span>' : ''}</span>`;
}

export function renderRunsHtml(runs: DocRun[], context: RunRenderContext = {}): string {
  const parts: string[] = [];
  for (const run of runs) {
    if (run.image) {
      parts.push(imageHtml(run, context));
      continue;
    }
    if (run.lineBreak) {
      parts.push(context.showFormattingMarks ? '<span data-line-break="1" aria-hidden="true" style="display:inline-block;width:0;height:0;vertical-align:baseline;color:#4b9cff;font-size:.72em;line-height:0;user-select:none;pointer-events:none;">↵</span><br>' : '<br>');
      continue;
    }
    if (run.tab) {
      // 功能：制表符渲染为占位块，宽度由布局后置处理按制表位实测设置，不再固定两个全角空格；时间：2026-10-30；作者：lq。
      // 功能：制表符占位块高度归零，避免 inline-block 的 1em 高度把固定行距的行盒撑高；时间：2026-10-31；作者：lq。
      parts.push(tabPlaceholderHtml(context));
      continue;
    }
    // 功能：字段显示为不可拆开的占位块，保留原字符样式；时间：2026-10-24；作者：lq。
    if (run.mergeField) {
      parts.push(`<span data-merge-field="${escapeHtml(run.mergeField)}"${inkAttr(run.style)} contenteditable="false" style="${runStyleCss(run.style)}"><span style="background:#e8f1ff;border-radius:3px;color:#245694;">«${escapeHtml(run.mergeField)}»</span></span>`);
      continue;
    }
    let text = run.text ?? '';
    if (run.field === 'PAGE' && context.pageNumber !== undefined) {
      text = String(context.pageNumber);
    } else if (run.field === 'SECTIONPAGES' && (context.sectionPages ?? context.totalPages) !== undefined) {
      text = String(context.sectionPages ?? context.totalPages);
    } else if (run.field === 'NUMPAGES' && context.totalPages !== undefined) {
      text = String(context.totalPages);
    }
    if (!text) continue;
    // 功能：合成/导入文档的 w:t 可含原始 Tab，显示时接入相同布局通道，保留模型与源 XML；时间：2026-11-03；作者：lq。
    const escaped = escapeHtml(text).replace(/\n/g, '<br>').replace(/\t/g, () => tabPlaceholderHtml(context));
    const css = runStyleCss(run.style);
    let html = css ? `<span${inkAttr(run.style)} style="${css}">${escaped}</span>` : escaped;
    if (run.style.verticalAlign === 'super') html = `<sup>${html}</sup>`;
    else if (run.style.verticalAlign === 'sub') html = `<sub>${html}</sub>`;
    parts.push(html);
  }
  return parts.join('');
}

/**
 * 功能：渲染目录条目——标题 + 点线前导 + 右对齐页码；时间：2026-10-30；作者：lq。
 *
 * 用 flex 行而不是制表符排版，是为了让长标题自动省略、页码始终贴右边界，
 * 并且在缩放、打印与 PDF 导出时表现一致（都是普通 DOM，不需要额外测量）。
 */
export function renderTocHtml(entries: TocEntry[]): string {
  if (entries.length === 0) {
    return '<span style="color:#98a2b3;">（暂无标题，目录为空）</span>';
  }
  return entries
    .map(
      (entry) =>
        `<div data-toc-entry="${escapeHtml(entry.paragraphId)}" style="display:flex;align-items:baseline;gap:6px;">` +
        `<span style="flex:0 1 auto;min-width:0;padding-left:${(entry.level - 1) * 2}em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(entry.text)}</span>` +
        `<span style="flex:1 1 auto;min-width:16px;border-bottom:1px dotted currentColor;opacity:0.5;transform:translateY(-2px);"></span>` +
        `<span style="flex:0 0 auto;font-variant-numeric:tabular-nums;">${entry.page}</span>` +
        `</div>`,
    )
    .join('');
}

export function renderParagraphHtml(
  paragraph: DocParagraph,
  context: RunRenderContext = {},
): string {
  // 功能：目录段落不渲染缓存文字，直接按当前标题与分页生成条目；时间：2026-10-30；作者：lq。
  if (paragraph.toc) return renderTocHtml(context.tocEntries ?? []);
  let marker = '';
  if (paragraph.list) {
    marker = `<span data-list-marker="1" contenteditable="false" style="display:inline-block;min-width:1.5em;white-space:nowrap;">${escapeHtml(
      paragraph.list.marker,
    )}</span>`;
  }
  const content = renderRunsHtml(paragraph.runs, { ...context, indentLeftPx: paragraph.style.indentLeftPx ?? 0 });
  return `${marker}${content || '<br>'}`;
}

export function paragraphWrapperStyle(
  paragraph: DocParagraph,
  options: { dropSpaceBefore?: boolean; spaceBeforePx?: number; spaceAfterPx?: number } = {},
): React.CSSProperties {
  const style = paragraph.style;
  const borderCss = (border: NonNullable<ParagraphStyle['borders']>['top']) => border && border.style !== 'none' ? `${Math.max(0.5, border.widthPx)}px ${border.style} ${border.color}` : undefined;
  return {
    // 功能：段前距在页首按 Word 规则丢弃，避免每个新页顶部多出一段空白；时间：2026-11-02；作者：lq。
    paddingTop: options.dropSpaceBefore ? 0 : options.spaceBeforePx ?? style.spaceBeforePx ?? 0,
    paddingBottom: options.spaceAfterPx ?? style.spaceAfterPx ?? 0,
    marginLeft: style.indentLeftPx ?? 0,
    marginRight: style.indentRightPx ?? 0,
    textAlign: style.alignment ?? 'left',
    backgroundColor: shadeColor(style.backgroundColor),
    borderTop: borderCss(style.borders?.top),
    borderRight: borderCss(style.borders?.right),
    borderBottom: borderCss(style.borders?.bottom),
    borderLeft: borderCss(style.borders?.left),
  };
}

export function paragraphTextStyle(
  paragraph: DocParagraph,
  defaultFont: string,
  defaultFontSizePt: number,
  gridLinePitchPx?: number,
): React.CSSProperties {
  const style: ParagraphStyle = paragraph.style;
  // 功能：snapToGrid=false 的段落不使用 docGrid 的 linePitch，回退到字体度量；时间：2026-11-03；作者：lq。
  const effectiveGridPitch = style.snapToGrid === false ? undefined : gridLinePitchPx;
  // 功能：继承字体与行盒使用同一实际 run，防止中文字体跨两格时被默认 Calibri 压成一格；时间：2026-11-05；作者：lq。
  const metrics = paragraphLineMetrics(paragraph.runs, style, defaultFont, defaultFontSizePt, effectiveGridPitch);
  const fontSizePt = metrics.sizePt;
  const fontFamily = fontStackFor(metrics.family, metrics.alternateFamily) || DOC_FONT_STACK;
  const computedLineHeight = metrics.heightPx;
  // 功能：无网格的纯内嵌图片自动行距由图片撑高，倍数只增加字体行高的额外间距，避免 CSS 下伸部重复占位；时间：2026-11-05；作者：lq。
  const imageOnlyAutoLine = !paragraph.list && !paragraph.toc && paragraph.runs.length > 0
    && paragraph.runs.every((run) => run.image && (run.image.wrap ?? 'inline') === 'inline'
      && !run.image.anchor && !run.text && !run.tab && !run.lineBreak && !run.field && !run.mergeField)
    && style.lineHeightExactPx === undefined && style.lineHeightAtLeastPx === undefined
    && !effectiveGridPitch;
  // 功能：保留自动倍数相对单倍增加的间距；固定值、最小值、网格和文字混排仍沿用完整文字行盒；时间：2026-11-05；作者：lq。
  const imageLineExtra = imageOnlyAutoLine
    ? Math.max(0, computedLineHeight - metrics.naturalHeightPx) : 0;
  return {
    fontFamily,
    fontSize: `${fontSizePt}pt`,
    // 功能：固定值/最小值/倍数三种行距都落到同一个像素高度，保证测量与渲染一致；时间：2026-11-02；作者：lq。
    lineHeight: imageOnlyAutoLine ? '0px' : `${computedLineHeight}px`,
    // 功能：使用继承变量把间距传入图片帧；普通段落显式归零，避免嵌套段落继承外层图片行的间距；时间：2026-11-05；作者：lq。
    ...{ '--mo-inline-image-line-extra': `${imageLineExtra}px` },
    // 功能：为浮动图片（绝对定位）提供定位上下文；时间：2026-10-31；作者：lq。
    position: 'relative',
    // 功能：只含浮动图的段落没有行内内容会塌成 0 高，用一行最小高度保住 Word 的行盒；时间：2026-10-31；作者：lq。
    minHeight: `${computedLineHeight}px`,
    textIndent: style.firstLineIndentPx ? `${style.firstLineIndentPx}px` : undefined,
    whiteSpace: 'pre-wrap',
    overflowWrap: 'break-word',
    wordBreak: 'break-word',
  };
}

export function paragraphPlainText(paragraph: DocParagraph, context: RunRenderContext = {}): string {
  return paragraph.runs
    .map((run) => {
      if (run.image) return '[图片]';
      if (run.tab) return '\t';
      if (run.lineBreak) return '\n';
      if (run.field === 'PAGE') return String(context.pageNumber ?? run.text ?? '');
      if (run.field === 'SECTIONPAGES') return String(context.sectionPages ?? context.totalPages ?? run.text ?? '');
      if (run.field === 'NUMPAGES') return String(context.totalPages ?? run.text ?? '');
      return run.text ?? '';
    })
    .join('');
}
