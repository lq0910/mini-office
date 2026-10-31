import type { DocParagraph, DocRun, ParagraphStyle, RunStyle, WordTabStop } from '@mini-office/document-model';
import { DOC_FONT_STACK, fontStackFor } from '@mini-office/design-system';
import { escapeHtml } from '@mini-office/shared';
import { measureNaturalLineHeightPx } from './font-metrics';

export interface RunRenderContext {
  pageNumber?: number;
  totalPages?: number;
  /** 功能：SECTIONPAGES 使用当前报告节的总页数；时间：2026-10-24；作者：lq。 */
  sectionPages?: number;
  /** 功能：段落制表位，供 tab 运行按左右/居中对齐跳转；时间：2026-10-30。 */
  tabStops?: WordTabStop[];
  /** 功能：默认制表位间距（px），无显式制表位时按它的整数倍跳转；时间：2026-10-30。 */
  defaultTabStopPx?: number;
  /**
   * 功能：本段所在栏/内容区左缘的页面横坐标（含页边距与栏偏移）——VML 浮层的 page 参照系据此换算成
   * 相对所属段落的 left，从而不依赖挂载后的坐标测量；时间：2026-10-30。
   */
  vmlColumnLeftPx?: number;
  /** 功能：段落左缩进（px），与栏左缘一起把页面横坐标换算成段落内偏移；时间：2026-10-30。 */
  vmlIndentLeftPx?: number;
}

/** 功能：Word 倍数行距基于字体自然行高（中文字体约 1.2×字号），CSS line-height 单位数是基于字号，需乘此系数对齐 Word 表现；时间：2026-10-30。 */
const FONT_LINE_HEIGHT_RATIO = 1.2;

function runStyleCss(style: RunStyle): string {
  const css: string[] = [];
  if (style.fontFamily) {
    // 功能：中西文分开组栈——西文用主字体，中文插 eastAsia 字体（Word 的 ascii/eastAsia 双槽模型）；时间：2026-10-30。
    let stack = fontStackFor(style.fontFamily).replace(/"/g, "'");
    if (style.eastAsiaFont && style.eastAsiaFont !== style.fontFamily) {
      const eastAsiaStack = fontStackFor(style.eastAsiaFont).replace(/"/g, "'");
      const firstComma = stack.indexOf(',');
      stack =
        firstComma === -1
          ? `${stack}, ${eastAsiaStack}`
          : `${stack.slice(0, firstComma)}, ${eastAsiaStack}${stack.slice(firstComma)}`;
    }
    css.push(`font-family:${stack}`);
  }
  if (style.fontSizePt) css.push(`font-size:${style.fontSizePt}pt`);
  if (style.bold) css.push('font-weight:700');
  if (style.italic) css.push('font-style:italic');
  const decorations: string[] = [];
  if (style.underline) decorations.push('underline');
  if (style.strike) decorations.push('line-through');
  if (decorations.length) css.push(`text-decoration:${decorations.join(' ')}`);
  if (style.color) css.push(`color:${style.color}`);
  if (style.highlight) css.push(`background-color:${style.highlight}`);
  if (style.letterSpacingPx) css.push(`letter-spacing:${style.letterSpacingPx}px`);
  if (style.allCaps) css.push('text-transform:uppercase');
  if (style.smallCaps) css.push('font-variant:small-caps');
  return css.join(';');
}

function imageHtml(run: DocRun, context: RunRenderContext = {}): string {
  const image = run.image;
  if (!image) return '';
  const alt = escapeHtml(image.alt ?? image.name ?? '');
  const imageId = escapeHtml(image.id ?? image.relId ?? '');
  const wrap = image.wrap ?? 'inline';

  // 功能：VML rect 绝对定位浮层，横坐标按参照系换算到页面坐标，纵坐标留待挂载后按锚定段落实测补偿；
  // 不占行内空间；时间：2026-10-30。
  if (image.vmlRect) {
    const { left, top, width, height, xRelative, checked } = image.vmlRect;
    const stroke = image.strokeWidthPx ?? 1;
    const color = image.strokeColor ?? '#000000';
    // 功能：方框内的勾用真实对钩绘制（字号随框高走），勾选状态存于模型并可写回 v:textbox；时间：2026-10-30。
    const check = checked
      ? `<span data-vml-check="1" style="position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-size:${Math.max(8, height * 0.95)}px;line-height:1;font-family:'Segoe UI Symbol','宋体',sans-serif;color:${escapeHtml(color)};">✓</span>`
      : '';
    // 功能：方框的包含块就是所属段落文本框（paragraphTextStyle 设 position:relative），
    // left 由「页面横坐标 − 栏左缘 − 段落左缩进」换算成段落内偏移，纵向直接用 margin-top；
    // 这样缩放、字体加载与重排都不会让方框漂移，语义与 Word 的 relative-to:paragraph 一致；
    // 时间：2026-10-30。
    const columnLeft = context.vmlColumnLeftPx ?? 0;
    const indent = context.vmlIndentLeftPx ?? 0;
    const paragraphLeft = (xRelative === 'page' ? left - columnLeft : left) - indent;
    return `<span data-vml-rect="1" data-vml-id="${escapeHtml(image.id ?? '')}" data-vml-checked="${checked ? '1' : '0'}" data-vml-left="${left}" data-vml-top="${top}" data-vml-xrel="${xRelative ?? 'text'}" data-vml-stroke="${escapeHtml(color)}" data-vml-stroke-width="${stroke}" data-vml-width="${width}" data-vml-height="${height}" contenteditable="false" title="点击勾选" style="position:absolute;left:${paragraphLeft}px;top:${top}px;width:${width}px;height:${height}px;box-sizing:border-box;border:${Math.max(0.5, stroke)}px solid ${escapeHtml(color)};line-height:0;cursor:pointer;pointer-events:auto;" role="checkbox" aria-checked="${checked ? 'true' : 'false'}">${check}</span>`;
  }

  const layout = wrap === 'square-left'
    ? 'float:left;margin:0 12px 8px 0;'
    : wrap === 'square-right'
      ? 'float:right;margin:0 0 8px 12px;'
      : wrap === 'top-bottom'
        ? 'display:block;margin:8px auto;'
        : 'display:inline-block;vertical-align:baseline;';
  const frameStyle = `position:relative;line-height:0;${layout}`;
  const imageStyle = `display:block;width:${image.widthPx}px;height:${image.heightPx}px;max-width:100%;object-fit:contain;`;
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

/**
 * 功能：内容控件渲染成可交互 chip——复选框显示勾选方块、下拉/日期显示当前值与提示图标；
 * 全部 contenteditable=false，点击由编辑器接管，避免输入时破坏控件结构；时间：2026-10-30。
 */
function controlHtml(run: DocRun, runIndex: number): string | undefined {
  const control = run.control;
  if (!control) return undefined;
  const text = run.text ?? '';
  const style = runStyleCss(run.style);
  const box = 'display:inline-block;box-sizing:border-box;vertical-align:baseline;';
  if (control.kind === 'checkbox') {
    const mark = control.checked ? '☑' : '☐';
    return `<span data-control="checkbox" data-run-index="${runIndex}" data-control-checked="${control.checked ? '1' : '0'}" contenteditable="false" title="点击勾选" role="checkbox" aria-checked="${control.checked ? 'true' : 'false'}" style="${style}${style ? ';' : ''}${box}cursor:pointer;user-select:none;">${mark}</span>`;
  }
  const chip = 'border:1px solid #bfbfbf;border-radius:3px;padding:0 4px;background:#f3f3f3;color:inherit;cursor:pointer;user-select:none;';
  const value = text || control.alias || '请选择';
  const placeholder = control.placeholder || !text ? ';opacity:0.65;font-style:italic;' : '';
  // 功能：候选与日期格式随 DOM 一起输出，编辑回读时据此还原控件而无需再查模型；时间：2026-10-30。
  const meta = [
    `data-control-value="${escapeHtml(text)}"`,
    control.choices?.length ? `data-control-choices="${escapeHtml(JSON.stringify(control.choices))}"` : '',
    control.dateFormat ? `data-control-format="${escapeHtml(control.dateFormat)}"` : '',
    control.placeholder || !text ? 'data-control-placeholder="1"' : '',
  ].filter(Boolean).join(' ');
  return `<span data-control="${control.kind}" data-run-index="${runIndex}" ${meta} contenteditable="false" title="${control.kind === 'date' ? '点击选择日期' : '点击选择内容'}" style="${style}${style ? ';' : ''}${box}${chip}${placeholder}">${escapeHtml(value)}<span style="margin-left:3px;font-size:0.85em;">▾</span></span>`;
}

export function renderRunsHtml(runs: DocRun[], context: RunRenderContext = {}): string {
  const parts: string[] = [];

  // 缓冲区：合并连续的文本 run（逐字母模式 + 负字距），避免多栏布局下单词被拆分；
  // 时间：2026-10-30。
  let bufText = '';
  let bufCss = '';

  const flushBuf = () => {
    if (!bufText) return;
    const escaped = escapeHtml(bufText).replace(/\n/g, '<br>');
    parts.push(bufCss ? `<span style="${bufCss} white-space:nowrap;">${escaped}</span>` : escaped);
    bufText = '';
    bufCss = '';
  };

  // 制表位渲染状态：currentX 为已渲染内容的估计横坐标（CSS px）；
  // pendingAlign 为待渲染段落的对齐方式（center/right 由制表位决定）；
  // pendingTargetX 为该段落应对齐的目标位置，pendingStartX 为段落在行内的起点。
  let currentX = 0;
  let pendingAlign: 'center' | 'right' | null = null;
  let pendingTargetX = 0;
  let pendingStartX = 0;
  let pendingParts: string[] = [];

  const charWidthPx = (run: DocRun): number => {
    const sizePt = run.style.fontSizePt ?? 10.5;
    // 估算平均字符宽度 ≈ 字号 × 0.55（中英文混合近似），用于制表位位置估算；
    // pt → px ≈ ×0.75；时间：2026-10-30。
    return sizePt * 0.55 * 0.75;
  };

  const flushPending = () => {
    if (!pendingParts.length || !pendingAlign) return;
    const gap = Math.max(0, pendingTargetX - pendingStartX);
    const contentWidth = Math.max(0, currentX - pendingStartX);
    // 功能：右/居中制表位下内容比空隙宽时不再折行——按内容宽度撑开盒子并向左溢出错位，
    // 保证右缘仍落在制表位上（Word 的分行处理无法在流式编辑中复现）；时间：2026-10-30。
    const overflow = Math.max(0, contentWidth - gap);
    const width = gap + overflow;
    const marginLeft = overflow > 0 ? -overflow : 0;
    parts.push(
      `<span style="display:inline-block;width:${width}px;${marginLeft ? `margin-left:${marginLeft}px;` : ''}text-align:${pendingAlign};white-space:nowrap">${pendingParts.join('')}</span>`,
    );
    currentX = pendingTargetX;
    pendingAlign = null;
    pendingTargetX = 0;
    pendingStartX = 0;
    pendingParts = [];
  };

  const tabStops = context.tabStops ?? [];
  const usedTabStops = new Set<number>();
  // 功能：默认制表位——Word 在无显式制表位时按 defaultTabStop（缺省 720 缇 = 48px）整数倍跳转；时间：2026-10-30。
  const defaultTabPitch = context.defaultTabStopPx && context.defaultTabStopPx > 0 ? context.defaultTabStopPx : 48;

  const findNextTabStop = (afterX: number): WordTabStop => {
    const explicit = tabStops.find((t, i) => t.posPx > afterX && !usedTabStops.has(i));
    if (explicit) return explicit;
    const next = (Math.floor(afterX / defaultTabPitch) + 1) * defaultTabPitch;
    // 功能：显式制表位用尽后仍按默认制表位跳转，避免回退成固定宽度空格导致行宽失真；时间：2026-10-30。
    return { posPx: next, kind: 'left' };
  };

  const markTabStopUsed = (t: WordTabStop) => {
    const idx = tabStops.findIndex((s) => s.posPx === t.posPx && s.kind === t.kind);
    if (idx >= 0) usedTabStops.add(idx);
  };

  for (const [runIndex, run] of runs.entries()) {
    if (run.textbox) {
      // 功能：文本框只读渲染——段落锚定的绝对定位容器，内容用正文渲染器；时间：2026-10-31。
      flushPending();
      flushBuf();
      const box = run.textbox;
      const columnLeft = context.vmlColumnLeftPx ?? 0;
      const indent = context.vmlIndentLeftPx ?? 0;
      const left = (box.xRelative === 'page' ? box.left - columnLeft : box.left) - indent;
      const inner = box.blocks
        .map((block) => (block.type === 'paragraph' ? `<div data-textbox-block>${renderParagraphHtml(block, context)}</div>` : ''))
        .join('');
      parts.push(`<span data-textbox="1" data-textbox-left="${box.left}" data-textbox-top="${box.top}" data-textbox-width="${box.width}" data-textbox-height="${box.height}" data-textbox-xrel="${box.xRelative ?? 'text'}" contenteditable="false" aria-hidden="true" style="position:absolute;left:${left}px;top:${box.top}px;width:${box.width}px;height:${box.height}px;box-sizing:border-box;background:#ffffff;overflow:hidden;text-align:left;line-height:normal;">${inner}</span>`);
      continue;
    }
    if (run.control) {
      // 功能：控件 chip 作为整体渲染，不并入文本缓冲；时间：2026-10-30。
      flushPending();
      flushBuf();
      const html = controlHtml(run, runIndex);
      if (html) parts.push(html);
      continue;
    }
    if (run.image) {
      flushPending();
      flushBuf();
      parts.push(imageHtml(run, context));
      continue;
    }
    if (run.lineBreak) {
      flushPending();
      flushBuf();
      parts.push('<br>');
      continue;
    }
    if (run.pageBreak) {
      flushPending();
      flushBuf();
      parts.push('<span data-page-break="1" contenteditable="false" style="display:inline-block;width:0;height:0;overflow:hidden;"></span>');
      continue;
    }
    if (run.tab) {
      flushPending();
      flushBuf();
      const tabStop = findNextTabStop(currentX);
      markTabStopUsed(tabStop);
      const gap = Math.max(0, tabStop.posPx - currentX);
      if (tabStop.kind === 'left' || tabStop.kind === 'decimal') {
        parts.push(`<span style="display:inline-block;width:${gap}px"></span>`);
        currentX = tabStop.posPx;
      } else {
        // center / right：将后续文本缓冲，在目标位置内按 text-align 对齐；
        pendingAlign = tabStop.kind;
        pendingTargetX = tabStop.posPx;
        pendingStartX = currentX;
      }
      continue;
    }
    if (run.mergeField) {
      flushPending();
      flushBuf();
      parts.push(`<span data-merge-field="${escapeHtml(run.mergeField)}" contenteditable="false" style="${runStyleCss(run.style)}"><span style="background:#e8f1ff;border-radius:3px;color:#245694;">«${escapeHtml(run.mergeField)}»</span></span>`);
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

    const escaped = escapeHtml(text).replace(/\n/g, '<br>');
    let css = runStyleCss(run.style);
    // 功能：批注范围高亮 + 修订标记渲染（插入绿底/删除红底删除线，Word 惯例配色）；时间：2026-10-31。
    const commentAttr = run.commentIds?.length ? ` data-comment-ids="${escapeHtml(run.commentIds.join(','))}"` : '';
    let revOpen = '';
    let revClose = '';
    if (run.revision) {
      css += run.revision.type === 'insert' ? ';background-color:#e2f0d9;' : ';background-color:#fbdbdb;text-decoration:line-through;';
      revOpen = `<span data-revision="${run.revision.type}"${run.revision.author ? ` data-revision-author="${escapeHtml(run.revision.author)}"` : ''} contenteditable="false" style="display:inline;">`;
      revClose = '</span>';
    }
    if (commentAttr) {
      css += ';background-color:#fff3bf;';
    }

    // 逐字母模式：连续文本 run 且有 letterSpacing，合并到 buffer，包在 nowrap span 里；
    // 普通文本（无 letterSpacing）直接渲染，遇到 buffer 非空时先 flush。
    const isPlainText = !run.field && !run.anchor && !run.href && run.style.verticalAlign !== 'super' && run.style.verticalAlign !== 'sub';
    // 功能：右/居中制表位的待对齐内容必须进入缓冲盒，否则会提前落到制表位之前；时间：2026-10-30。
    if (isPlainText && css && !pendingAlign && !commentAttr && !run.revision) {
      if (!bufCss) bufCss = css;
      bufText += text;
      currentX += text.length * charWidthPx(run);
    } else {
      flushBuf();
      if (pendingAlign) {
        // center / right 制表位下：缓冲到 pendingParts，等 flushPending 一并对齐；
        pendingParts.push(css ? `<span style="${css}">${escaped}</span>` : escaped);
        currentX += text.length * charWidthPx(run);
      } else {
        let html = css ? `<span${commentAttr} style="${css}">${escaped}</span>` : `<span${commentAttr}>${escaped}</span>`;
        if (run.style.verticalAlign === 'super') html = `<sup>${html}</sup>`;
        else if (run.style.verticalAlign === 'sub') html = `<sub>${html}</sub>`;
        if (run.revision) html = revOpen + html + revClose;
        if (run.anchor) {
          html = `<a data-anchor="${escapeHtml(run.anchor)}" style="color:inherit;text-decoration:none;cursor:pointer;">${html}</a>`;
        } else if (run.href) {
          html = `<a href="${escapeHtml(run.href)}" style="color:#0563c1;text-decoration:underline;cursor:pointer;">${html}</a>`;
        }
        parts.push(html);
        currentX += text.length * charWidthPx(run);
      }
    }
  }

  flushBuf();
  flushPending();
  return parts.join('');
}

/** 功能：判断是否为 Word 自动目录段落（TOC1~TOC9 样式），作为制表位信息缺失时的兜底；时间：2026-10-30。 */
function isTocParagraph(paragraph: DocParagraph): boolean {
  return !!paragraph.styleId && /^TOC\s*\d+$/i.test(paragraph.styleId);
}

/** 功能：段落含「右对齐/居中制表位 + 前导符」，应渲染为「标题 + 前导点 + 页码右对齐」；时间：2026-10-30。 */
function hasLeaderTab(paragraph: DocParagraph): boolean {
  return (
    paragraph.style.tabStops?.some(
      (t) => (t.kind === 'right' || t.kind === 'center') && t.leader && t.leader !== 'none',
    ) ?? false
  );
}

/** 功能：是否走前导点布局（标准 TOC，或带前导符制表位的手动目录如样式"7"）；时间：2026-10-30。 */
function useLeaderLayout(paragraph: DocParagraph): boolean {
  return hasLeaderTab(paragraph) || isTocParagraph(paragraph);
}

/**
 * 功能：前导符字符——用真实字符绘制点线，密度与 Word 一致（border 虚线过于稀疏，且打印会丢）；
 * 时间：2026-10-30。
 */
const LEADER_GLYPH: Record<string, string> = {
  dot: '.',
  middleDot: '·',
  hyphen: '-',
  underline: '_',
};

/**
 * 功能：把目录类段落渲染成「标题 + 前导点 + 页码右对齐」的 flex 结构，对齐 Word/WPS 目录外观；
 * 前导符按制表位 leader 类型选择字符（dot 用点、hyphen 用连字符）；时间：2026-10-30。
 */
function renderLeaderTabHtml(paragraph: DocParagraph, context: RunRenderContext): string {
  const runs = paragraph.runs;
  const tabIndex = runs.findIndex((run) => run.tab);
  if (tabIndex === -1) return renderRunsHtml(runs, context) || '<br>';
  const left = runs.slice(0, tabIndex);
  const right = runs.slice(tabIndex + 1);
  const leftHtml = renderRunsHtml(left, context) || '';
  const rightHtml = renderRunsHtml(right, context) || '';
  const tabStop = paragraph.style.tabStops?.find(
    (t) => (t.kind === 'right' || t.kind === 'center') && t.leader && t.leader !== 'none',
  );
  const leaderChar = LEADER_GLYPH[tabStop?.leader ?? 'dot'] ?? '.';
  // 功能：前导点用足够多的真实字符铺满，再由 overflow 裁剪到弹性宽度，点的疏密与 Word 接近；时间：2026-10-30。
  const leaderHtml = leaderChar.repeat(240);
  return (
    `<span class="mo-toc-title" style="flex:0 0 auto;white-space:pre-wrap;">${leftHtml}</span>` +
    `<span class="mo-toc-leader" contenteditable="false" style="flex:1 1 auto;overflow:hidden;white-space:nowrap;letter-spacing:0.22em;margin:0 4px;min-width:10px;">${leaderHtml}</span>` +
    `<span class="mo-toc-page" style="flex:0 0 auto;white-space:pre-wrap;">${rightHtml}</span>`
  );
}

export function renderParagraphHtml(
  paragraph: DocParagraph,
  context: RunRenderContext = {},
): string {
  let marker = '';
  // 功能：空标记（numFmt="none" 的无编号列表）不渲染占位 span；时间：2026-10-31。
  if (paragraph.list && paragraph.list.marker) {
    marker = `<span data-list-marker="1" contenteditable="false" style="display:inline-block;min-width:1.5em;white-space:nowrap;">${escapeHtml(
      paragraph.list.marker,
    )}</span>`;
  }
  const content = useLeaderLayout(paragraph)
    ? renderLeaderTabHtml(paragraph, context)
    : renderRunsHtml(paragraph.runs, {
        ...context,
        // 功能：制表位位置以版面文字边界为基准，段落左缩进后需从相对坐标中扣除；时间：2026-10-30。
        tabStops: paragraph.style.tabStops?.map((stop) => ({
          ...stop,
          posPx: stop.posPx - (paragraph.style.indentLeftPx ?? 0),
        })),
        // 功能：VML 浮层的段落内横向偏移需要扣掉左缩进；时间：2026-10-30。
        vmlIndentLeftPx: paragraph.style.indentLeftPx ?? 0,
      });
  return `${marker}${content || '<br>'}`;
}

export function paragraphWrapperStyle(paragraph: DocParagraph): React.CSSProperties {
  const style = paragraph.style;
  const borderCss = (border: NonNullable<ParagraphStyle['borders']>['top']) => border && border.style !== 'none' ? `${Math.max(0.5, border.widthPx)}px ${border.style} ${border.color}` : undefined;
  return {
    paddingTop: style.spaceBeforePx ?? 0,
    paddingBottom: style.spaceAfterPx ?? 0,
    marginLeft: style.indentLeftPx ?? 0,
    marginRight: style.indentRightPx ?? 0,
    textAlign: style.alignment ?? 'left',
    backgroundColor: style.backgroundColor,
    borderTop: borderCss(style.borders?.top),
    borderRight: borderCss(style.borders?.right),
    borderBottom: borderCss(style.borders?.bottom),
    borderLeft: borderCss(style.borders?.left),
  };
}

export function paragraphTextStyle(paragraph: DocParagraph, defaultFont: string, defaultFontSizePt: number, linePitchPx?: number): React.CSSProperties {
  const style: ParagraphStyle = paragraph.style;
  // 功能：计算段落自然行高（CSS px）——优先用逐字体实测的自然行高（与渲染引擎同源、按字体真实度量），
  // 无 DOM 环境时回退系数近似；Word 未写 w:spacing 时是「单倍行距」（自然行高），不是 1.5 倍；时间：2026-10-31。
  const fontStack = fontStackFor(defaultFont) || DOC_FONT_STACK;
  const naturalSinglePx = measureNaturalLineHeightPx(fontStack, defaultFontSizePt);
  const naturalLineHeightPx = style.lineHeightExactPx || style.lineHeightAtLeastPx
    ? style.lineHeightExactPx ?? style.lineHeightAtLeastPx
    : naturalSinglePx > 0
      ? naturalSinglePx * (style.lineHeight ?? 1)
      : (defaultFontSizePt * (style.lineHeight ?? 1) * FONT_LINE_HEIGHT_RATIO) / 0.75;
  // 功能：若文档定义了 w:docGrid w:linePitch，段落行高向上取整到网格行高的整数倍，解决封面空行偏矮、整体上移问题；
  // 固定行距与显式关闭「对齐到网格」的段落不参与取整（与 Word 一致）；时间：2026-10-30。
  const snapToGrid = style.snapToGrid !== false && style.lineHeightExactPx === undefined;
  const lineHeightValue = linePitchPx && snapToGrid && naturalLineHeightPx
    ? Math.ceil(naturalLineHeightPx / linePitchPx) * linePitchPx
    : naturalLineHeightPx;
  const base: React.CSSProperties = {
    fontFamily: fontStack,
    fontSize: `${defaultFontSizePt}pt`,
    // 功能：作为 VML 浮层的包含块，让方框按段落相对坐标定位；时间：2026-10-30。
    position: 'relative',
    // 功能：固定值与最小值行距都用 px 表达，CSS 无“至少”语义，取相同高度即可保证分页一致；时间：2026-10-29；作者：lq。
    lineHeight: `${lineHeightValue}px`,
    textIndent: style.firstLineIndentPx ? `${style.firstLineIndentPx}px` : undefined,
    whiteSpace: 'pre-wrap',
    overflowWrap: 'break-word',
    wordBreak: 'break-word',
  };
  // 功能：目录类段落改用 flex，让前导点与页码右对齐；时间：2026-10-30。
  if (useLeaderLayout(paragraph)) {
    return { ...base, display: 'flex', alignItems: 'baseline' };
  }
  return base;
}


/**
 * 功能：段落的有效行高（CSS px）——固定值/atLeast 直接用，其余按逐字体自然行高 × 倍数；
 * 供孤寡控制把断页高度换算成行数；时间：2026-10-31。
 */
export function paragraphLineHeightPx(paragraph: DocParagraph, defaultFont: string, defaultFontSizePt: number): number {
  if (paragraph.style.lineHeightExactPx) return paragraph.style.lineHeightExactPx;
  const fontStack = fontStackFor(defaultFont) || DOC_FONT_STACK;
  const natural = measureNaturalLineHeightPx(fontStack, defaultFontSizePt);
  const single = natural > 0 ? natural : (defaultFontSizePt * FONT_LINE_HEIGHT_RATIO) / 0.75;
  const multiple = paragraph.style.lineHeight ?? 1;
  const value = single * multiple;
  return paragraph.style.lineHeightAtLeastPx ? Math.max(value, paragraph.style.lineHeightAtLeastPx) : value;
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
