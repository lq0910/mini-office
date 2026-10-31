import type { DocBlock, DocImage, DocParagraph, DocRun, RunStyle } from '@mini-office/document-model';
import { FONT_CHOICES, fontStackFor } from '@mini-office/design-system';
import { uid } from '@mini-office/shared';

const NAMED_COLORS: Record<string, string> = {
  black: '#000000',
  white: '#FFFFFF',
  red: '#FF0000',
  green: '#008000',
  blue: '#0000FF',
  yellow: '#FFFF00',
  gray: '#808080',
  grey: '#808080',
  orange: '#FFA500',
  purple: '#800080',
  transparent: '',
};

export function cssColorToHex(value: string | undefined | null): string | undefined {
  if (!value) return undefined;
  const color = value.trim().toLowerCase();
  if (!color || color === 'transparent' || color === 'rgba(0, 0, 0, 0)') return undefined;
  if (color.startsWith('#')) {
    if (/^#[0-9a-f]{3}$/.test(color)) {
      return `#${color
        .slice(1)
        .split('')
        .map((char) => char + char)
        .join('')
        .toUpperCase()}`;
    }
    if (/^#[0-9a-f]{6}$/.test(color)) return color.toUpperCase();
    return undefined;
  }
  const rgb = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(color);
  if (rgb) {
    return `#${[rgb[1], rgb[2], rgb[3]]
      .map((part) => Number(part).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase()}`;
  }
  return NAMED_COLORS[color] || undefined;
}

export function parseFontSizePt(value: string | undefined | null): number | undefined {
  if (!value) return undefined;
  const match = /^([\d.]+)\s*(pt|px|em|rem)?$/i.exec(value.trim());
  if (!match) return undefined;
  const number = Number(match[1]);
  if (!Number.isFinite(number)) return undefined;
  const unit = (match[2] ?? 'px').toLowerCase();
  if (unit === 'pt') return number;
  if (unit === 'px') return number * 0.75;
  if (unit === 'em' || unit === 'rem') return number * 12;
  return number;
}

/**
 * 功能：把渲染用的字体栈首项还原为模型中的字体名；时间：2026-10-29；作者：lq。
 *
 * 背景：段落渲染时会把「宋体」展开成 `"Songti SC", "SimSun", …`，若直接把这个族名写回模型，
 * 用户输入一个字后字体就会变成 "Songti SC"，工具栏字体框也会变空。这里按 FONT_CHOICES
 * 反查一次，保证样式套用与输入后字体名保持稳定。
 */
const FAMILY_TO_CHOICE: Map<string, string> = (() => {
  const map = new Map<string, string>();
  for (const choice of FONT_CHOICES) {
    if (choice === '默认') continue;
    const first = fontStackFor(choice).split(',')[0]?.trim().replace(/^['"]|['"]$/g, '');
    if (first) map.set(first.toLowerCase(), choice);
  }
  return map;
})();

function firstFontFamily(value: string): string | undefined {
  const first = value.split(',')[0]?.trim().replace(/^['"]|['"]$/g, '');
  if (!first) return undefined;
  if (/^(serif|sans-serif|monospace|system-ui|inherit|initial|unset)$/i.test(first)) return undefined;
  return FAMILY_TO_CHOICE.get(first.toLowerCase()) ?? first;
}

function styleIsEqual(a: RunStyle, b: RunStyle): boolean {
  return (
    a.fontFamily === b.fontFamily &&
    a.fontSizePt === b.fontSizePt &&
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.strike === b.strike &&
    a.color === b.color &&
    a.highlight === b.highlight &&
    a.verticalAlign === b.verticalAlign &&
    a.letterSpacingPx === b.letterSpacingPx &&
    a.allCaps === b.allCaps &&
    a.smallCaps === b.smallCaps
  );
}

function mergeTextRuns(runs: DocRun[]): DocRun[] {
  const result: DocRun[] = [];
  for (const run of runs) {
    const previous = result.at(-1);
    if (
      previous &&
      !previous.image &&
      !previous.tab &&
      !previous.lineBreak &&
      !previous.field && !previous.mergeField &&
      // 功能：内容控件是独立整体，不能被相邻同样式文本合并掉；时间：2026-10-30。
      !previous.control &&
      !run.image &&
      !run.tab &&
      !run.lineBreak &&
      !run.field && !run.mergeField &&
      !run.control &&
      // 功能：批注范围与修订边界不能被合并掉（高亮与接受/拒绝依赖 run 边界）；时间：2026-10-31。
      (previous.commentIds?.join(',') ?? '') === (run.commentIds?.join(',') ?? '') &&
      previous.revision?.type === run.revision?.type &&
      styleIsEqual(previous.style, run.style)
    ) {
      previous.text = (previous.text ?? '') + (run.text ?? '');
    } else {
      result.push({ ...run });
    }
  }
  return result;
}

/** 功能：从 DOM 属性还原候选列表（JSON），解析失败时退回按分隔符拆分；时间：2026-10-30。 */
function parseControlChoices(raw: string): string[] {
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) return parsed.map((item) => String(item));
  } catch {
    // 回退到旧格式
  }
  return raw.split('');
}

/**
 * 功能：在 [start,end) 处插入一个制表符 run——折叠光标即插入，选区则整体替换为一个 Tab；
 * Word 标准：正文 Tab 跳到下一个制表位而不是插入空格；时间：2026-10-30。
 */
export function insertTabAt(runs: DocRun[], start: number, end: number, style: RunStyle): DocRun[] {
  const total = runs.reduce((sum, run) => sum + (run.tab || run.image || run.lineBreak || run.mergeField ? 1 : (run.text ?? '').length), 0);
  const safeStart = Math.max(0, Math.min(start, total));
  const safeEnd = Math.max(safeStart, Math.min(end, total));
  const tabRun: DocRun = { text: '', style: { ...style }, tab: true };
  const consumed = (run: DocRun) => (run.tab || run.image || run.lineBreak || run.mergeField ? 1 : (run.text ?? '').length);
  const output: DocRun[] = [];
  let position = 0;
  let inserted = false;
  for (const run of runs) {
    const runEnd = position + consumed(run);
    if (!inserted && runEnd > safeStart) {
      // 功能：插入点落在本 run 内或起点——保留删除区间前的尾部，插入 Tab，再保留删除区间后的头部；时间：2026-10-30。
      if (position < safeStart) output.push(sliceRunText(run, 0, safeStart - position));
      output.push(tabRun);
      inserted = true;
      if (runEnd > safeEnd) output.push(sliceRunText(run, safeEnd - position, (run.text ?? '').length));
    } else if (inserted && position < safeEnd && runEnd > safeEnd) {
      // 删除区间跨到本 run：只留区间之后的部分
      output.push(sliceRunText(run, safeEnd - position, (run.text ?? '').length));
    } else if (inserted && runEnd <= safeEnd) {
      // 完整落在删除区间内：丢弃
    } else {
      output.push(run);
    }
    position = runEnd;
  }
  if (!inserted) output.push(tabRun);
  return output;
}

/** 功能：按文本区间截取 run；原子 run（tab/图片/换行）原样返回；时间：2026-10-30。 */
function sliceRunText(run: DocRun, from: number, to: number): DocRun {
  if (run.tab || run.image || run.lineBreak || run.mergeField) return { ...run };
  return { ...run, text: (run.text ?? '').slice(from, to) };
}

function elementBoxWidth(node: Element): number {
  const width = parseFloat((node as HTMLElement).style.width ?? '');
  return Number.isFinite(width) ? width : 0;
}

function elementBoxHeight(node: Element): number {
  const height = parseFloat((node as HTMLElement).style.height ?? '');
  return Number.isFinite(height) ? height : 0;
}

export function serializeDomRuns(
  root: HTMLElement,
  fallback: RunStyle,
  imagesById: Map<string, DocImage>,
): DocRun[] {
  const runs: DocRun[] = [];

  const walk = (node: Node, style: RunStyle, commentIds?: string[], revision?: DocRun['revision']) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent ?? '';
      if (text) runs.push({ text, style: { ...style }, ...(commentIds?.length ? { commentIds } : {}), ...(revision ? { revision } : {}) });
      return;
    }
    if (!(node instanceof Element)) return;
    // 功能：批注范围与修订标记沿 DOM 继承到子文本；时间：2026-10-31。
    const nodeComments = node.getAttribute('data-comment-ids');
    const nodeRevision = node.getAttribute('data-revision');
    const nextComments = nodeComments ? nodeComments.split(',') : commentIds;
    let nextRevision = revision;
    if (nodeRevision === 'insert' || nodeRevision === 'delete') {
      nextRevision = { type: nodeRevision, ...(node.getAttribute('data-revision-author') ? { author: node.getAttribute('data-revision-author') as string } : {}) };
    }
    if (node.getAttribute('data-list-marker') === '1') return;
    // 功能：VML 浮层方框在编辑回读时还原为图片 run（含勾选状态），避免编辑段落时丢失复选框；时间：2026-10-30。
    if (node.hasAttribute('data-vml-rect')) {
      const number = (name: string, fallback: number) => {
        const value = Number(node.getAttribute(name));
        return Number.isFinite(value) && node.getAttribute(name) !== null ? value : fallback;
      };
      const xRelative = node.getAttribute('data-vml-xrel') === 'page' ? 'page' as const : 'text' as const;
      const width = number('data-vml-width', 0) || elementBoxWidth(node);
      const height = number('data-vml-height', 0) || elementBoxHeight(node);
      runs.push({
        text: '',
        style: { ...style },
        image: {
          id: node.getAttribute('data-vml-id') ?? undefined,
          src: '',
          widthPx: width,
          heightPx: height,
          wrap: 'none',
          vmlRect: {
            left: number('data-vml-left', 0),
            top: number('data-vml-top', 0),
            width,
            height,
            xRelative,
            ...(node.getAttribute('data-vml-checked') === '1' ? { checked: true } : {}),
          },
          strokeColor: node.getAttribute('data-vml-stroke') ?? undefined,
          strokeWidthPx: number('data-vml-stroke-width', 1),
        },
      });
      return;
    }
    // 功能：文本框在编辑回读时还原为带 textbox 的 run，内层段落按块重建；时间：2026-10-31。
    if (node.hasAttribute('data-textbox')) {
      const number = (name: string): number => {
        const value = Number(node.getAttribute(name));
        return Number.isFinite(value) && value !== 0 ? value : 0;
      };
      const width = number('data-textbox-width') || 100;
      const height = number('data-textbox-height') || 60;
      const blocks: DocBlock[] = Array.from(node.querySelectorAll('[data-textbox-block]')).map((blockEl) => {
        const inner = blockEl as HTMLElement;
        const runs = serializeDomRuns(inner, style, imagesById);
        return { type: 'paragraph' as const, id: uid('p'), runs, style: {} };
      });
      runs.push({
        text: '',
        style: { ...style },
        textbox: {
          left: number('data-textbox-left'),
          top: number('data-textbox-top'),
          width,
          height,
          xRelative: node.getAttribute('data-textbox-xrel') === 'page' ? 'page' as const : 'text' as const,
          blocks,
        },
      });
      return;
    }
    // 功能：内容控件 chip 在编辑回读时还原为带 control 的 run，保存后仍是可交互控件；时间：2026-10-30。
    const controlKind = node.getAttribute('data-control');
    if (controlKind) {
      const kind = controlKind as NonNullable<DocRun['control']>['kind'];
      const text = node.getAttribute('data-control-value') ?? '';
      const choices = node.getAttribute('data-control-choices');
      const format = node.getAttribute('data-control-format');
      runs.push({
        text,
        style: { ...style },
        control: {
          kind,
          ...(kind === 'checkbox' ? { checked: node.getAttribute('data-control-checked') === '1' } : {}),
          ...(choices ? { choices: parseControlChoices(choices) } : {}),
          ...(format ? { dateFormat: format } : {}),
          ...(node.getAttribute('data-control-placeholder') === '1' || !text ? { placeholder: true } : {}),
        },
      });
      return;
    }
    if (node.hasAttribute('data-tab')) { runs.push({ text: '', style: { ...style }, tab: true }); return; }
    const tag = node.tagName;
    if (tag === 'BR') {
      runs.push({ text: '', style: { ...style }, lineBreak: true });
      return;
    }
    if (tag === 'IMG') {
      const id = node.getAttribute('data-image-id') ?? '';
      const stored = imagesById.get(id);
      if (stored) {
        runs.push({ text: '', style: { ...style }, image: stored });
      } else {
        const src = (node as HTMLImageElement).src;
        if (src) {
          runs.push({
            text: '',
            style: { ...style },
            image: {
              id,
              src,
              widthPx: (node as HTMLImageElement).width || 160,
              heightPx: (node as HTMLImageElement).height || 120,
            },
          });
        }
      }
      return;
    }

    const next: RunStyle = { ...style };
    if (tag === 'B' || tag === 'STRONG') next.bold = true;
    if (tag === 'I' || tag === 'EM') next.italic = true;
    if (tag === 'U') next.underline = true;
    if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') next.strike = true;
    if (tag === 'SUP') next.verticalAlign = 'super';
    if (tag === 'SUB') next.verticalAlign = 'sub';
    if (tag === 'FONT') {
      const element = node as HTMLFontElement;
      if (element.face) next.fontFamily = element.face;
      const color = cssColorToHex(element.color);
      if (color) next.color = color;
      const sizeMap: Record<string, number> = {
        '1': 8,
        '2': 10,
        '3': 12,
        '4': 14,
        '5': 18,
        '6': 24,
        '7': 36,
      };
      if (element.size && sizeMap[element.size]) next.fontSizePt = sizeMap[element.size];
    }

    if (node instanceof HTMLElement) {
      const inline = node.style;
      if (inline.fontWeight) {
        const weight = inline.fontWeight;
        next.bold = weight === 'bold' || weight === 'bolder' || Number(weight) >= 600;
      }
      if (inline.fontStyle === 'italic' || inline.fontStyle === 'oblique') next.italic = true;
      if (inline.textDecorationLine?.includes('underline') || inline.textDecoration?.includes('underline')) {
        next.underline = true;
      }
      if (inline.textDecorationLine?.includes('line-through') || inline.textDecoration?.includes('line-through')) {
        next.strike = true;
      }
      const color = cssColorToHex(inline.color);
      if (color) next.color = color;
      const background = cssColorToHex(inline.backgroundColor);
      if (background) next.highlight = background;
      const fontSize = parseFontSizePt(inline.fontSize);
      if (fontSize) next.fontSizePt = fontSize;
      if (inline.fontFamily) {
        const family = firstFontFamily(inline.fontFamily);
        if (family) next.fontFamily = family;
      }
      if (inline.letterSpacing) {
        const spacing = parseFloat(inline.letterSpacing);
        if (Number.isFinite(spacing)) next.letterSpacingPx = spacing;
      }
      if (inline.textTransform === 'uppercase') next.allCaps = true;
      if (inline.fontVariant?.includes('small-caps')) next.smallCaps = true;
    }

    // 功能：字段原子沿用节点原有字形样式，输入不会将其拆散；时间：2026-10-24；作者：lq。
    if (node.hasAttribute('data-merge-field')) {
      const name = node.getAttribute('data-merge-field')!;
      runs.push({ text: `«${name}»`, mergeField: name, style: next });
      return;
    }
    // 功能：段内分页符标记还原为 pageBreak run，保证编辑后不丢分页位置；时间：2026-10-30；作者：lq。
    if (node.getAttribute('data-page-break') === '1') {
      runs.push({ text: '', style: next, pageBreak: true });
      return;
    }
    for (const child of Array.from(node.childNodes)) walk(child, next, nextComments, nextRevision);
  };

  for (const child of Array.from(root.childNodes)) walk(child, fallback);
  const merged = mergeTextRuns(runs);
  if (merged.length === 0) merged.push({ text: '', style: { ...fallback } });
  return merged;
}

function isAtomicElement(element: Element): boolean {
  return (
    element.tagName === 'IMG' ||
    element.tagName === 'BR' ||
    element.getAttribute('data-tab') === '1' ||
    element.getAttribute('data-page-break') === '1' ||
    element.hasAttribute('data-merge-field')
  );
}

/** Caret offset in atomic text units (images, tabs and breaks count as 1). */
function offsetAt(root: HTMLElement, container: Node, offset: number): number {
  const measure = document.createRange();
  measure.selectNodeContents(root);
  measure.setEnd(container, offset);
  const fragment = measure.cloneContents();
  let count = 0;
  const walk = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      count += (node.textContent ?? '').length;
      return;
    }
    if (!(node instanceof Element)) return;
    if (node.getAttribute('data-list-marker') === '1') return;
    if (isAtomicElement(node)) {
      count += 1;
      return;
    }
    for (const child of Array.from(node.childNodes)) walk(child);
  };
  for (const child of Array.from(fragment.childNodes)) walk(child);
  return count;
}

export function getCaretOffset(root: HTMLElement): number {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return 0;
  const range = selection.getRangeAt(0);
  if (!root.contains(range.endContainer)) return 0;
  return offsetAt(root, range.endContainer, range.endOffset);
}

/** Capture both ends of the current selection, even if focus moves to a toolbar select. */
export function getSelectionOffsets(root: HTMLElement): { start: number; end: number } | null {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return null;
  const range = selection.getRangeAt(0);
  if (!root.contains(range.startContainer) || !root.contains(range.endContainer)) return null;
  const start = offsetAt(root, range.startContainer, range.startOffset);
  const end = offsetAt(root, range.endContainer, range.endOffset);
  return { start: Math.min(start, end), end: Math.max(start, end) };
}

/** Place the caret at an atomic unit offset. */
export function setCaretAtOffset(root: HTMLElement, target: number): void {
  const selection = window.getSelection();
  if (!selection) return;
  let remaining = Math.max(0, target);
  let placed = false;

  const place = (node: Node, offset: number) => {
    const range = document.createRange();
    range.setStart(node, offset);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
    placed = true;
  };

  const walk = (node: Node): boolean => {
    if (node.nodeType === Node.TEXT_NODE) {
      const length = (node.textContent ?? '').length;
      if (remaining <= length) {
        place(node, remaining);
        return true;
      }
      remaining -= length;
      return false;
    }
    if (!(node instanceof Element)) return false;
    if (node.getAttribute('data-list-marker') === '1') return false;
    // 功能：光标始终位于字段块外侧，保证插域后继续输入的位置正确；时间：2026-10-24；作者：lq。
    if (isAtomicElement(node)) {
      if (remaining <= 0) {
        place(node.parentNode ?? root, Array.from((node.parentNode ?? root).childNodes).indexOf(node));
        return true;
      }
      remaining -= 1;
      return false;
    }
    for (const child of Array.from(node.childNodes)) {
      if (walk(child)) return true;
    }
    return false;
  };

  for (const child of Array.from(root.childNodes)) {
    if (walk(child)) break;
  }
  if (!placed) {
    const range = document.createRange();
    range.selectNodeContents(root);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
  }
}

export function getParagraphPlainTextFromDom(root: HTMLElement): string {
  let text = '';
  const walk = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      text += node.textContent ?? '';
      return;
    }
    if (!(node instanceof Element)) return;
    if (node.getAttribute('data-list-marker') === '1') return;
    if (node.tagName === 'BR') {
      text += '\n';
      return;
    }
    if (node.tagName === 'IMG' || node.hasAttribute('data-merge-field')) {
      text += '￼';
      return;
    }
    if (node.getAttribute('data-tab') === '1') {
      text += '\t';
      return;
    }
    // 功能：段内分页符在 DOM 纯文本中按一个换行计；时间：2026-10-30；作者：lq。
    if (node.getAttribute('data-page-break') === '1') {
      text += '\n';
      return;
    }
    for (const child of Array.from(node.childNodes)) walk(child);
  };
  for (const child of Array.from(root.childNodes)) walk(child);
  return text;
}

export function runsPlainText(runs: DocRun[]): string {
  return runs
    .map((run) => {
      if (run.image || run.mergeField) return '￼';
      if (run.tab) return '\t';
      if (run.lineBreak) return '\n';
      // 功能：段内分页符在纯文本中按一个换行计，与 offset 口径一致；时间：2026-10-30；作者：lq。
      if (run.pageBreak) return '\n';
      return run.text ?? '';
    })
    .join('');
}

export interface FindHit {
  paragraphId: string;
  index: number;
  length: number;
  pageIndex?: number;
}

export function findInParagraph(
  paragraph: DocParagraph,
  query: string,
  caseSensitive = false,
): Array<{ index: number; length: number }> {
  if (!query) return [];
  const text = runsPlainText(paragraph.runs);
  const hits: Array<{ index: number; length: number }> = [];
  if (caseSensitive) {
    let from = 0;
    while (from <= text.length) {
      const index = text.indexOf(query, from);
      if (index < 0) break;
      hits.push({ index, length: query.length });
      from = index + Math.max(1, query.length);
    }
  } else {
    const haystack = text.toLowerCase();
    const needle = query.toLowerCase();
    let from = 0;
    while (from <= haystack.length) {
      const index = haystack.indexOf(needle, from);
      if (index < 0) break;
      hits.push({ index, length: query.length });
      from = index + Math.max(1, query.length);
    }
  }
  return hits;
}

export function replaceInRuns(
  runs: DocRun[],
  query: string,
  replacement: string,
  limit: number,
  caseSensitive = false,
): { runs: DocRun[]; count: number } {
  if (!query) return { runs, count: 0 };
  let remaining = limit;
  let count = 0;
  const next = runs.map((run) => {
    if (remaining <= 0 || !run.text || run.image || run.tab || run.mergeField) return run;
    let text = run.text;
    let changed = true;
    while (changed && remaining > 0) {
      changed = false;
      const haystack = caseSensitive ? text : text.toLowerCase();
      const needle = caseSensitive ? query : query.toLowerCase();
      const index = haystack.indexOf(needle);
      if (index >= 0) {
        text = text.slice(0, index) + replacement + text.slice(index + query.length);
        remaining -= 1;
        count += 1;
        changed = true;
      }
    }
    return text === run.text ? run : { ...run, text };
  });
  return { runs: next, count };
}

export function stripRunsForSearch(runs: DocRun[]): string {
  return runsPlainText(runs);
}

interface DomPosition {
  node: Node;
  offset: number;
}

function locateOffset(root: HTMLElement, target: number, bias: 'start' | 'end'): DomPosition | null {
  let remaining = Math.max(0, target);
  let result: DomPosition | null = null;

  const visit = (node: Node): boolean => {
    if (node.nodeType === Node.TEXT_NODE) {
      const length = (node.textContent ?? '').length;
      if (remaining <= length) {
        result = { node, offset: remaining };
        return true;
      }
      remaining -= length;
      return false;
    }
    if (!(node instanceof Element)) return false;
    if (node.getAttribute('data-list-marker') === '1') return false;
    if (isAtomicElement(node)) {
      if (remaining <= 0) {
        const parent = node.parentNode ?? root;
        const index = Array.from(parent.childNodes).indexOf(node);
        result = { node: parent, offset: bias === 'end' ? index + 1 : index };
        return true;
      }
      remaining -= 1;
      return false;
    }
    for (const child of Array.from(node.childNodes)) {
      if (visit(child)) return true;
    }
    return false;
  };

  for (const child of Array.from(root.childNodes)) {
    if (visit(child)) break;
  }
  if (!result) {
    result = { node: root, offset: root.childNodes.length };
  }
  return result;
}

export function selectRangeByOffset(
  root: HTMLElement,
  start: number,
  end: number,
  focus = true,
): void {
  const startPosition = locateOffset(root, start, 'start');
  const endPosition = locateOffset(root, end, 'end');
  if (!startPosition || !endPosition) return;
  const selection = window.getSelection();
  if (!selection) return;
  try {
    const range = document.createRange();
    range.setStart(startPosition.node, startPosition.offset);
    range.setEnd(endPosition.node, endPosition.offset);
    selection.removeAllRanges();
    selection.addRange(range);
    if (focus) root.focus({ preventScroll: true });
  } catch {
    // The DOM may have changed between lookup and selection.
  }
}


/** 功能：将跨段落 DOM 选区转换为各段原子偏移，不包含页眉页脚；时间：2026-10-18；作者：lq。 */
export function selectedParagraphSlices(root: HTMLElement, range: Range): Array<{ paragraphId: string; start: number; end: number }> {
  if (range.collapsed || !root.contains(range.startContainer) || !root.contains(range.endContainer)) return [];
  const slices: Array<{ paragraphId: string; start: number; end: number }> = [];
  root.querySelectorAll<HTMLElement>('[data-paragraph-id]').forEach((element) => {
    if (!range.intersectsNode(element)) return;
    const start = element.contains(range.startContainer) ? offsetAt(element, range.startContainer, range.startOffset) : 0;
    const end = element.contains(range.endContainer) ? offsetAt(element, range.endContainer, range.endOffset)
      : offsetAt(element, element, element.childNodes.length);
    if (end > start || !element.textContent) slices.push({ paragraphId: element.dataset.paragraphId!, start, end });
  });
  return slices;
}

