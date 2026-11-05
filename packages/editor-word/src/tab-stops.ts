/**
 * 功能：Word 制表位的定位与渲染后置处理；时间：2026-10-30；作者：lq。
 *
 * 为什么需要「渲染后再算宽度」：
 * 制表位的宽度取决于「制表符当前所在的横向位置」，而位置又取决于前面的文字排到哪里，
 * 这只有浏览器排版完才知道。因此渲染时先输出零宽占位块，再由本模块按实测位置补上宽度。
 * 一次左→右的顺序遍历即可收敛：每个制表符只受它左边的制表符影响。
 */
import type { TabAlignment, TabLeader, TabStop } from '@mini-office/document-model';

/** Word 默认制表位间距：720 twips = 0.5 英寸 = 48px。 */
export const DEFAULT_TAB_STOP_PX = 48;

export interface TabTarget {
  posPx: number;
  alignment: TabAlignment;
  leader?: TabLeader;
}

/**
 * 功能：求出 x 右侧的下一个制表位——优先用自定义制表位，没有则按默认间距等距推进；
 * 时间：2026-10-30；作者：lq。
 */
export function nextTabStop(stops: TabStop[] | undefined, x: number, defaultStepPx = DEFAULT_TAB_STOP_PX): TabTarget {
  const step = defaultStepPx > 0 ? defaultStepPx : DEFAULT_TAB_STOP_PX;
  const list = (stops ?? []).filter((stop) => Number.isFinite(stop.posPx)).sort((a, b) => a.posPx - b.posPx);
  for (const stop of list) {
    if (stop.posPx > x + 0.5) return { posPx: stop.posPx, alignment: stop.alignment, leader: stop.leader };
  }
  // 越过最后一个自定义制表位后回到默认间距，与 Word 的行为一致。
  const next = (Math.floor((x + 0.5) / step) + 1) * step;
  return { posPx: next, alignment: 'left' };
}

/** 功能：制表位前导符转成下边框样式，点线/虚线与 Word 观感一致；时间：2026-10-30；作者：lq。 */
function leaderStyle(leader: TabLeader | undefined): string {
  switch (leader) {
    case 'dot':
    case 'middleDot':
      return '1px dotted currentColor';
    case 'hyphen':
      return '1px dashed currentColor';
    case 'underscore':
      return '1px solid currentColor';
    case 'heavy':
      return '2px solid currentColor';
    default:
      return '';
  }
}

/** 功能：量出制表符后方跨样式节点的文字宽度，多行时返回 null 表示无法右对齐；时间：2026-11-05；作者：lq。 */
function measureTrailing(tab: HTMLElement, paragraph: HTMLElement, scale: number): number | null {
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_ELEMENT);
  let end: Node | null = null;
  let passedTab = false;
  while (walker.nextNode()) {
    const element = walker.currentNode as HTMLElement;
    if (element === tab) {
      passedTab = true;
      continue;
    }
    if (!passedTab) continue;
    if (element.hasAttribute('data-tab') || element.tagName === 'BR') {
      end = element;
      break;
    }
  }
  const range = document.createRange();
  range.setStartAfter(tab);
  if (end) range.setEndBefore(end);
  else range.setEnd(paragraph, paragraph.childNodes.length);
  if (range.collapsed) return 0;
  // 功能：嵌套样式 span 与文字会在同一行返回多个矩形；按共同垂直带和水平并集量宽，不把它误判为换行；时间：2026-11-03；作者：lq。
  const rects = Array.from(range.getClientRects()).filter((rect) => rect.width > 0 && rect.height > 0);
  if (rects.length > 1 && Math.max(...rects.map((rect) => rect.top)) >= Math.min(...rects.map((rect) => rect.bottom))) return null;
  const width = rects.length ? Math.max(...rects.map((rect) => rect.right)) - Math.min(...rects.map((rect) => rect.left)) : range.getBoundingClientRect().width;
  return Number.isFinite(width) ? width / scale : null;
}

export interface TabStopSource {
  stops?: TabStop[];
  defaultStepPx?: number;
  /** 功能：段落相对正文栏的左缩进，用于恢复制表位的原始坐标基准；时间：2026-11-03；作者：lq。 */
  indentLeftPx?: number;
}

/**
 * 功能：按实测位置为容器内所有制表符占位块设置宽度与前导符；返回是否有改动。
 * 时间：2026-10-30；作者：lq。
 *
 * 说明：容器既可能是屏幕上的分页 DOM，也可能是离屏测量容器，两者的段落结构与宽度一致，
 * 因此同一套计算能得到相同结果，分页高度与屏幕显示不会互相打架。
 */
export function applyTabStops(
  root: HTMLElement,
  resolve: (paragraphId: string) => TabStopSource | undefined,
): boolean {
  let changed = false;
  const paragraphs = root.querySelectorAll<HTMLElement>('[data-paragraph-id]');
  paragraphs.forEach((paragraphElement) => {
    const tabs = Array.from(paragraphElement.querySelectorAll<HTMLElement>('[data-tab]'));
    if (tabs.length === 0) return;
    const config = resolve(paragraphElement.dataset.paragraphId ?? '');
    if (!config) return;
    const box = paragraphElement.getBoundingClientRect();
    const layoutWidth = paragraphElement.offsetWidth;
    const scale = layoutWidth > 0 ? box.width / layoutWidth : 1;
    if (!Number.isFinite(scale) || scale <= 0) return;
    // 功能：段落外层 margin-left 已移动文字盒，测量原点须退回正文栏以免二次叠加缩进；时间：2026-11-03；作者：lq。
    const origin = box.left - (config.indentLeftPx ?? 0) * scale;
    // 最多两轮：第一轮设宽度可能引起换行，第二轮修正被挤到下一行的制表符。
    for (let pass = 0; pass < 2; pass += 1) {
      let passChanged = false;
      for (const tab of tabs) {
        const rect = tab.getBoundingClientRect();
        const x = (rect.left - origin) / scale;
        const target = nextTabStop(config.stops, x, config.defaultStepPx);
        let width = Math.max(0, target.posPx - x);
        if (target.alignment === 'right' || target.alignment === 'center' || target.alignment === 'decimal') {
          const trailing = measureTrailing(tab, paragraphElement, scale);
          if (trailing !== null) {
            width = Math.max(0, target.posPx - x - (target.alignment === 'center' ? trailing / 2 : trailing));
          }
        }
        const cssWidth = `${width.toFixed(2)}px`;
        if (tab.style.width !== cssWidth) {
          tab.style.width = cssWidth;
          passChanged = true;
        }
        const leader = leaderStyle(target.leader);
        if ((tab.style.borderBottom || '') !== leader) {
          tab.style.borderBottom = leader;
          passChanged = true;
        }
      }
      if (!passChanged) break;
      changed = true;
    }
  });
  return changed;
}
