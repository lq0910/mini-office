/** 功能：跨页表格的同一行各单元格同步留出页间距离，装饰不进入保存或撤销；时间：2026-11-11；作者：lq。 */
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view';
import type { LayoutBatch } from '../upstream/apps/docs/src/renderer/editor/pagination-gaps';
// 功能：整组居中文字无法放入一页时，按可容纳的完整行绘制到各页正文区；时间：2026-11-11；作者：lq。
import { wordCellLineShifts } from './cell-line-placement';

const key = new PluginKey<DecorationSet>('miniOfficeCellGaps');
// 功能：纵向居中的跨页合并格单独持有显示位移，不写回原始对齐属性或正文；时间：2026-11-11；作者：lq。
const alignKey = new PluginKey<DecorationSet>('miniOfficeMergedCellAlignment');
/** 功能：只保存伴随占位；主页缝仍由同源分页插件计入页界台账，避免多列重复计算高度；时间：2026-11-11；作者：lq。 */
export const wordCellGaps = Extension.create({
  name: 'miniOfficeCellGaps',
  addProseMirrorPlugins: () => [key, alignKey].map(key => new Plugin<DecorationSet>({ key,
    state: { init: () => DecorationSet.empty, apply: (tr, old) => tr.getMeta(key) ?? old.map(tr.mapping, tr.doc) },
    props: { decorations: state => key.getState(state) },
  })),
});
/** 功能：一次事务同步所有单元格续行；同样的几何重用 widget，不产生整表重排；时间：2026-11-11；作者：lq。 */
export function setWordCellGaps(view: EditorView, gaps: Array<{ pos: number; height: number; page: number }>, batch: LayoutBatch): void {
  const signature = gaps.map(gap => `${gap.page}:${gap.pos}:${gap.height.toFixed(2)}`).join('|');
  const old = key.getState(view.state);
  if (old?.find()[0]?.spec.signature === signature || (!gaps.length && !old?.find().length)) return;
  const decorations = gaps.map(gap => Decoration.widget(gap.pos, () => {
    const element = document.createElement('span'); element.className = 'mo-word-cell-gap'; element.contentEditable = 'false';
    element.style.height = `${gap.height}px`; element.dataset.page = String(gap.page); return element;
  }, { side: -1, key: `cell-gap-${gap.page}-${gap.pos}-${gap.height.toFixed(2)}`, signature }));
  batch.set(key, DecorationSet.create(view.state.doc, decorations));
}

/** 功能：跨页行及合并格的短居中文字不能因分页占位落进页脚或页缝；按最近正文区整体移动；时间：2026-11-11；作者：lq。 */
export function alignWordTableCellText(view: EditorView, pages: readonly { top: number; bottom: number }[], factor: number, batch: LayoutBatch): void {
  const root = view.dom;
  const items: Array<{ pos: number; end: number; shift: number }> = [];
  const origin = root.getBoundingClientRect().top;
  // 功能：读几何时撤去上一轮纯显示平移，防止缩放/重排累计位移；时间：2026-11-11；作者：lq。
  root.classList.add('mo-measuring-cell-alignment');
  try {
    for (const cell of root.querySelectorAll<HTMLTableCellElement>('td, th')) {
      if (cell.closest('.page-gap, .page-repeat-header, .doc-nested-table')) continue;
      // 功能：显示层可能已将跨页行顶对齐，判断短居中格必须读取来源 vAlign，不能读校正后的计算值；时间：2026-11-11；作者：lq。
      const sourceAlign = (cell as unknown as { pmViewDesc?: { node?: { attrs: { vAlign?: string } } } }).pmViewDesc?.node?.attrs.vAlign;
      // 功能：顶对齐的纵向合并格也可能跨越行间页缝，长文字同样需要完整行避让；时间：2026-11-11；作者：lq。
      if (!/^(center|middle|bottom)$/.test(sourceAlign ?? cell.style.verticalAlign ?? '') && (Number(cell.dataset.baseRowspan) || cell.rowSpan) < 2) continue;
      const paragraphs = Array.from(cell.children).filter((el): el is HTMLElement => el instanceof HTMLElement && !!(el as unknown as { pmViewDesc?: { node?: { isTextblock: boolean } } }).pmViewDesc?.node?.isTextblock);
      if (!paragraphs.length) continue;
      const range = document.createRange(); let top = Infinity, bottom = -Infinity;
      for (const para of paragraphs) {
        // 功能：仅测真实字形，不把页缝 widget 的高度误认为居中文字的高度；时间：2026-11-11；作者：lq。
        const walker = document.createTreeWalker(para, NodeFilter.SHOW_TEXT);
        let text: Node | null;
        while ((text = walker.nextNode())) {
          if (!text.textContent?.trim() || text.parentElement?.closest('.page-gap, .mo-word-cell-gap, .page-repeat-header')) continue;
          range.selectNodeContents(text);
          for (const rect of range.getClientRects()) if (rect.width > 0.1 && rect.height > 0.1) {
            top = Math.min(top, (rect.top - origin) / factor); bottom = Math.max(bottom, (rect.bottom - origin) / factor);
          }
        }
      }
      if (!Number.isFinite(top)) continue;
      const cellRect = cell.getBoundingClientRect(), style = getComputedStyle(cell);
      const min = (cellRect.top - origin) / factor + (parseFloat(style.paddingTop) || 0);
      const max = (cellRect.bottom - origin) / factor - (parseFloat(style.paddingBottom) || 0);
      // 功能：整组文字必须能放入同一正文片段；长内容继续由原分页处理，不裁剪或丢弃文字；时间：2026-11-11；作者：lq。
      let shift: number | undefined;
      for (const page of pages) {
        const start = Math.max(min, page.top), end = Math.min(max, page.bottom);
        if (end - start < bottom - top) continue;
        const candidate = Math.max(start, Math.min(top, end - (bottom - top))) - top;
        if (shift === undefined || Math.abs(candidate) < Math.abs(shift)) shift = candidate;
      }
      if (shift === undefined) {
        // 功能：仅对跨页且无法整体容纳的长文字采集字符位置；正文模型与行宽保持不变；时间：2026-11-11；作者：lq。
        const segments = pages.map(page => ({ top: Math.max(min, page.top), bottom: Math.min(max, page.bottom) })).filter(slot => slot.bottom > slot.top);
        if (segments.length < 2) continue;
        const lines: Array<{ top: number; bottom: number; chars: Array<{ pos: number; end: number }> }> = [];
        for (const para of paragraphs) {
          const walker = document.createTreeWalker(para, NodeFilter.SHOW_TEXT); let text: Node | null;
          while ((text = walker.nextNode())) {
            if (!text.textContent?.trim() || text.parentElement?.closest('.page-gap, .mo-word-cell-gap, .page-repeat-header')) continue;
            const pos = view.posAtDOM(text, 0);
            for (let offset = 0; offset < text.textContent.length; offset++) {
              range.setStart(text, offset); range.setEnd(text, offset + 1);
              const rect = range.getClientRects()[0]; if (!rect || rect.width < 0.1 || rect.height < 0.1) continue;
              const top = (rect.top - origin) / factor, bottom = (rect.bottom - origin) / factor;
              let line = lines.find(line => Math.abs(line.top - top) < 1);
              if (!line) { line = { top, bottom, chars: [] }; lines.push(line); }
              line.bottom = Math.max(line.bottom, bottom); line.chars.push({ pos: pos + offset, end: pos + offset + 1 });
            }
          }
        }
        lines.sort((a, b) => a.top - b.top);
        const shifts = wordCellLineShifts(lines, segments);
        if (shifts) lines.forEach((line, index) => {
          if (Math.abs(shifts[index]) < 0.5) return;
          // 功能：相邻同位移字符合并为一个装饰，保留字体/域/原生选区与保存内容；时间：2026-11-11；作者：lq。
          for (const char of line.chars.sort((a, b) => a.pos - b.pos)) {
            const last = items[items.length - 1];
            if (last?.end === char.pos && last.shift === shifts[index]) last.end = char.end;
            else items.push({ ...char, shift: shifts[index] });
          }
        });
        continue;
      }
      if (Math.abs(shift) < 0.5) continue;
      for (const para of paragraphs) {
        const position = view.state.doc.resolve(view.posAtDOM(para, 0));
        for (let depth = position.depth; depth > 0; depth--) if (position.node(depth).isTextblock) {
          if (view.nodeDOM(position.before(depth)) !== para) break;
          // 功能：仅移动文字运行，不能平移段落里的页缝 widget，否则会反向改变整篇物理页界；时间：2026-11-11；作者：lq。
          position.node(depth).forEach((node, offset) => {
            if (node.isText && node.text?.trim()) items.push({ pos: position.start(depth) + offset, end: position.start(depth) + offset + node.nodeSize, shift });
          });
          break;
        }
      }
    }
  } finally { root.classList.remove('mo-measuring-cell-alignment'); }
  const signature = items.map(item => `${item.pos}:${item.shift.toFixed(2)}`).join('|'), old = alignKey.getState(view.state);
  if (old?.find()[0]?.spec.signature === signature || (!items.length && !old?.find().length)) return;
  const decorations = items.map(item => Decoration.inline(item.pos, item.end, { class: 'mo-word-merged-cell-text', style: `--mo-cell-page-shift:${item.shift.toFixed(2)}px` }, { signature }));
  batch.set(alignKey, DecorationSet.create(view.state.doc, decorations));
}
