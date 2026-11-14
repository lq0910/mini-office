/** 功能：以原生装饰提供各分节正文净宽，布局代码不能直接改写可编辑段落 DOM；时间：2026-11-11；作者：lq。 */
import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import type { Node } from '@tiptap/pm/model';
import type { ParsedDocFull } from '../upstream/packages/docx-engine/src/index';
import { wordSections } from './page-layout';

export function wordSectionWidths(parsed: ParsedDocFull) {
  // 功能：同一不可变正文的装饰只构建一次，选区/分页装饰事务复用；时间：2026-11-13；作者：lq。
  const cache = new WeakMap<Node, DecorationSet>();
  return Extension.create({
    name: 'miniOfficeSectionWidths',
    addProseMirrorPlugins: () => [new Plugin({ props: {
      // 功能：统一宽度继承根样式，与 GenOffice 的单流画布一致，避免数千个相同节点装饰反复创建；时间：2026-11-13；作者：lq。
      attributes(state): Record<string, string> {
        const settings = wordSections(state.doc, parsed)[0]?.settings;
        return settings ? { style: `--section-content-w:${(settings.pageWidth - settings.marginLeft - settings.marginRight) / 15}px` } : {};
      },
      decorations(state) {
      const previous = cache.get(state.doc);
      if (previous) return previous;
      const sections = wordSections(state.doc, parsed);
      const widths = sections.map(section => (section.settings.pageWidth - section.settings.marginLeft - section.settings.marginRight) / 15);
      if (widths.every(width => width === widths[0])) { cache.set(state.doc, DecorationSet.empty); return DecorationSet.empty; }
      const decorations: Decoration[] = [];
      let sectionIndex = 0;
      state.doc.forEach((node, pos) => {
        const section = sections[Math.min(sectionIndex, sections.length - 1)];
        if (!section) return;
        const s = section.settings;
        decorations.push(Decoration.node(pos, pos + node.nodeSize, { style: `--section-content-w:${(s.pageWidth - s.marginLeft - s.marginRight) / 15}px` }));
        if (node.attrs.docxIndex === section.lastBlockIndex) sectionIndex++;
      });
      const result = DecorationSet.create(state.doc, decorations);
      cache.set(state.doc, result);
      return result;
    } } })],
  });
}
