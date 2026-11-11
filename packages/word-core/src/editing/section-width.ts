/** 功能：以原生装饰提供各分节正文净宽，布局代码不能直接改写可编辑段落 DOM；时间：2026-11-11；作者：lq。 */
import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import type { ParsedDocFull } from '../upstream/packages/docx-engine/src/index';
import { wordSections } from './page-layout';

export function wordSectionWidths(parsed: ParsedDocFull) {
  return Extension.create({
    name: 'miniOfficeSectionWidths',
    addProseMirrorPlugins: () => [new Plugin({ props: { decorations(state) {
      const sections = wordSections(state.doc, parsed);
      const decorations: Decoration[] = [];
      let sectionIndex = 0;
      state.doc.forEach((node, pos) => {
        const section = sections[Math.min(sectionIndex, sections.length - 1)];
        if (!section) return;
        const s = section.settings;
        decorations.push(Decoration.node(pos, pos + node.nodeSize, { style: `--section-content-w:${(s.pageWidth - s.marginLeft - s.marginRight) / 15}px` }));
        if (node.attrs.docxIndex === section.lastBlockIndex) sectionIndex++;
      });
      return DecorationSet.create(state.doc, decorations);
    } } })],
  });
}
