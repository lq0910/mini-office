/** 功能：从唯一原生正文计算字数及字符数，统计面板和状态栏采用相同口径；时间：2026-11-11；作者：lq。 */
import type { Node } from '@tiptap/pm/model';
import type { TextboxDisplay } from '../upstream/packages/docx-engine/src/index';

export function wordDocumentStats(doc: Node): { words: number; characters: number; paragraphs: number } {
  const texts: string[] = [];
  doc.descendants(node => {
    if (node.isTextblock) {
      texts.push(node.textBetween(0, node.content.size, '', leaf => leaf.type.name === 'hardBreak' ? '\n' : typeof leaf.attrs.text === 'string' ? leaf.attrs.text : ''));
      return false;
    }
    // 功能：文本框内容由原生对象属性持有；不统计编辑器装饰、页间占位或隐藏续格文字；时间：2026-11-11；作者：lq。
    const boxes = (node.attrs.textboxes ?? node.attrs.boxes) as TextboxDisplay[] | null;
    for (const box of boxes ?? []) for (const paragraph of box.paras ?? []) texts.push(paragraph.runs.map(run => run.image ? '' : run.text).join(''));
  });
  let words = 0; let characters = 0;
  for (const text of texts) {
    words += (text.match(/\p{Script=Han}/gu)?.length ?? 0) + text.replace(/\p{Script=Han}/gu, ' ').trim().split(/\s+/u).filter(Boolean).length;
    characters += Array.from(text.replace(/\s/gu, '')).length;
  }
  return { words, characters, paragraphs: texts.length };
}
