/** 功能：格内段落来源锚点随原生节点历史保存，拆段及新建格不继承旧书签；时间：2026-11-11；作者：lq。 */
import { Extension } from '@tiptap/core';
export const wordTableSource = Extension.create({
  name: 'miniOfficeTableSource',
  addGlobalAttributes() {
    return [
      { types: ['docParagraph', 'docListItem'], attributes: { miniOfficeCellSource: { default: null, rendered: false, keepOnSplit: false } } },
      // 功能：续格跟随原行与合并起格身份，插删其他行列后仍能找回隐藏内容及原始属性；时间：2026-11-11；作者：lq。
      { types: ['docTableRow'], attributes: { miniOfficeSourceRow: { default: null, rendered: false, keepOnSplit: false } } },
      { types: ['docTableCell', 'docTableHeader'], attributes: { miniOfficeSourceColumn: { default: null, rendered: false, keepOnSplit: false } } },
    ];
  },
});
