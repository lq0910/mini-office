/** 功能：提供定制 Word 的编辑、模型转换和选区入口；时间：2026-11-08；作者：lq。 */
export { editorExtensions } from '../upstream/apps/docs/src/renderer/editor/extensions';
export { blocksToPmDoc, pmDocOptions, pmDocToSavePlan, type SavePlan } from '../upstream/apps/docs/src/renderer/editor/convert';
export { getActiveSubEditor, setActiveSubEditor, dropActiveSubEditor } from '../upstream/apps/docs/src/renderer/editor/active-editor';
export { Editor } from '@tiptap/core';
// 功能：表格 UI 使用原生矩形选区，不伪造跨单元格文本选区；时间：2026-11-10；作者：lq。
export { CellSelection } from '@tiptap/pm/tables';
// 功能：公开原生格式摘要和局部事务，供现有工具栏直接调用；时间：2026-11-10；作者：lq。
export { readWordFormatting, applyWordTextFormat, applyWordParagraphFormat, toggleWordInlineFormat, type WordTextFormat, type WordParagraphFormat } from './formatting';
export { insertPageBreak } from '../upstream/apps/docs/src/renderer/editor/page-break';
// 功能：公开列表状态及同源文字查找/替换操作，供现有 UI 调用；时间：2026-11-10；作者：lq。
export { readWordList, changeWordListLevel } from './lists';
export { findWordText, highlightWordSearch, replaceWordText, type WordSearchResults } from './search';
// 功能：公开样式目录及格式刷事务，工具栏不再依赖旧正文模型；时间：2026-11-10；作者：lq。
export { captureWordFormat, applyWordFormatPainter, type WordFormatSnapshot } from './format-painter';
export { wordStyles, wordStyleFormat, defineWordStyle, applyWordStyle, importWordStyles, renameWordStyle, deleteWordStyle } from './styles';
// 功能：标尺和目录使用原生选区/事务及同源分页缓存，复用现有产品入口；时间：2026-11-11；作者：lq。
export { applyWordRuler, wordRulerTabs } from './ruler';
export { insertWordToc, followWordToc, wordHeadings, wordTocSnapshot } from './toc';
