/** 功能：提供定制 Word 的编辑、模型转换和选区入口；时间：2026-11-08；作者：lq。 */
export { editorExtensions } from '../upstream/apps/docs/src/renderer/editor/extensions';
export { blocksToPmDoc, pmDocOptions, pmDocToSavePlan, type SavePlan } from '../upstream/apps/docs/src/renderer/editor/convert';
export { getActiveSubEditor, setActiveSubEditor, dropActiveSubEditor } from '../upstream/apps/docs/src/renderer/editor/active-editor';
export { Editor } from '@tiptap/core';
