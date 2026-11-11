/** 功能：图片和表格入口直接复用同源模型、结构命令与历史；时间：2026-11-11；作者：lq。 */
import type { Editor } from '@tiptap/core';
import type { Node as PmNode } from '@tiptap/pm/model';
import { NodeSelection, TextSelection, type Command } from '@tiptap/pm/state';
import { closeHistory } from '@tiptap/pm/history';
import { CellSelection, TableMap, deleteRow, deleteTable, isInTable, mergeCells, selectedRect, splitCell } from '@tiptap/pm/tables';
import { insertRowsOrColumns, deleteTableColumn, selectTablePart, tableCellsSelection, distributeRowsEvenly, setCellProps, setRowProps } from '../upstream/apps/docs/src/renderer/editor/table-ops';
import { distributeSelectedColumns, setSelectedColumnWidth, setTableGridAtCell, tableGridAtCell } from '../upstream/apps/docs/src/renderer/editor/table-sizing';
import { setTableAutoFit, toggleRepeatHeaderRows, repeatHeaderState } from '../upstream/apps/docs/src/renderer/editor/table-properties';
import { setSelectionBorders, type TableBorderMode } from '../upstream/apps/docs/src/renderer/editor/table-borders';
import { wordTableCell, wordTableModel } from '../upstream/apps/docs/src/renderer/editor/new-table';
import { tableModelToPmNode } from '../upstream/apps/docs/src/renderer/editor/convert';
import { insertPictureAtSelection, type NewPicture } from '../upstream/apps/docs/src/renderer/editor/insert-picture';
import type { TableCell } from '../upstream/packages/docx-engine/src/types';

/** 功能：跨浮层的表格和图片命令只在可编辑且非输入法组合状态执行；时间：2026-11-11；作者：lq。 */
const editable = (editor: Editor) => !editor.isDestroyed && editor.isEditable && !editor.view.composing;
/** 功能：独立工具栏操作分隔连续输入，纯选择操作不进入历史；时间：2026-11-11；作者：lq。 */
function dispatchCommand(editor: Editor, command: Command, change = true): boolean {
  if (!editable(editor)) return false;
  const done = command(editor.state, tr => editor.view.dispatch(change ? closeHistory(tr) : tr), editor.view);
  if (done && change) editor.view.dispatch(closeHistory(editor.state.tr));
  return done;
}
/** 功能：整表节点选区转换为原生矩形选区，单元格命令保持整表作用范围；时间：2026-11-11；作者：lq。 */
function enterTable(editor: Editor): boolean {
  const cells = tableCellsSelection(editor.state);
  if (cells) editor.view.dispatch(editor.state.tr.setSelection(cells));
  return isInTable(editor.state);
}
/** 功能：读取当前节真实可用宽度，未挂载时按调用者提供的节尺寸计算；时间：2026-11-11；作者：lq。 */
export function wordContentWidth(editor: Editor, fallback = 624): number {
  const pos = editor.state.selection.$from.depth ? editor.state.selection.$from.before(1) : editor.state.selection.from;
  const element = editor.view.nodeDOM(pos);
  const width = Number.parseFloat(getComputedStyle(element instanceof HTMLElement ? element : editor.view.dom).getPropertyValue('--section-content-w'));
  return Number.isFinite(width) && width > 0 ? width : fallback;
}
/** 功能：插入有原生网格和边框的表格，空段落不吞掉，光标进入首格；时间：2026-11-11；作者：lq。 */
export function insertWordTable(editor: Editor, rows: number, columns: number): boolean {
  if (!editable(editor) || !editor.schema.nodes.docTable || !Number.isInteger(rows) || !Number.isInteger(columns) || rows < 1 || rows > 100 || columns < 1 || columns > 63) return false;
  const model = wordTableModel(Array.from({ length: rows }, () => Array.from({ length: columns }, () => wordTableCell(['']))), columns);
  const { $from } = editor.state.selection;
  for (let depth = $from.depth; depth > 0; depth--) {
    if (['docTableCell', 'docTableHeader'].includes($from.node(depth).type.name)) {
      const done = editor.chain().command(({ tr }) => { closeHistory(tr); return true; })
        .insertContentAt($from.end(depth), [{ type: 'docNestedTable', attrs: { model } }, { type: 'docParagraph' }]).run();
      editor.view.dispatch(closeHistory(editor.state.tr)); return done;
    }
  }
  const node = tableModelToPmNode(model);
  const block = $from.depth ? $from.node(1) : null;
  const at = block?.isTextblock && !block.content.size ? $from.before(1) : null;
  const chain = editor.chain().command(({ tr }) => { closeHistory(tr); return true; });
  const done = at == null ? chain.insertContent(node).run()
    : chain.insertContentAt(at, node).command(({ tr }) => { tr.setSelection(TextSelection.near(tr.doc.resolve(at + 1))); return true; }).run();
  // 功能：插入结束后隔开随后的连续输入，不能一次撤销同时吞掉表格与首格文字；时间：2026-11-11；作者：lq。
  editor.view.dispatch(closeHistory(editor.state.tr)); return done;
}

/** 功能：图片尺寸与数据校验后沿用同源插入，保留尾段和单步历史；时间：2026-11-11；作者：lq。 */
export function insertWordPicture(editor: Editor, picture: NewPicture): boolean {
  // 功能：页眉页脚与文本框采用原生行内图片节点，仍使用相同文件校验和单步历史；时间：2026-11-11；作者：lq。
  if (!editor.schema.nodes.docProtected && editor.schema.nodes.docInlineImage) return insertWordPictures(editor, [picture]);
  if (!editable(editor) || !editor.schema.nodes.docProtected || !/^image\/(png|jpeg|gif|webp|svg\+xml)$/.test(picture.mime)
    || !picture.base64 || !picture.dataUrl.startsWith(`data:${picture.mime};base64,`)
    || ![picture.widthPx, picture.heightPx].every(value => Number.isFinite(value) && value >= 1 && value <= 30000)) return false;
  editor.view.dispatch(closeHistory(editor.state.tr));
  if (!insertPictureAtSelection(editor, picture)) return false;
  const { doc, selection, schema } = editor.state;
  if (selection instanceof NodeSelection && selection.node.attrs.blockType === 'image') {
    const tr = editor.state.tr;
    const after = selection.to;
    if (after === doc.content.size) tr.insert(after, schema.nodes.docParagraph.create());
    tr.setSelection(TextSelection.near(tr.doc.resolve(after), 1));
    editor.view.dispatch(tr);
  }
  editor.view.dispatch(closeHistory(editor.state.tr));
  return true;
}

/** 功能：多张剪贴板图片一次插入和撤销，沿用正文图片/格内 run 的同源属性与宽度规则；时间：2026-11-11；作者：lq。 */
export function insertWordPictures(editor: Editor, pictures: NewPicture[]): boolean {
  if (!editable(editor) || !pictures.length || pictures.length > 32 || (!editor.schema.nodes.docProtected && !editor.schema.nodes.docInlineImage)) return false;
  if (pictures.some(picture => !/^image\/(png|jpeg|gif|webp|svg\+xml)$/.test(picture.mime) || !picture.base64 || !picture.dataUrl.startsWith(`data:${picture.mime};base64,`) || ![picture.widthPx, picture.heightPx].every(value => Number.isFinite(value) && value >= 1 && value <= 30000))) return false;
  let available = editor.schema.nodes.docProtected ? wordContentWidth(editor) : editor.view.dom.clientWidth || 624; let insideCell = !editor.schema.nodes.docProtected;
  const { $from } = editor.state.selection;
  for (let depth = $from.depth; depth > 0; depth--) if (['docTableCell', 'docTableHeader', 'hfCell'].includes($from.node(depth).type.name)) {
    insideCell = true;
    const dom = editor.view.nodeDOM($from.before(depth));
    if (dom instanceof HTMLElement && dom.clientWidth) {
      const css = getComputedStyle(dom);
      available = dom.clientWidth - (Number.parseFloat(css.paddingLeft) || 0) - (Number.parseFloat(css.paddingRight) || 0);
    }
    break;
  }
  const nodes = pictures.map(picture => {
    const scale = Math.min(1, available / picture.widthPx);
    const widthPx = Math.max(1, Math.round(picture.widthPx * scale)), heightPx = Math.max(1, Math.round(picture.heightPx * scale));
    return insideCell ? { type: 'docInlineImage', attrs: { dataUrl: picture.dataUrl, widthPx, heightPx } }
      : { type: 'docProtected', attrs: { docxIndex: null, blockType: 'image', label: picture.label, imageDataUrl: picture.dataUrl, imageWidthPx: widthPx, imageHeightPx: heightPx, genImage: { base64: picture.base64, mime: picture.mime, widthPx, heightPx } } };
  });
  const done = editor.chain().command(({ tr }) => { closeHistory(tr); return true; }).insertContent(nodes).command(({ tr }) => {
    if (tr.doc.lastChild?.type.name === 'docProtected') tr.insert(tr.doc.content.size, editor.schema.nodes.docParagraph.create());
    if (tr.selection instanceof NodeSelection && tr.selection.node.attrs.blockType === 'image') tr.setSelection(TextSelection.near(tr.doc.resolve(tr.selection.to), 1));
    tr.setMeta('uiEvent', 'paste');
    return true;
  }).run();
  editor.view.dispatch(closeHistory(editor.state.tr)); return done;
}

export type WordTableAction = 'row-above' | 'row-below' | 'column-before' | 'column-after' | 'row-delete' | 'column-delete' | 'table-delete'
  | 'merge' | 'split' | 'select-rows' | 'select-columns' | 'select-table' | 'distribute-columns' | 'distribute-rows' | 'fit-page' | 'header';
/** 功能：右键菜单摘要基于实际矩形选区，不把数组单元格索引当网格列；时间：2026-11-11；作者：lq。 */
export function wordTableState(editor: Editor) {
  const cells = tableCellsSelection(editor.state);
  const state = cells ? editor.state.apply(editor.state.tr.setSelection(cells)) : editor.state;
  if (!isInTable(state)) return undefined;
  const rect = selectedRect(state);
  const cell = state.doc.nodeAt(rect.tableStart + rect.map.map[rect.top * rect.map.width + rect.left])!;
  return { table: rect.table, cell, rows: rect.map.height, columns: rect.map.width,
    canMerge: mergeCells(state), canSplit: splitCell(state), header: repeatHeaderState(state),
    row: rect.table.child(rect.top), rect };
}
/** 功能：拆分导入的纵向合并格时，将隐藏续格恢复为可编辑原生内容和属性，整次拆分只有一步历史；时间：2026-11-11；作者：lq。 */
const splitWordCell: Command = (state, dispatch) => {
  if (!isInTable(state)) return false;
  const rect = selectedRect(state);
  const first = rect.table.nodeAt(rect.map.map[rect.top * rect.map.width + rect.left]);
  const column = first?.attrs.miniOfficeSourceColumn;
  let records: { row: number; column: number; cell: TableCell }[] = [];
  try { records = JSON.parse(rect.table.attrs.continuationCells ?? '[]'); } catch { /* 功能：无来源续格时保持原生命令；时间：2026-11-11；作者：lq。 */ }
  return splitCell(state, dispatch ? tr => {
    if (column != null && first && first.attrs.rowspan > 1) for (let row = rect.top + 1; row < rect.bottom; row++) {
      const sourceRow = rect.table.child(row).attrs.miniOfficeSourceRow;
      const original = records.find(record => record.row === sourceRow && record.column === column);
      if (!original) continue;
      const table = tr.doc.nodeAt(rect.tableStart - 1)!;
      const map = TableMap.get(table);
      const at = rect.tableStart + map.map[row * map.width + rect.left];
      const current = tr.doc.nodeAt(at)!;
      const restored = state.schema.nodeFromJSON(tableModelToPmNode({ rows: [[{ ...original.cell, vMerge: undefined, colSpan: 1 }]] })).firstChild!.firstChild!;
      tr.setNodeMarkup(at, undefined, { ...restored.attrs, colwidth: current.attrs.colwidth, colspan: 1, rowspan: 1, clipHeightTwips: null, miniOfficeSourceColumn: column });
      tr.replaceWith(at + 1, at + current.nodeSize - 1, restored.content);
    }
    dispatch(tr);
  } : undefined);
};
/** 功能：整表选择、行列增删、合并拆分和布局全部进入唯一核心事务；时间：2026-11-11；作者：lq。 */
export function runWordTableAction(editor: Editor, action: WordTableAction): boolean {
  if (!editable(editor) || !enterTable(editor)) return false;
  if (action === 'distribute-rows') {
    editor.view.dispatch(closeHistory(editor.state.tr));
    return distributeRowsEvenly(editor.view);
  }
  const width = wordContentWidth(editor);
  const rect = selectedRect(editor.state);
  const commands: Record<Exclude<WordTableAction, 'distribute-rows'>, Command> = {
    'row-above': insertRowsOrColumns(1, 'above', width), 'row-below': insertRowsOrColumns(1, 'below', width),
    'column-before': insertRowsOrColumns(1, 'left', width), 'column-after': insertRowsOrColumns(1, 'right', width),
    'row-delete': deleteRow, 'column-delete': rect.right - rect.left === rect.map.width ? deleteTable : deleteTableColumn, 'table-delete': deleteTable,
    merge: mergeCells, split: splitWordCell, 'select-rows': selectTablePart('row'), 'select-columns': selectTablePart('column'),
    'select-table': selectTablePart('table'), 'distribute-columns': distributeSelectedColumns(width),
    'fit-page': setTableAutoFit('window', width), header: toggleRepeatHeaderRows(),
  };
  return dispatchCommand(editor, commands[action], !action.startsWith('select-'));
}

/** 功能：表格属性只提交用户修改字段，原始合并格、行属性和未知 XML 继续保留；时间：2026-11-11；作者：lq。 */
export interface WordTablePatch {
  widthPx?: number; columnWidthPx?: number; rowHeightTwips?: number; rowHeightRule?: 'atLeast' | 'exact';
  align?: 'left' | 'center' | 'right'; fill?: string | null; vAlign?: 'top' | 'center' | 'bottom';
  border?: { mode: TableBorderMode; color: string; widthPt: number; style: 'single' | 'double' | 'dashed' | 'dotted' };
}
/** 功能：把尺寸、行高、底纹及边框组合为一次属性编辑和一次撤销；时间：2026-11-11；作者：lq。 */
export function applyWordTableProperties(editor: Editor, patch: WordTablePatch): boolean {
  if (!editable(editor) || !Object.keys(patch).length) return false;
  if ([patch.widthPx, patch.columnWidthPx, patch.rowHeightTwips].some(value => value !== undefined && (!Number.isFinite(value) || value < 1 || value > 31680))) return false;
  if (patch.fill && !/^[a-f\d]{6}$/i.test(patch.fill)) return false;
  if (patch.border && (!/^[a-f\d]{6}$/i.test(patch.border.color) || !Number.isFinite(patch.border.widthPt) || patch.border.widthPt < 0.25 || patch.border.widthPt > 12)) return false;
  if (!enterTable(editor)) return false;
  const initial = editor.state;
  let state = initial;
  let combined = closeHistory(initial.tr);
  const collect = (command: Command) => command(state, transaction => {
    transaction.steps.forEach(step => combined.step(step));
    state = state.apply(transaction);
  }, editor.view);
  const width = wordContentWidth(editor);
  if (patch.columnWidthPx !== undefined) collect(setSelectedColumnWidth(patch.columnWidthPx, width));
  if (patch.rowHeightTwips !== undefined || patch.rowHeightRule !== undefined) collect(setRowProps({ ...(patch.rowHeightTwips !== undefined ? { heightTwips: Math.round(patch.rowHeightTwips) } : {}), ...(patch.rowHeightRule ? { heightRule: patch.rowHeightRule } : {}) }));
  if ('fill' in patch || patch.vAlign) collect(setCellProps({ ...('fill' in patch ? { fill: patch.fill } : {}), ...(patch.vAlign ? { vAlign: patch.vAlign } : {}) }));
  if (patch.border) collect(setSelectionBorders(patch.border.mode, { style: patch.border.style, color: patch.border.color, szEighths: Math.round(patch.border.widthPt * 8) }));
  // 功能：表宽必须同时更新网格及每个合并格的宽度，防止保存后回到旧网格总宽；时间：2026-11-11；作者：lq。
  if (patch.widthPx !== undefined) {
    const rect = selectedRect(state);
    const first = rect.tableStart + rect.map.map[0];
    const grid = tableGridAtCell(state, first, width);
    if (!grid) return false;
    const total = grid.widths.reduce((sum, value) => sum + value, 0);
    collect(setTableGridAtCell(first, grid.widths.map(value => value * patch.widthPx! / total)));
  }
  const rect = selectedRect(state);
  const tableAttrs: Record<string, unknown> = {};
  if (patch.align) tableAttrs.tblAlign = patch.align;
  if (patch.widthPx !== undefined) Object.assign(tableAttrs, { widthPx: patch.widthPx, widthPct: null, tblAutoFit: 'fixed', tblAutoFitEdited: true });
  if (Object.keys(tableAttrs).length) combined.setNodeMarkup(rect.tableStart - 1, undefined, { ...rect.table.attrs, ...tableAttrs });
  if (!combined.docChanged) return false;
  editor.view.dispatch(combined);
  return true;
}

/** 功能：右键定位原生格内或图片选区，选中矩形内右键不破坏已有矩形；时间：2026-11-11；作者：lq。 */
export function selectWordObjectAt(editor: Editor, target: HTMLElement): 'table' | 'image' | undefined {
  if (!editable(editor) || !editor.view.dom.contains(target)) return undefined;
  let image: { node: PmNode; pos: number } | undefined;
  editor.state.doc.descendants((node, pos) => {
    if (image) return false;
    if (node.type.name !== 'docInlineImage' && !(node.type.name === 'docProtected' && node.attrs.blockType === 'image')) return;
    const dom = editor.view.nodeDOM(pos);
    if (dom === target || dom instanceof Element && dom.contains(target)) image = { node, pos };
  });
  if (image) { editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, image.pos))); return 'image'; }
  const td = target.closest('td, th');
  if (!td || !editor.view.dom.contains(td)) return undefined;
  // 功能：整表手柄选中后在内部右键仍作用于全表，不缩小为单格；时间：2026-11-11；作者：lq。
  if (editor.state.selection instanceof NodeSelection && editor.state.selection.node.type.name === 'docTable') {
    const dom = editor.view.nodeDOM(editor.state.selection.from);
    if (dom instanceof Element && dom.contains(td)) return 'table';
  }
  const inside = editor.view.posAtDOM(td, 0);
  const selection = editor.state.selection;
  let selected = false;
  if (selection instanceof CellSelection) selection.forEachCell((_node, pos) => { if (pos + 1 === inside) selected = true; });
  if (!selected) editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(inside))));
  return isInTable(editor.state) ? 'table' : undefined;
}
/** 功能：图片环绕和尺寸保留来源节点所有非目标绘图属性；时间：2026-11-11；作者：lq。 */
export function applyWordImage(editor: Editor, patch: { wrap?: 'inline' | 'square-left' | 'square-right' | 'topBottom'; widthPx?: number; heightPx?: number; delete?: boolean }): boolean {
  if (!editable(editor)) return false;
  const selection = editor.state.selection;
  if (!(selection instanceof NodeSelection)) return false;
  const node = selection.node;
  const inline = node.type.name === 'docInlineImage';
  if (!inline && !(node.type.name === 'docProtected' && node.attrs.blockType === 'image')) return false;
  if ([patch.widthPx, patch.heightPx].some(value => value !== undefined && (!Number.isFinite(value) || value < 1 || value > 30000))) return false;
  const tr = closeHistory(editor.state.tr);
  if (patch.delete) tr.deleteSelection();
  else {
    const attrs = { ...node.attrs };
    if (patch.wrap !== undefined) {
      attrs[inline ? 'wrap' : 'imageWrap'] = patch.wrap === 'inline' ? null : patch.wrap;
      // 功能：从浮动位置切回预设环绕时清除旧偏移，其他绘图效果保留；时间：2026-11-11；作者：lq。
      for (const key of inline ? ['offsetXEmu', 'offsetYEmu', 'posH', 'posV'] : ['imageOffsetXEmu', 'imageOffsetYEmu', 'imagePosH', 'imagePosV']) attrs[key] = null;
    }
    if (patch.widthPx !== undefined) attrs[inline ? 'widthPx' : 'imageWidthPx'] = patch.widthPx;
    if (patch.heightPx !== undefined) attrs[inline ? 'heightPx' : 'imageHeightPx'] = patch.heightPx;
    tr.setNodeMarkup(selection.from, undefined, attrs);
  }
  if (tr.doc.eq(editor.state.doc)) return false;
  editor.view.dispatch(tr); return true;
}
