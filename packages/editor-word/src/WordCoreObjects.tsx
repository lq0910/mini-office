/** 功能：复用图片文件选择、原右键菜单和属性控件，操作唯一 Word 核心；时间：2026-11-11；作者：lq。 */
import { useEffect, useRef, useState, type MouseEvent } from 'react';
import { Button, ContextMenu, Dialog, Field, Input, Select, type MenuItem } from '@mini-office/ui';
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import type { Editor } from '../../word-core/src/editing';
import { applyWordImage, applyWordTableProperties, insertWordPicture, runWordTableAction, selectWordObjectAt, wordContentWidth, wordTableState, type WordTableAction, type WordTablePatch } from '../../word-core/src/editing/objects';
import { readWordPictureFile } from '../../word-core/src/editing/picture-files';

/** 功能：文件异步加载和属性草稿带文档身份，拒绝覆盖期间已改变的正文；时间：2026-11-11；作者：lq。 */
type ObjectTarget = { editor: Editor; doc: Editor['state']['doc']; bookmark: ReturnType<Editor['state']['selection']['getBookmark']> };
const capture = (editor: Editor): ObjectTarget => ({ editor, doc: editor.state.doc, bookmark: editor.state.selection.getBookmark() });
/** 功能：浮层关闭后还原原生选区，不用 DOM 偏移反推文档位置；时间：2026-11-11；作者：lq。 */
function restore(target: ObjectTarget): boolean {
  if (target.editor.isDestroyed || !target.editor.isEditable || target.editor.view.composing || target.editor.state.doc !== target.doc) return false;
  target.editor.view.dispatch(target.editor.state.tr.setSelection(target.bookmark.resolve(target.doc)));
  return true;
}
/** 功能：编辑区右键和插入图片共用目标校验及错误入口；时间：2026-11-11；作者：lq。 */
export function useWordCoreObjects(core: WordCoreSession, readonly: boolean, onError: (message: string | undefined) => void) {
  const input = useRef<HTMLInputElement>(null);
  const pending = useRef<ObjectTarget>();
  const mounted = useRef(true);
  const [menu, setMenu] = useState<{ x: number; y: number; kind: 'table' | 'image'; target: ObjectTarget }>();
  const [properties, setProperties] = useState<ObjectTarget>();
  useEffect(() => { mounted.current = true; return () => { mounted.current = false; pending.current = undefined; }; }, [core]);
  /** 功能：菜单执行前校验捕获版本，取消或延迟文件读取不改变历史；时间：2026-11-11；作者：lq。 */
  const execute = (command: (editor: Editor) => unknown) => {
    if (!menu || readonly) return;
    if (!restore(menu.target)) { onError('文档已变化，请重新选择对象。'); return; }
    try { command(menu.target.editor); menu.target.editor.view.focus(); onError(undefined); }
    catch (error) { onError(error instanceof Error ? error.message : String(error)); }
  };
  const table = menu?.kind === 'table' && !menu.target.editor.isDestroyed ? wordTableState(menu.target.editor) : undefined;
  const action = (id: WordTableAction, label: string, disabled = false, danger = false): MenuItem => ({ id, label, disabled, danger, onSelect: () => execute(editor => runWordTableAction(editor, id)) });
  const tableItems: MenuItem[] = [
    { type: 'label', label: '插入' }, action('row-above', '在上方插入行'), action('row-below', '在下方插入行'),
    action('column-before', '在左侧插入列'), action('column-after', '在右侧插入列'), { type: 'separator' },
    action('row-delete', '删除行'), action('column-delete', '删除列'), action('table-delete', '删除表格', false, true), { type: 'separator' },
    action('merge', '合并单元格', !table?.canMerge), action('split', '拆分单元格', !table?.canSplit), { type: 'separator' },
    action('select-rows', '选择行'), action('select-columns', '选择列'), action('select-table', '选择整个表格'), { type: 'separator' },
    action('distribute-columns', '平均分布各列'), action('distribute-rows', '平均分布各行'), action('fit-page', '适应页宽'),
    { ...action('header', '重复标题行', !table?.header.enabled), checked: table?.header.active }, { type: 'separator' },
    { id: 'properties', label: '表格属性…', onSelect: () => { if (menu) setProperties(menu.target); } },
  ];
  const imageNode = menu?.kind === 'image' ? menu.target.editor.state.doc.nodeAt(menu.target.editor.state.selection.from) : undefined;
  const currentWrap = imageNode?.attrs[imageNode.type.name === 'docInlineImage' ? 'wrap' : 'imageWrap'] ?? 'inline';
  const imageItems: MenuItem[] = [
    { type: 'label', label: '图片文字环绕' },
    ...(['inline', 'square-left', 'square-right', 'topBottom'] as const).map((wrap, index) => ({ id: wrap, label: ['嵌入型', '左侧环绕', '右侧环绕', '上下型'][index], checked: currentWrap === wrap, onSelect: () => execute(editor => applyWordImage(editor, { wrap })) })),
    { type: 'separator' }, { id: 'image-delete', label: '删除图片', danger: true, onSelect: () => execute(editor => applyWordImage(editor, { delete: true })) },
  ];
  return {
    // 功能：在浏览器右键默认行为前固定矩形选区，避免 contextmenu 到来时选区已缩成光标；时间：2026-11-11；作者：lq。
    onMouseDown: (event: MouseEvent) => {
      if (event.button !== 2 || readonly) return;
      if (selectWordObjectAt(core.commandEditor(), event.target as HTMLElement)) { event.preventDefault(); event.stopPropagation(); }
    },
    /** 功能：只在已迁移的活动 schema 开放图片文件选择；时间：2026-11-11；作者：lq。 */
    chooseImage: (editor: Editor) => { if (!readonly && (editor.schema.nodes.docProtected || editor.schema.nodes.docInlineImage)) { pending.current = capture(editor); input.current?.click(); } },
    onContextMenu: (event: MouseEvent) => {
      if (readonly) return;
      const editor = core.commandEditor();
      const kind = selectWordObjectAt(editor, event.target as HTMLElement);
      if (!kind) return;
      event.preventDefault(); event.stopPropagation();
      setMenu({ x: event.clientX, y: event.clientY, kind, target: capture(editor) });
    },
    elements: <>
      <input ref={input} type="file" aria-label="插入图片文件" accept="image/png,image/jpeg,image/gif,image/webp,image/svg+xml" className="hidden" onChange={event => {
        const file = event.target.files?.[0]; const target = pending.current; pending.current = undefined; event.target.value = '';
        if (!file || !target) return;
        void readWordPictureFile(file).then(picture => {
          if (!mounted.current) return;
          if (!restore(target)) throw new Error('图片加载期间文档已变化，请重新插入。');
          const scale = Math.min(1, wordContentWidth(target.editor) / picture.widthPx);
          if (!insertWordPicture(target.editor, { ...picture, widthPx: Math.max(1, Math.round(picture.widthPx * scale)), heightPx: Math.max(1, Math.round(picture.heightPx * scale)) })) throw new Error('当前编辑区域不能插入图片。');
          target.editor.view.focus(); onError(undefined);
        }).catch(error => { if (mounted.current) onError(error instanceof Error ? error.message : String(error)); });
      }} />
      {menu ? <ContextMenu x={menu.x} y={menu.y} items={menu.kind === 'table' ? tableItems : imageItems} onClose={() => setMenu(undefined)} /> : null}
      {properties ? <WordCoreTableProperties target={properties} onClose={() => setProperties(undefined)} onError={onError} /> : null}
    </>,
  };
}

/** 功能：沿用原表格属性字段，长度使用 px/twips 局部换算，一次应用可撤销；时间：2026-11-11；作者：lq。 */
function WordCoreTableProperties({ target, onClose, onError }: { target: ObjectTarget; onClose: () => void; onError: (message: string | undefined) => void }) {
  const [patch, setPatch] = useState<WordTablePatch>({});
  const [error, setError] = useState('');
  const state = wordTableState(target.editor);
  if (!state) return null;
  const update = (part: WordTablePatch) => { setPatch(previous => ({ ...previous, ...part })); setError(''); };
  const apply = () => {
    if (!restore(target)) { setError('文档已变化，请取消并重新打开属性。'); return; }
    if (!applyWordTableProperties(target.editor, patch)) { setError('请检查尺寸、颜色和边框参数。'); return; }
    target.editor.view.focus(); onError(undefined); onClose();
  };
  const border = patch.border ?? { mode: 'all', color: '333333', widthPt: 0.5, style: 'single' };
  return <Dialog open title="表格属性" description="尺寸和对齐应用于当前表格；行高、列宽、底纹和垂直对齐应用于所选行列或单元格。" onClose={onClose}
    footer={<><Button onClick={onClose}>取消</Button><Button variant="primary" disabled={!Object.keys(patch).length} onClick={apply}>应用表格属性</Button></>}>
    <div className="grid grid-cols-2 gap-3 text-[12px]">
      <Field label="表格宽度 (px)"><Input aria-label="表格宽度 (px)" type="number" min={1} max={31680} value={patch.widthPx ?? state.table.attrs.widthPx ?? ''} placeholder="自动" onChange={event => update({ widthPx: event.target.value === '' ? NaN : Number(event.target.value) })} /></Field>
      <Field label="表格对齐"><Select aria-label="表格对齐" value={patch.align ?? state.table.attrs.tblAlign ?? 'left'} onChange={event => update({ align: event.target.value as WordTablePatch['align'] })}><option value="left">左对齐</option><option value="center">居中</option><option value="right">右对齐</option></Select></Field>
      <Field label="列宽 (px)"><Input aria-label="列宽 (px)" type="number" min={1} value={patch.columnWidthPx ?? state.cell.attrs.colwidth?.[0] ?? ''} placeholder="自动" onChange={event => update({ columnWidthPx: event.target.value === '' ? NaN : Number(event.target.value) })} /></Field>
      <Field label="行高 (px)"><Input aria-label="行高 (px)" type="number" min={1} value={(patch.rowHeightTwips ?? state.row.attrs.heightTwips) == null ? '' : (patch.rowHeightTwips ?? state.row.attrs.heightTwips) / 15} placeholder="自动" onChange={event => update({ rowHeightTwips: event.target.value === '' ? NaN : Number(event.target.value) * 15 })} /></Field>
      <Field label="行高规则"><Select aria-label="行高规则" value={patch.rowHeightRule ?? state.row.attrs.heightRule ?? 'atLeast'} onChange={event => update({ rowHeightRule: event.target.value as 'atLeast' | 'exact' })}><option value="atLeast">最小值</option><option value="exact">固定值</option></Select></Field>
      <Field label="垂直对齐"><Select aria-label="垂直对齐" value={patch.vAlign ?? state.cell.attrs.vAlign ?? 'top'} onChange={event => update({ vAlign: event.target.value as WordTablePatch['vAlign'] })}><option value="top">顶端</option><option value="center">居中</option><option value="bottom">底端</option></Select></Field>
      <Field label="单元格底纹"><div className="flex gap-2"><input type="color" aria-label="单元格底纹" value={`#${patch.fill ?? state.cell.attrs.fill ?? 'ffffff'}`} onChange={event => update({ fill: event.target.value.slice(1) })} /><Button size="sm" onClick={() => update({ fill: null })}>清除底纹</Button></div></Field>
      <Field label="边框范围"><Select aria-label="边框范围" value={patch.border?.mode ?? ''} onChange={event => { if (event.target.value) update({ border: { ...border, mode: event.target.value as typeof border.mode } }); }}><option value="">保持原有边框</option><option value="all">所有边框</option><option value="none">无边框</option><option value="outer">外侧边框</option><option value="inner">内部边框</option><option value="top">上边框</option><option value="bottom">下边框</option><option value="left">左边框</option><option value="right">右边框</option></Select></Field>
      <Field label="边框颜色"><input aria-label="边框颜色" type="color" value={`#${border.color}`} onChange={event => update({ border: { ...border, color: event.target.value.slice(1) } })} /></Field>
      <Field label="边框线宽 (pt)"><Input aria-label="边框线宽 (pt)" type="number" min={0.25} max={12} step={0.25} value={border.widthPt} onChange={event => update({ border: { ...border, widthPt: Number(event.target.value) } })} /></Field>
      <Field label="边框线型"><Select aria-label="边框线型" value={border.style} onChange={event => update({ border: { ...border, style: event.target.value as typeof border.style } })}><option value="single">实线</option><option value="double">双线</option><option value="dashed">虚线</option><option value="dotted">点线</option></Select></Field>
    </div>
    {error ? <p role="alert" className="mt-3 text-xs text-[var(--mo-danger)]">{error}</p> : null}
  </Dialog>;
}
