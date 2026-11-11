/** 功能：复用既有样式库和命名对话框，所有文档修改直接提交核心事务；时间：2026-11-10；作者：lq。 */
import { useEffect, useMemo, useState } from 'react';
import type { DocStyleDefinition } from '@mini-office/document-model';
import { Button, Dialog, Field, Input } from '@mini-office/ui';
import type { WordCoreSession } from '../../word-core/src/index';
import { applyWordStyle, captureWordFormat, defineWordStyle, deleteWordStyle, importWordStyles, renameWordStyle, wordStyleFormat, type Editor, type WordFormatSnapshot } from '../../word-core/src/editing/index';
import { coreToolbarFormatting } from './core-formatting';
import { NewStyleDialog } from './NewStyleDialog';
import { wordStyleCatalogOf } from '../../word-core/src/editing/styles';
import { SaveStyleSetDialog } from './StyleSetDialog';
import { BUILTIN_STYLES, newStyleId, addStyleSet, readStyleSets, removeStyleSet } from './styles';
import type { StylePickerProps } from './StylePicker';

/** 功能：样式集和预设仅转换格式键及单位，禁止通过旧 DocModel 写回正文；时间：2026-11-10；作者：lq。 */
export function coreStyleDefinition(style: DocStyleDefinition, core: WordCoreSession) {
  const p = style.paragraph, t = style.run;
  const base = wordStyleFormat(core.formattingSource().styles.get('Normal') ?? { styleId: 'Normal', name: 'Normal', type: 'paragraph' }, core.parsed);
  const format: WordFormatSnapshot = {
    text: { ...base.text, font: t.fontFamily ?? base.text.font, fontAscii: t.fontFamily ?? base.text.fontAscii,
      sizeHalfPoints: (t.fontSizePt ?? (base.text.sizeHalfPoints ?? 22) / 2) * 2,
      bold: t.bold ?? false, italic: t.italic ?? false, underline: t.underline ?? false, strike: t.strike ?? false, color: t.color, highlight: t.highlight },
    paragraph: { align: p.alignment ?? 'left', indentLeft: Math.round((p.indentLeftPx ?? 0) * 15), indentRight: Math.round((p.indentRightPx ?? 0) * 15),
      indentFirstLine: Math.round((p.firstLineIndentPx ?? 0) * 15), spaceBefore: Math.round((p.spaceBeforePx ?? 0) * 15), spaceAfter: Math.round((p.spaceAfterPx ?? 0) * 15),
      lineRule: p.lineHeightExactPx != null ? 'exact' : p.lineHeightAtLeastPx != null ? 'atLeast' : 'auto',
      lineRawTwips: Math.round((p.lineHeightExactPx ?? p.lineHeightAtLeastPx ?? (p.lineHeight ?? 1) * 16) * 15),
      lineSpacing: p.lineHeightExactPx != null || p.lineHeightAtLeastPx != null ? null : p.lineHeight ?? 1,
      keepNext: p.keepNext ?? false, keepLines: p.keepLines ?? false, widowControl: p.widowControl ?? true, pageBreakBefore: p.pageBreakBefore ?? false,
      borders: p.borders ? Object.entries({ t: p.borders.top, l: p.borders.left, b: p.borders.bottom, r: p.borders.right }).filter(([, line]) => line && line.style !== 'none').map(([side]) => side).join('') : '',
      borderLines: p.borders ? JSON.stringify(Object.fromEntries(Object.entries({ t: p.borders.top, l: p.borders.left, b: p.borders.bottom, r: p.borders.right }).filter(([, line]) => line).map(([side, line]) => [side, { color: line!.color.replace(/^#/, ''), szPt: line!.widthPx * 0.75, val: line!.style, ...(line!.style === 'none' ? { none: true } : {}) }]))) : null,
      shadingFill: p.backgroundColor?.replace(/^#/, '') ?? null,
      tabStops: p.tabStops?.length ? JSON.stringify(p.tabStops.map(tab => ({ pos: Math.round(tab.posPx * 15), val: tab.alignment, leader: tab.leader }))) : null },
    ...(style.list ? { list: { kind: style.list.ordered ? 'ordered' as const : 'bullet' as const, level: style.list.level, numId: core.styleNumbering(style.list.ordered ? 'ordered' : 'bullet') } } : {}),
  };
  const definition = defineWordStyle(style.id, style.name, format, p.outlineLevel == null ? undefined : p.outlineLevel + 1);
  definition.upsert.builtIn = style.origin === 'preset';
  return definition;
}

/** 功能：只把核心样式摘要转换为卡片预览，不构建旧正文树；时间：2026-11-10；作者：lq。 */
function previewStyle(id: string, name: string, f: WordFormatSnapshot, headingLevel?: number): DocStyleDefinition {
  const { text: t, paragraph: p } = f;
  // 功能：样式集导出保留列表、段落边框、底纹及制表位，重导入不退化为纯文字样式；时间：2026-11-11；作者：lq。
  const lines = p.borderLines ? JSON.parse(p.borderLines) : {};
  const borders = Object.fromEntries(Object.entries({ t: 'top', l: 'left', b: 'bottom', r: 'right' }).filter(([side]) => lines[side]).map(([side, name]) => [name, { style: lines[side].none ? 'none' : lines[side].val ?? 'single', widthPx: (lines[side].szPt ?? 0.5) / 0.75, color: `#${lines[side].color ?? '444444'}` }]));
  return { id, name, origin: 'custom', run: { fontFamily: t.font, fontSizePt: t.sizeHalfPoints == null ? undefined : t.sizeHalfPoints / 2,
    bold: t.bold, italic: t.italic, underline: t.underline, strike: t.strike, color: t.color, highlight: t.highlight },
    paragraph: { alignment: p.align as 'left', outlineLevel: headingLevel == null ? undefined : headingLevel - 1, indentLeftPx: Number(p.indentLeft ?? 0) / 15, indentRightPx: Number(p.indentRight ?? 0) / 15,
      firstLineIndentPx: Number(p.indentFirstLine ?? 0) / 15, spaceBeforePx: Number(p.spaceBefore ?? 0) / 15, spaceAfterPx: Number(p.spaceAfter ?? 0) / 15,
      lineHeight: p.lineRule === 'auto' ? p.lineSpacing ?? 1 : undefined,
      lineHeightExactPx: p.lineRule === 'exact' ? Number(p.lineRawTwips) / 15 : undefined,
      lineHeightAtLeastPx: p.lineRule === 'atLeast' ? Number(p.lineRawTwips) / 15 : undefined,
      keepNext: p.keepNext ?? undefined, keepLines: p.keepLines ?? undefined, pageBreakBefore: p.pageBreakBefore,
      borders: Object.keys(borders).length ? borders : undefined, backgroundColor: p.shadingFill ? `#${p.shadingFill}` : undefined,
      tabStops: p.tabStops ? JSON.parse(p.tabStops).map((tab: { pos: number; val: string; leader?: string }) => ({ posPx: tab.pos / 15, alignment: tab.val, leader: tab.leader })) : undefined },
    ...(f.list ? { list: { ordered: f.list.kind === 'ordered', level: f.list.level } } : {}),
  };
}

/** 功能：现有样式库的状态与对话框，只持有 UI 状态，文档定义保存在核心中；时间：2026-11-10；作者：lq。 */
export function useWordCoreStyles(core: WordCoreSession, readonly: boolean, run: (command: (editor: Editor) => unknown) => void, capture: () => void, onError: (error: string) => void) {
  const [open, setOpen] = useState(false), [more, setMore] = useState(false), [createOpen, setCreateOpen] = useState(false);
  // 功能：经典菜单与命令面板进入同一新核心样式库，保留当前区域选区；时间：2026-11-11；作者：lq。
  useEffect(() => {
    const show = () => { capture(); setOpen(true); };
    window.addEventListener('mo:styles', show);
    return () => window.removeEventListener('mo:styles', show);
  }, [capture]);
  const [sets, setSets] = useState(readStyleSets), [saveSetOpen, setSaveSetOpen] = useState(false);
  const [renaming, setRenaming] = useState<string>(), [name, setName] = useState('');
  // 功能：样式目录未改时复用格式预览，滚动和选区变化不逐个重建整套样式卡片；时间：2026-11-11；作者：lq。
  const catalog = wordStyleCatalogOf(core.editor.state.doc);
  const source = useMemo(() => core.formattingSource(), [core, catalog]);
  const styles = source.styles;
  const builtins = BUILTIN_STYLES;
  const builtinIds = new Set(builtins.map(item => item.id));
  const previews = useMemo(() => [...styles.values()].filter(item => item.type === 'paragraph' && !(item.styleId.startsWith('MoRetained_') && item.semiHidden)).map(style => previewStyle(style.styleId, style.name, wordStyleFormat(style, source), style.headingLevel)), [source, styles]);
  const presets = builtins.map(item => {
    const existing = previews.find(style => style.id === item.id);
    return existing ? { ...existing, name: styles.get(item.id)!.name.toLowerCase() === (item.builtinName ?? '').toLowerCase() ? item.name : existing.name } : item;
  });
  const extras = previews.filter(item => !builtinIds.has(item.id));
  const exportableIds = new Set(wordStyleCatalogOf(core.editor.state.doc).definitions.filter(item => item.format && !core.parsed.styles.has(item.upsert.styleId)).map(item => item.upsert.styleId));
  const exportable = extras.filter(item => exportableIds.has(item.id));
  const ids = new Set<string>();
  const editor = core.commandEditor();
  for (const { $from, $to } of editor.state.selection.ranges) editor.state.doc.nodesBetween($from.pos, $to.pos, node => { if (node.isTextblock) ids.add(String(node.attrs.styleId ?? 'Normal')); });
  if (editor.state.selection.empty) ids.add(String(editor.state.selection.$from.parent.attrs.styleId ?? 'Normal'));
  const currentId = ids.size === 1 ? [...ids][0] : undefined;
  const format = coreToolbarFormatting(core);
  const names = previews.map(item => item.name);
  /** 功能：预设定义只在首次使用时写入，已有导入定义以文档自身内容为准；时间：2026-11-10；作者：lq。 */
  const apply = (id: string) => run(target => {
    // 功能：正文与受控页眉页脚均持有主目录投影，其他未接线区域保持禁用；时间：2026-11-11；作者：lq。
    if (!('miniOfficeStyles' in target.state.doc.attrs)) return;
    const preset = builtins.find(item => item.id === id);
    applyWordStyle(target, core.parsed, id, styles.has(id) ? undefined : preset ? coreStyleDefinition(preset, core) : undefined);
    setOpen(false);
  });
  const invalidName = !name.trim() || name.trim().length > 40 || previews.some(item => item.id !== renaming && item.name === name.trim());
  const picker: StylePickerProps = {
    presets, documentStyles: extras, currentStyleId: currentId, currentStyleName: ids.size > 1 ? '混合样式' : presets.find(item => item.id === currentId)?.name ?? extras.find(item => item.id === currentId)?.name ?? '自定义格式',
    showMore: more, defaultFont: source.docDefaults?.eastAsiaFont ?? source.docDefaults?.asciiFont ?? '宋体', defaultFontSizePt: (source.docDefaults?.sizeHalfPoints ?? 22) / 2,
    readonly: readonly || !('miniOfficeStyles' in editor.state.doc.attrs) || !editor.schema.nodes.docParagraph, styleSets: sets, open, onOpenChange: setOpen,
    // 功能：清除已迁移的字符/段落格式回到文档正文样式，保留链接及语言来源；时间：2026-11-10；作者：lq。
    onApply: apply, onClearFormatting: () => apply('Normal'),
    onNewStyle: () => { capture(); setOpen(false); setCreateOpen(true); }, onToggleMore: () => setMore(value => !value),
    canDeleteStyle: id => id !== 'Normal' && !core.parsed.styles.get(id)?.isDefault,
    canSaveStyleSet: exportable.length > 0,
    onDeleteStyle: id => run(() => deleteWordStyle(core.editor, core.parsed, id)),
    onRenameStyle: id => { capture(); setRenaming(id); setName(styles.get(id)?.name ?? ''); setOpen(false); },
    onSaveStyleSet: () => { setOpen(false); setSaveSetOpen(true); },
    onApplyStyleSet: set => run(target => {
      try {
        const taken = new Set(styles.keys());
        // 功能：先校验整组再提交，拒绝部分导入；每组仅一个事务及一次撤销；时间：2026-11-10；作者：lq。
        const definitions = set.styles.map(item => { const id = taken.has(item.id) ? newStyleId(taken) : item.id; taken.add(id); return coreStyleDefinition({ ...item, id }, core); });
        importWordStyles(target, definitions);
      } catch (error) { onError(String(error)); }
    }),
    onDeleteStyleSet: id => setSets(removeStyleSet(id)), onBeforeInteraction: capture,
  };
  return { picker, apply, dialogs: <>
    <NewStyleDialog open={createOpen} formatPreview={{ run: format.activeStyle, style: format.paragraph }}
      doc={{ defaultFont: picker.defaultFont, defaultFontSizePt: picker.defaultFontSizePt }} existingNames={names} defaultName="新样式"
      onClose={() => setCreateOpen(false)} onCreate={(value, base) => run(target => {
        const snapshot = base === 'current' ? captureWordFormat(target, source) : wordStyleFormat(source.styles.get('Normal') ?? { styleId: 'Normal', name: 'Normal', type: 'paragraph' }, source);
        const id = newStyleId(styles.keys());
        // 功能：从标题新建时保留原大纲级别，选择正文基准时回到普通段落；时间：2026-11-10；作者：lq。
        const headingLevel = base === 'current' ? target.state.selection.$from.parent.type.name === 'docHeading' ? Number(target.state.selection.$from.parent.attrs.level) : source.styles.get(String(target.state.selection.$from.parent.attrs.styleId))?.headingLevel : undefined;
        applyWordStyle(target, core.parsed, id, defineWordStyle(id, value, snapshot, headingLevel));
      })} />
    <SaveStyleSetDialog open={saveSetOpen} defaultName="我的样式集" styleCount={exportable.length} existingNames={sets.map(item => item.name)}
      onClose={() => setSaveSetOpen(false)} onSave={value => setSets(addStyleSet(value, exportable))} />
    <Dialog open={!!renaming} onClose={() => setRenaming(undefined)} title="重命名样式" width="sm" footer={<>
      <Button onClick={() => setRenaming(undefined)}>取消</Button><Button variant="primary" disabled={invalidName} onClick={() => { run(target => renameWordStyle(target, core.parsed, renaming!, name)); setRenaming(undefined); }}>确定</Button></>}>
      <Field label="样式名称" hint="最多 40 个字符，不能与其他样式同名"><Input autoFocus value={name} maxLength={40} onChange={event => setName(event.target.value)} /></Field>
    </Dialog>
  </> };
}
