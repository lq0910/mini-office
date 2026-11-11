/** 功能：在现有外壳/工具栏中挂载唯一 Word 会话，承接 M2 链路与 M3 格式迁移；时间：2026-11-10；作者：lq。 */
import { useCallback, useEffect, useRef, useState } from 'react';
import { useWorkspaceStore, type DocumentSession } from '@mini-office/document-model';
import { WordToolbar } from './WordToolbar';
import { WordCoreSurfaceView } from './WordCoreSurface';
import type { WordCoreSurface } from './core-surface';
import { insertPageBreak, toggleWordInlineFormat, readWordList, changeWordListLevel, captureWordFormat, applyWordFormatPainter, type WordFormatSnapshot, type Editor } from '../../word-core/src/editing/index';
// 功能：从编辑器公开类型推导选区书签，保持依赖方向；时间：2026-11-10；作者：lq。
type SelectionBookmark = ReturnType<Editor['state']['selection']['getBookmark']>;
import { coreFontFormat, coreParagraphFormat, coreToolbarFormatting } from './core-formatting';
import { ParagraphDialog } from './ParagraphDialog';
// 功能：沿用查找浮层并开放当前列表状态；时间：2026-11-10；作者：lq。
import { WordCoreFindBar } from './WordCoreFindBar';
// 功能：原样式库、命名对话框和格式刷复用唯一新核心；时间：2026-11-10；作者：lq。
import { useWordCoreStyles } from './WordCoreStyles';
// 功能：沿用属性面板和页眉页脚弹窗入口，设置只写入核心事务；时间：2026-11-10；作者：lq。
// 功能：标尺和目录接入现有视图/引用入口，不经过旧文档模型；时间：2026-11-11；作者：lq。
import { WordCoreRuler } from './WordCoreRuler';
import { insertWordToc, followWordToc } from '../../word-core/src/editing/toc';
import { WordCorePagePanel } from './WordCorePagePanel';
import { WordCoreHeaderFooterDialog } from './WordCoreHeaderFooterDialog';
// 功能：图片与表格菜单、属性和原生结构命令接入当前核心；时间：2026-11-11；作者：lq。
import { useWordCoreObjects } from './WordCoreObjects';
import { insertWordTable } from '../../word-core/src/editing/objects';
// 功能：显示活动编辑区粘贴格式切换，不维护另一份正文；时间：2026-11-11；作者：lq。
import { WordCorePasteOptions } from './WordCorePasteOptions';
// 功能：原生字段与邮件数据 UI 共用确定版本快照，合并不改源模板；时间：2026-11-11；作者：lq。
import { WordCoreMailMerge } from './WordCoreMailMerge';
import { insertWordMergeField } from '../../word-core/src/editing/mail-merge';
// 功能：打印和 PDF 使用当前原生排版快照；时间：2026-11-11；作者：lq。
import { useWordCoreOutput } from './WordCoreOutput';
// 功能：Ctrl+滚轮只缩放当前正文，沿用状态栏范围和未保存状态；时间：2026-11-11；作者：lq。
import { useWordCoreWheelZoom } from './core-wheel-zoom';

/** 功能：正文、保存和撤销共享工作区核心，组件卸载不销毁标签会话；时间：2026-11-09；作者：lq。 */
export function WordCoreEditor({ session }: { session: DocumentSession }) {
  const core = session.wordCore!;
  const [, refresh] = useState(0);
  const [error, setError] = useState<string>();
  const [marks, setMarks] = useState(false);
  const [paragraphOpen, setParagraphOpen] = useState(false);
  const [propertiesOpen, setPropertiesOpen] = useState(false);
  const [headerFooter, setHeaderFooter] = useState<'header' | 'footer'>();
  const [findOpen, setFindOpen] = useState(false);
  const [mailOpen, setMailOpen] = useState(false);
  const [painter, setPainter] = useState<{ format: WordFormatSnapshot; sticky: boolean }>();
  const paintGesture = useRef(false);
  // 功能：目录拖选不误触导航，只有指针停留在同一位置的单击才跳转；时间：2026-11-11；作者：lq。
  const tocPointer = useRef<{ x: number; y: number }>();
  const paintFrame = useRef<number>();
  const painterRef = useRef(painter);
  painterRef.current = painter;
  // 功能：浮层/原生选择框夺焦前记录当前活动编辑器和选区，执行命令时恢复；时间：2026-11-10；作者：lq。
  const selection = useRef<{ editor: Editor; bookmark: SelectionBookmark }>();
  const scroll = useRef<HTMLDivElement>(null);
  const surface = useRef<WordCoreSurface>();
  // 功能：一帧只读取一次当前页几何，快速拖动滚动条不逐事件强制布局；时间：2026-11-11；作者：lq。
  const pageFrame = useRef<number>();
  const rulerOpen = useWorkspaceStore(state => state.settings.wordRuler === true);
  const zoom = session.view.zoom ?? 1;
  useWordCoreWheelZoom(scroll, session.id);
  /** 功能：当前页按真实纸张和可见窗口计算，滚动/缩放只更新视图而不改文档历史；时间：2026-11-11；作者：lq。 */
  const updateCurrentPage = useCallback(() => {
    // 功能：工作区先移除会话后 React 才清理订阅，关闭期间拒绝迟到测量；时间：2026-11-11；作者：lq。
    if (core.editor.isDestroyed) return;
    const viewport = scroll.current, layout = surface.current?.layout.current();
    const paper = core.editor.view.dom.closest<HTMLElement>('.page-wrap');
    if (!viewport || !layout?.frames.length || !paper || !viewport.clientHeight) return;
    const box = paper.getBoundingClientRect(), factor = box.width / paper.offsetWidth || 1;
    const y = (viewport.getBoundingClientRect().top + Math.min(160, viewport.clientHeight / 3) - box.top) / factor;
    let page = 0;
    layout.frames.forEach((frame, index) => { if (frame.top <= y) page = index; });
    const store = useWorkspaceStore.getState();
    if (store.sessions[session.id]?.view.page !== page) store.updateView(session.id, { page });
  }, [core, session.id]);
  const scheduleCurrentPage = useCallback(() => {
    if (pageFrame.current != null) return;
    pageFrame.current = requestAnimationFrame(() => { pageFrame.current = undefined; updateCurrentPage(); });
  }, [updateCurrentPage]);
  // 功能：缩放和窗口尺寸改变后使用新几何同步当前页，卸载取消尚未执行的测量；时间：2026-11-11；作者：lq。
  useEffect(() => {
    scheduleCurrentPage(); const observer = new ResizeObserver(scheduleCurrentPage);
    if (scroll.current) observer.observe(scroll.current);
    return () => { if (pageFrame.current != null) cancelAnimationFrame(pageFrame.current); pageFrame.current = undefined; observer.disconnect(); };
  }, [zoom, scheduleCurrentPage]);
  const save = (asNew = false) => { void useWorkspaceStore.getState().saveSession(session.id, { asNew }); };
  useEffect(() => core.subscribe(() => refresh(revision => revision + 1)), [core]);
  // 功能：异步剪贴板图片失败沿用编辑器错误入口，不能静默丢失截图或覆盖新正文；时间：2026-11-11；作者：lq。
  useEffect(() => {
    const report = (event: Event) => setError(String((event as CustomEvent).detail));
    // 功能：清理持有的 DOM，不再访问已销毁编辑器的 view 代理；时间：2026-11-11；作者：lq。
    const dom = core.editor.view.dom;
    dom.addEventListener('mo:word-paste-error', report);
    return () => dom.removeEventListener('mo:word-paste-error', report);
  }, [core]);
  // 功能：菜单、命令面板与 Mod+F 共用外壳事件，只在当前会话打开查找；时间：2026-11-10；作者：lq。
  useEffect(() => {
    const toggle = () => { selection.current = undefined; setFindOpen(value => !value); };
    window.addEventListener('mo:find', toggle);
    return () => window.removeEventListener('mo:find', toggle);
  }, [core]);
  const state = core.state();
  const format = coreToolbarFormatting(core);
  const list = readWordList(core.commandEditor());
  /** 功能：保持当前文本框或表格选区，命令执行后将焦点还给同一编辑区；时间：2026-11-10；作者：lq。 */
  const run = (command: (editor: Editor) => unknown) => {
    if (session.readonly) return;
    const target = selection.current?.editor ?? core.commandEditor();
    if (target.isDestroyed || target.view.composing) return;
    if (selection.current) {
      const restored = selection.current.bookmark.resolve(target.state.doc);
      if (!restored.eq(target.state.selection)) target.view.dispatch(target.state.tr.setSelection(restored));
    }
    command(target);
    selection.current = undefined;
    // 功能：颜色输入逐键生效时保留输入框焦点，其他命令同步回到原编辑区以承接紧随的键盘输入；时间：2026-11-10；作者：lq。
    // 功能：样式目录事务可能重挂文本框，焦点恢复必须使用当前有效区域；时间：2026-11-11；作者：lq。
    if (!document.activeElement?.matches('input[aria-label="颜色值"], input[type="color"]')) core.commandEditor().view.focus();
  };
  /** 功能：对话框夺焦前保存选区，样式事务恢复到相同编辑区；时间：2026-11-10；作者：lq。 */
  const captureSelection = () => { const editor = core.commandEditor(); selection.current = { editor, bookmark: editor.state.selection.getBookmark() }; };
  const styles = useWordCoreStyles(core, session.readonly, run, captureSelection, setError);
  const objects = useWordCoreObjects(core, session.readonly, setError);
  const output = useWordCoreOutput(session);
  /** 功能：Esc 在任意焦点退出格式刷，不改正文或历史；时间：2026-11-10；作者：lq。 */
  useEffect(() => {
    const cancel = (event: KeyboardEvent) => { if (event.key === 'Escape' && painter) { painterRef.current = undefined; setPainter(undefined); paintGesture.current = false; if (paintFrame.current != null) cancelAnimationFrame(paintFrame.current); } };
    window.addEventListener('keydown', cancel);
    return () => window.removeEventListener('keydown', cancel);
  }, [painter]);
  // 功能：标签卸载或切换时取消尚未执行的拖选提交，防止隐藏文档继续被刷格式；时间：2026-11-10；作者：lq。
  useEffect(() => () => { if (paintFrame.current != null) cancelAnimationFrame(paintFrame.current); }, [core]);
  // 功能：工具栏操作当前区域能力，表格和目录等结构入口按区域 schema 限定；时间：2026-11-11；作者：lq。
  return <WordCoreMailMerge session={session} open={mailOpen} onClose={() => setMailOpen(false)} onInsertField={name => run(editor => insertWordMergeField(editor, name))}>{mail => <div className="flex h-full min-h-0 flex-col" data-word-engine="word-core" onMouseDownCapture={objects.onMouseDown} onContextMenuCapture={objects.onContextMenu}
    onPointerDownCapture={event => { tocPointer.current = { x: event.clientX, y: event.clientY }; }}
    onClickCapture={event => {
      // 功能：单击目录直接导航，修饰键单击继续兼容；拖选或组合输入期间不夺取正文选区；时间：2026-11-11；作者：lq。
      if (core.editor.view.composing || (tocPointer.current && Math.hypot(event.clientX - tocPointer.current.x, event.clientY - tocPointer.current.y) > 4)) return;
      const line = (event.target as HTMLElement).closest<HTMLElement>('.doc-toc-line');
      if (!line || !core.editor.view.dom.contains(line)) return;
      event.preventDefault(); event.stopPropagation(); selection.current = undefined;
      if (!followWordToc(core.editor, core.formattingSource(), Array.from(core.editor.view.dom.querySelectorAll('.doc-toc-line')).indexOf(line))) setError('此目录条目未找到对应标题或书签，请更新目录后重试。');
    }}
    onKeyDownCapture={event => {
      // 功能：沿用字符/对齐/段落快捷键，仅在编辑区执行且优先保护中文组合输入；时间：2026-11-10；作者：lq。
      const target = event.target as HTMLElement;
      if (!target.closest('.ProseMirror') || target.closest('input, textarea, select') || event.nativeEvent.isComposing || event.keyCode === 229 || core.commandEditor().view.composing || session.readonly) return;
      if (!(event.metaKey || event.ctrlKey)) return;
      const key = event.key.toLowerCase();
      if ((event.altKey && /^[1-4]$/.test(event.key)) || (event.shiftKey && !event.altKey && key === 'n')) {
        event.preventDefault(); event.stopPropagation(); selection.current = undefined;
        styles.apply(event.altKey ? `Heading${event.key}` : 'Normal');
      } else if (event.altKey && key === 'm') {
        event.preventDefault(); event.stopPropagation();
        const editor = core.commandEditor(); selection.current = { editor, bookmark: editor.state.selection.getBookmark() };
        setParagraphOpen(true);
      } else if (!event.altKey && !event.shiftKey && ['b', 'i', 'u', 'e', 'l', 'r', 'j'].includes(key)) {
        event.preventDefault(); event.stopPropagation(); selection.current = undefined;
        if (['b', 'i', 'u'].includes(key)) run(editor => toggleWordInlineFormat(editor, core.formattingSource(), key === 'b' ? 'bold' : key === 'i' ? 'italic' : 'underline'));
        else run(editor => coreParagraphFormat(editor, { alignment: ({ e: 'center', l: 'left', r: 'right', j: 'justify' } as const)[key as 'e' | 'l' | 'r' | 'j'] }));
      }
    }}>
    <WordToolbar doc={{ defaultFontSizePt: (core.parsed.docDefaults?.sizeHalfPoints ?? 22) / 2 }} migrationPending migrated={['font', 'paragraph', 'pageBreak', 'find', 'pageLayout', 'headerFooter', 'ruler', 'mailMerge', 'print', 'export', ...(core.commandEditor().schema.nodes.docTable ? ['table' as const] : []), ...((core.commandEditor().schema.nodes.docProtected || core.commandEditor().schema.nodes.docInlineImage) ? ['image' as const] : []), ...(core.commandEditor() === core.editor ? ['toc' as const] : []), 'formatPainter', ...((core.commandEditor().schema.nodes.docListItem || core.commandEditor().schema.nodes.docParagraph?.spec.attrs?.listMarker) ? ['list' as const] : [])]} readonly={session.readonly || mail.previewOn}
      stylePicker={styles.picker}
      listKind={list.kind} listLevel={list.level} listMixed={list.mixed}
      onChangeListLevel={delta => run(editor => changeWordListLevel(editor, delta))}
      activeStyle={format.activeStyle} paragraphStyle={format.paragraph} mixedText={format.mixedText} mixedParagraph={format.mixedParagraph}
      onBeforeInteraction={captureSelection}
      canUndo={state.canUndo} canRedo={state.canRedo} saving={session.state === 'SAVING'}
      onSave={() => save()} onSaveAs={() => save(true)} onUndo={() => run(editor => editor.commands.undo())} onRedo={() => run(editor => editor.commands.redo())}
      onInlineCommand={command => run(editor => toggleWordInlineFormat(editor, core.formattingSource(), command === 'strikeThrough' ? 'strike' : command))}
      zoom={zoom} pageCount={session.view.pageCount ?? 1} currentPage={session.view.page ?? 0}
      onZoomChange={value => useWorkspaceStore.getState().updateView(session.id, { zoom: value })}
      onFitWidth={() => { const width = scroll.current?.clientWidth; if (width) useWorkspaceStore.getState().updateView(session.id, { zoom: Math.min(2.5, Math.max(0.4, (width - 48) / (core.editor.view.dom.offsetWidth || 794))) }); }}
      onShowCompatibility={() => useWorkspaceStore.getState().openDialog({ type: 'compatibility', sessionId: session.id })}
      showFormattingMarks={marks} onToggleFormattingMarks={() => setMarks(value => !value)}
      findOpen={findOpen} propertiesOpen={propertiesOpen} rulerOpen={rulerOpen} formatPainterActive={!!painter}
      onFontFormat={patch => run(editor => coreFontFormat(editor, patch))} onColorFormat={patch => run(editor => coreFontFormat(editor, patch))}
      onParagraphFormat={patch => run(editor => coreParagraphFormat(editor, patch))} onInsertTable={(rows, columns) => run(editor => insertWordTable(editor, rows, columns))} onInsertImage={() => run(editor => objects.chooseImage(editor))} onPageBreak={() => run(insertPageBreak)}
      onHeaderFooter={() => { selection.current = undefined; setHeaderFooter('footer'); }} onMailMerge={() => setMailOpen(value => !value)} onToggleFind={() => { selection.current = undefined; setFindOpen(value => !value); }} onToggleProperties={() => setPropertiesOpen(value => !value)} onPrint={output.openPrint} onExport={output.openExport}
      onAddBullets={() => run(editor => core.toggleList('bullet', editor))} onAddNumbering={() => run(editor => core.toggleList('ordered', editor))}
      onFormatPainter={sticky => {
        if (painter && !sticky) { setPainter(undefined); selection.current = undefined; return; }
        run(editor => setPainter({ format: painter?.format ?? captureWordFormat(editor, core.formattingSource()), sticky }));
      }} onOpenParagraphDialog={() => setParagraphOpen(true)} onToggleRuler={() => useWorkspaceStore.getState().updateSettings({ wordRuler: !rulerOpen })}
      onInsertToc={() => { try { run(editor => insertWordToc(editor, core.formattingSource())); setError(undefined); } catch (caught) { setError(String(caught instanceof Error ? caught.message : caught)); } }}
      onUpdateToc={() => { try { run(editor => insertWordToc(editor, core.formattingSource(), true)); setError(undefined); } catch (caught) { setError(String(caught instanceof Error ? caught.message : caught)); } }} />
    {mail.bar}
    {styles.dialogs}
    {objects.elements}
    {output.elements}
    <WordCorePasteOptions core={core} />
    {headerFooter && !session.readonly ? <WordCoreHeaderFooterDialog core={core} pageCount={session.view.pageCount ?? 1} initialRegion={headerFooter} onClose={() => { setHeaderFooter(undefined); selection.current = undefined; }} /> : null}
    <ParagraphDialog open={paragraphOpen} applyChangedOnly style={format.paragraph} fontSizePt={format.activeStyle.fontSizePt ?? 11}
      selectionHint={format.mixedParagraph.length ? '选区包含混合段落格式；仅应用本次修改的设置。' : undefined}
      onClose={() => { setParagraphOpen(false); selection.current = undefined; }}
      onApply={patch => run(editor => coreParagraphFormat(editor, patch))} />
    {error ? <div role="alert" className="shrink-0 px-4 py-2 text-sm text-[var(--mo-danger)]">{error}</div> : null}
    <div className="flex min-h-0 flex-1"><div className="relative min-w-0 flex-1">
    {findOpen ? <WordCoreFindBar core={core} readonly={session.readonly} onClose={() => { selection.current = undefined; setFindOpen(false); }} /> : null}
    {mail.previewOn ? mail.canvas : null}
    <div ref={scroll} className="mo-scroll h-full overflow-auto bg-[var(--mo-bg)] p-6" style={{ ...(painter ? { cursor: 'crosshair' } : {}), ...(mail.previewOn ? { display: 'none' } : {}) }}
      onScroll={scheduleCurrentPage}
      onPointerDownCapture={event => { paintGesture.current = !!painter && !!(event.target as HTMLElement).closest('.ProseMirror'); }}
      onPointerUp={event => {
        if (!paintGesture.current || !painter || event.button !== 0) return;
        paintGesture.current = false; selection.current = undefined;
        // 功能：等待 ProseMirror 接收本次 DOM 选区后再刷格式，避免用拖选前的旧范围；时间：2026-11-10；作者：lq。
        paintFrame.current = requestAnimationFrame(() => {
          if (core.commandEditor().isDestroyed || painterRef.current !== painter) return;
          if (applyWordFormatPainter(core.commandEditor(), painter.format) && !painter.sticky) setPainter(undefined);
        });
      }}>
      {rulerOpen ? <WordCoreRuler core={core} zoom={zoom} readonly={session.readonly} /> : null}
      <div className={`doc-zoom ${marks ? 'show-marks' : ''}`} style={{ zoom }}>
        <WordCoreSurfaceView session={core} onReady={ready => { surface.current = ready; setError(undefined); updateCurrentPage(); }} onError={caught => setError(String(caught))}
          onLayoutError={caught => setError(String(caught))}
          onLayout={layout => { setError(undefined); const current = useWorkspaceStore.getState().sessions[session.id]; if (current && current.view.pageCount !== layout.pageCount) useWorkspaceStore.getState().updateView(session.id, { pageCount: layout.pageCount }); updateCurrentPage(); }} />
      </div>
    </div>
    </div>
    {mail.panel}
    {propertiesOpen && !mailOpen ? <WordCorePagePanel core={core} session={session} readonly={session.readonly} onClose={() => setPropertiesOpen(false)} onApplied={() => { selection.current = undefined; }}
      onOpenParagraph={() => { captureSelection(); setParagraphOpen(true); }} onCompatibility={() => useWorkspaceStore.getState().openDialog({ type: 'compatibility', sessionId: session.id })} /> : null}
    </div>
  </div>}</WordCoreMailMerge>;
}
