/** 功能：在 Mini Office 现有外壳和工具栏中挂载唯一新 Word 会话，完成 M2 产品纵向接线；时间：2026-11-09；作者：lq。 */
import { useEffect, useRef, useState } from 'react';
import { useWorkspaceStore, type DocumentSession } from '@mini-office/document-model';
import { WordToolbar } from './WordToolbar';
import { WordCoreSurfaceView } from './WordCoreSurface';

/** 功能：正文、保存和撤销共享工作区核心，组件卸载不销毁标签会话；时间：2026-11-09；作者：lq。 */
export function WordCoreEditor({ session }: { session: DocumentSession }) {
  const core = session.wordCore!;
  const [, refresh] = useState(0);
  const [error, setError] = useState<string>();
  const [marks, setMarks] = useState(false);
  const scroll = useRef<HTMLDivElement>(null);
  const zoom = session.view.zoom ?? 1;
  const save = (asNew = false) => { void useWorkspaceStore.getState().saveSession(session.id, { asNew }); };
  useEffect(() => core.subscribe(() => refresh(revision => revision + 1)), [core]);
  const state = core.state();
  // 功能：M2 仅暴露已实际迁移的基础命令；其余入口在现有工具栏内禁用，按 M3 继续迁移；时间：2026-11-09；作者：lq。
  const pending = () => {};
  return <div className="flex h-full min-h-0 flex-col" data-word-engine="word-core">
    <WordToolbar doc={{ defaultFontSizePt: 11 }} migrationPending readonly={session.readonly}
      activeStyle={{ bold: core.editor.isActive('bold'), italic: core.editor.isActive('italic'), underline: core.editor.isActive('underline'), strike: core.editor.isActive('strike') }}
      canUndo={state.canUndo} canRedo={state.canRedo} saving={session.state === 'SAVING'}
      onSave={() => save()} onSaveAs={() => save(true)} onUndo={() => core.editor.chain().focus().undo().run()} onRedo={() => core.editor.chain().focus().redo().run()}
      onInlineCommand={command => { core.editor.chain().focus().toggleMark(command === 'strikeThrough' ? 'strike' : command).run(); }}
      zoom={zoom} pageCount={session.view.pageCount ?? 1} currentPage={session.view.page ?? 0}
      onZoomChange={value => useWorkspaceStore.getState().updateView(session.id, { zoom: value })}
      onFitWidth={() => { const width = scroll.current?.clientWidth; if (width) useWorkspaceStore.getState().updateView(session.id, { zoom: Math.min(2.5, Math.max(0.4, (width - 48) / (core.editor.view.dom.offsetWidth || 794))) }); }}
      onShowCompatibility={() => useWorkspaceStore.getState().openDialog({ type: 'compatibility', sessionId: session.id })}
      showFormattingMarks={marks} onToggleFormattingMarks={() => setMarks(value => !value)}
      findOpen={false} propertiesOpen={false} rulerOpen={false} formatPainterActive={false}
      onFontFormat={pending} onColorFormat={pending} onParagraphFormat={pending} onInsertTable={pending} onInsertImage={pending} onPageBreak={pending}
      onHeaderFooter={pending} onMailMerge={pending} onToggleFind={pending} onToggleProperties={pending} onPrint={pending} onExport={pending}
      onAddBullets={pending} onAddNumbering={pending} onFormatPainter={pending} onOpenParagraphDialog={pending} onToggleRuler={pending} onInsertToc={pending} />
    {error ? <div role="alert" className="shrink-0 px-4 py-2 text-sm text-[var(--mo-danger)]">{error}</div> : null}
    <div ref={scroll} className="mo-scroll min-h-0 flex-1 overflow-auto bg-[var(--mo-bg)] p-6">
      <div className={`doc-zoom ${marks ? 'show-marks' : ''}`} style={{ zoom }}>
        <WordCoreSurfaceView session={core} onReady={() => setError(undefined)} onError={caught => setError(String(caught))}
          onLayoutError={caught => setError(String(caught))}
          onLayout={layout => { setError(undefined); const current = useWorkspaceStore.getState().sessions[session.id]; if (current && current.view.pageCount !== layout.pageCount) useWorkspaceStore.getState().updateView(session.id, { pageCount: layout.pageCount }); }} />
      </div>
    </div>
  </div>;
}
