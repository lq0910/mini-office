/** 功能：原生 Word 打印预览/翻页/缩放/页边距及 PDF 范围、进度、取消，输出同源实际页面；时间：2026-11-11；作者：lq。 */
import { useEffect, useRef, useState } from 'react';
import { Button, Dialog, Field, Input, ProgressBar, Select } from '@mini-office/ui';
import { useWorkspaceStore, type DocumentSession } from '@mini-office/document-model';
import { attachPreviewMargins, getPlatformAdapter, printPages, type PrintPreviewFit } from '@mini-office/platform-adapter';
import { exportElementsToPdf } from '@mini-office/export-service';
import { stripExtension, type ExportProgress } from '@mini-office/shared';
import { coreWordOutputBytes, mountCoreWordReport, type CoreWordReport } from './core-output';
import { wordPageLayoutOf } from '../../word-core/src/editing/page-layout';

/** 功能：明确拒绝非法或越界页码，避免错误范围静默导出另一组页面；时间：2026-11-11；作者：lq。 */
export function validateCorePageRange(range: string, count: number): void {
  if (!range.trim()) return;
  if (range.length > 1000) throw new Error('页码范围过长。');
  for (const part of range.split(',')) {
    const match = /^\s*(\d+)\s*(?:[-–]\s*(\d+))?\s*$/.exec(part);
    const first = Number(match?.[1]), last = Number(match?.[2] ?? match?.[1]);
    if (!match || !Number.isInteger(first) || first < 1 || last < first || last > count) throw new Error(`页码范围无效，请填写 1 至 ${count} 内的页码，例如 1-3,5。`);
  }
}
/** 功能：工具栏、菜单和快捷键共用输出状态，离开会话或取消时清理隐藏排版与大页面；时间：2026-11-11；作者：lq。 */
export function useWordCoreOutput(session: DocumentSession) {
  const core = session.wordCore!; const [mode, setMode] = useState<'print' | 'pdf'>();
  const [report, setReport] = useState<CoreWordReport>(); const [error, setError] = useState(''), [loading, setLoading] = useState(false), [busy, setBusy] = useState(false);
  const [page, setPage] = useState(0), [fit, setFit] = useState<PrintPreviewFit>('page'), [margins, setMargins] = useState(false), [size, setSize] = useState({ width: 650, height: 600 });
  const [name, setName] = useState(''), [range, setRange] = useState(''), [progress, setProgress] = useState<ExportProgress>();
  const [orientationScope, setOrientationScope] = useState<'section' | 'all'>('section');
  const host = useRef<HTMLDivElement>(null), cancel = useRef(false), mounted = useRef(true); const [revision, setRevision] = useState(core.state().revision);
  // 功能：输出进行中固定使用已确认快照，后台正文事务不销毁正在捕获的页面；时间：2026-11-11；作者：lq。
  const frozenRevision = useRef(revision); if (!busy) frozenRevision.current = revision;
  useEffect(() => core.subscribe(state => setRevision(state.revision)), [core]);
  useEffect(() => { mounted.current = true; return () => { mounted.current = false; cancel.current = true; }; }, [core]);
  const open = (mode: 'print' | 'pdf') => { if (busy) return; setError(''); setProgress(undefined); setPage(0); setRange(''); setName(`${stripExtension(session.displayName)}.pdf`); setMode(mode); };
  useEffect(() => {
    const print = () => open('print'), pdf = () => open('pdf'); window.addEventListener('mo:print', print); window.addEventListener('mo:export-pdf', pdf);
    return () => { window.removeEventListener('mo:print', print); window.removeEventListener('mo:export-pdf', pdf); };
  }, [core, session.displayName, busy]);
  useEffect(() => {
    if (!mode) return;
    let active = true, current: CoreWordReport | undefined; setLoading(true); setReport(undefined); setError('');
    void coreWordOutputBytes(core).then(bytes => { if (!active) throw new Error('已取消输出。'); return mountCoreWordReport(bytes, () => !active); }).then(result => {
      current = result; if (!active) { result.dispose(); return; } setReport(result); setLoading(false); setPage(page => Math.min(page, result.pages.length - 1));
    }).catch(error => { if (active) { setError(error instanceof Error ? error.message : String(error)); setLoading(false); } });
    return () => { active = false; current?.dispose(); };
  }, [core, mode, frozenRevision.current]);
  useEffect(() => {
    const target = host.current; if (!target || mode !== 'print') return;
    const observer = new ResizeObserver(() => setSize({ width: target.clientWidth, height: target.clientHeight })); observer.observe(target); return () => observer.disconnect();
  }, [mode]);
  useEffect(() => {
    const target = host.current; if (!target || !report || mode !== 'print' || busy) return;
    target.replaceChildren(); const geometry = report.frames[page], element = report.pages[page]; if (!geometry || !element) return;
    const scale = fit === 'actual' ? 1 : Math.max(0.1, Math.min(1, (size.width - 32) / geometry.widthPx, fit === 'page' ? (size.height - 32) / geometry.heightPx : 1));
    const stage = document.createElement('div'); stage.style.cssText = `position:relative;width:${geometry.widthPx * scale}px;height:${geometry.heightPx * scale}px;margin:16px auto;`;
    element.style.transform = `scale(${scale})`; element.style.transformOrigin = 'top left'; stage.append(element); target.append(stage);
    const section = core.sections()[geometry.section];
    const detachMargins = margins && section ? attachPreviewMargins({ frame: stage, page: element, scale, index: page }, {
      readonly: session.readonly,
      margins: { top: section.settings.marginTop / 15, right: section.settings.marginRight / 15, bottom: section.settings.marginBottom / 15, left: section.settings.marginLeft / 15 }, pageWidthPx: geometry.widthPx, pageHeightPx: geometry.heightPx,
      onChange: (edge, px) => {
        if (session.readonly || busy) return;
        const field = { left: 'marginLeft', right: 'marginRight', top: 'marginTop', bottom: 'marginBottom' }[edge];
        const catalog = structuredClone(wordPageLayoutOf(core.editor.state.doc)); catalog.sections[section.lastBlockIndex] = { ...catalog.sections[section.lastBlockIndex], [field]: Math.round(px * 15) };
        try { core.commitPageLayout(catalog); } catch (error) { setError(error instanceof Error ? error.message : String(error)); }
      },
    }) : undefined;
    return () => { detachMargins?.(); target.replaceChildren(); };
  }, [core, report, page, fit, margins, size.width, size.height, mode, session.readonly, busy]);
  /** 功能：打印/PDF 捕获与预览相同的只读快照；用户确认后才调用外部平台，取消不保存部分结果；时间：2026-11-11；作者：lq。 */
  const confirm = async () => {
    if (!report || busy) return; cancel.current = false; setBusy(true); setError(''); setProgress(undefined);
    try {
      // 功能：预览页移动到了屏幕容器，输出前复位其比例并放回有真实尺寸的离屏捕获容器；时间：2026-11-11；作者：lq。
      const capture = report.surface.root.querySelector<HTMLElement>('[data-core-output-captures]')!;
      if (mode === 'pdf') validateCorePageRange(range, report.pages.length);
      report.pages.forEach(element => { element.style.transform = 'none'; capture.append(element); });
      if (mode === 'print') {
        const ok = await printPages({ pages: report.pages, pageSizes: [...report.frames], widthPx: report.frames[0].widthPx, heightPx: report.frames[0].heightPx, title: session.displayName, fullPaper: true });
        if (!ok) throw new Error('当前平台不能打开打印，请改用导出 PDF。');
        if (mounted.current) setMode(undefined);
      } else {
        const pdf = await exportElementsToPdf({ pages: report.pages, fileName: name, pageRange: range, isCancelled: () => cancel.current || !mounted.current, onProgress: value => { if (mounted.current) setProgress(value); } });
        if (cancel.current || !mounted.current) throw new Error('已取消导出，本次结果未保存。');
        const fileName = name.trim().replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').replace(/\.pdf$/i, '') || stripExtension(session.displayName);
        await getPlatformAdapter().saveBytes(pdf, { suggestedName: `${fileName}.pdf`, forcePicker: true }, 'application/pdf');
        if (mounted.current) { useWorkspaceStore.getState().notify('success', 'PDF 已导出', `${fileName}.pdf`); setMode(undefined); }
      }
    } catch (error) { if (mounted.current) setError(error instanceof Error ? error.message : String(error)); }
    finally { if (mounted.current) setBusy(false); }
  };
  const count = report?.pages.length ?? 0;
  const frame = report?.frames[page]; const section = frame ? core.sections()[frame.section] : undefined;
  /** 功能：预览中方向调整明确选择当前节或整篇，保留每节原纸张尺寸并一次进入历史；时间：2026-11-11；作者：lq。 */
  const orient = (orientation: 'portrait' | 'landscape') => {
    if (session.readonly || busy || !frame) return;
    const catalog = structuredClone(core.pageLayout());
    for (const [index, current] of core.sections().entries()) {
      if (orientationScope !== 'all' && index !== frame.section) continue;
      const settings = current.settings;
      if (orientation !== settings.orientation) catalog.sections[current.lastBlockIndex] = { ...catalog.sections[current.lastBlockIndex], orientation, pageWidth: settings.pageHeight, pageHeight: settings.pageWidth };
    }
    try { core.commitPageLayout(catalog); setError(''); } catch (error) { setError(error instanceof Error ? error.message : String(error)); }
  };
  const elements = <Dialog open={!!mode} title={mode === 'print' ? '打印预览' : '导出为 PDF'} description={mode === 'print' ? `按当前原生版式预览，共 ${count} 页。` : '使用当前原生分页版面生成 PDF，取消不会保存部分结果。'} width={mode === 'print' ? 'print' : 'sm'} onClose={() => { if (!busy) setMode(undefined); }} footer={<><Button disabled={busy} onClick={() => setMode(undefined)}>取消</Button><Button variant="primary" loading={busy} disabled={loading || !count} onClick={() => void confirm()}>{mode === 'print' ? '打印' : '开始导出'}</Button></>}>
    {mode === 'print' ? <div className="grid gap-3 text-xs sm:grid-cols-[minmax(0,1fr)_212px]">
      <div className="min-w-0 overflow-hidden rounded border border-[var(--mo-line)] bg-[var(--mo-surface-2)]"><div ref={host} aria-label="文档打印预览" className="h-[calc(100dvh-250px)] min-h-[280px] overflow-auto" /><div className="flex items-center justify-center gap-3 border-t border-[var(--mo-line)] p-2"><Button size="sm" disabled={busy || page <= 0} onClick={() => setPage(page - 1)}>上一页</Button><span>{count ? `${page + 1} / ${count}` : '0 / 0'}</span><Button size="sm" disabled={busy || page >= count - 1} onClick={() => setPage(page + 1)}>下一页</Button></div></div>
      <div className="space-y-4 rounded border border-[var(--mo-line)] p-3"><Field label="纸张尺寸">{frame ? `${Math.round(frame.widthPx * 25.4 / 96)} × ${Math.round(frame.heightPx * 25.4 / 96)} mm` : '正在计算…'}</Field><Field label="总页数">{count} 页</Field>
        <Field label="预览比例"><Select aria-label="打印预览比例" value={fit} onChange={event => setFit(event.target.value as PrintPreviewFit)}><option value="page">整页</option><option value="width">页宽</option><option value="actual">100%</option></Select></Field>
        <Field label="方向应用范围"><Select aria-label="打印方向应用范围" disabled={busy || session.readonly} value={orientationScope} onChange={event => setOrientationScope(event.target.value as 'section' | 'all')}><option value="section">当前节</option><option value="all">整篇文档</option></Select></Field>
        <Field label="纸张方向"><Select aria-label="打印纸张方向" disabled={busy || session.readonly || loading} value={section?.settings.orientation ?? 'portrait'} onChange={event => orient(event.target.value as 'portrait' | 'landscape')}><option value="portrait">纵向</option><option value="landscape">横向</option></Select></Field>
        <label className="flex items-center gap-2"><input type="checkbox" aria-label="显示页边距" checked={margins} disabled={busy} onChange={event => setMargins(event.target.checked)} />显示页边距{session.readonly ? '' : '（可拖动）'}</label>
        <p className="leading-5 text-[var(--mo-text-muted)]">每页保留所属节的纸张及边距。预览、打印和 PDF 使用同一份已完成排版的页面。</p>
        {section ? <p>当前第 {frame!.section + 1} 节 · {section.settings.pageWidth > section.settings.pageHeight ? '横向' : '纵向'}</p> : null}
      </div></div> : <div className="space-y-3"><Field label="文件名"><Input aria-label="PDF 文件名" value={name} disabled={busy} onChange={event => setName(event.target.value)} /></Field><Field label="页码范围" hint={`留空导出全部 ${count} 页，例如 1-3,5`}><Input aria-label="PDF 页码范围" value={range} disabled={busy} onChange={event => setRange(event.target.value)} /></Field>{progress ? <div className="space-y-2"><ProgressBar value={progress.current / Math.max(1, progress.total) * 100} /><p role="status" className="text-xs">{progress.message}</p></div> : null}{busy ? <Button onClick={() => { cancel.current = true; }}>取消导出</Button> : null}</div>}
    {loading ? <p role="status" className="py-3 text-sm">正在准备字体与分页版式…</p> : null}{error ? <p role="alert" className="py-3 text-sm text-[var(--mo-danger)]">{error}</p> : null}
  </Dialog>;
  return { openPrint: () => open('print'), openExport: () => open('pdf'), elements };
}
