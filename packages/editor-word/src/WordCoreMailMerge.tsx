/** 功能：现有邮件合并 UI 直接操作原生字段与保存快照，预览/Word/ZIP/PDF/打印均不回退旧模型；时间：2026-11-11；作者：lq。 */
import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import JSZip from 'jszip';
import { useWorkspaceStore, type DocumentSession } from '@mini-office/document-model';
import { getPlatformAdapter, printPages } from '@mini-office/platform-adapter';
import { stripExtension } from '@mini-office/shared';
import { exportElementsToPdf } from '@mini-office/export-service';
import { MailMergePanelView } from './MailMergePanel';
import { MergePreviewBar } from './MailMergePreview';
import { useMailMergeController, type MailMergeController, type NativeMailMergeTemplate } from './mail-merge-controller';
import { buildWordMailMergeXml, fileUriTarget } from './mail-merge-ooxml';
import { mergeFileName } from './mail-merge';
import { wordMergeFields } from '../../word-core/src/editing/mail-merge';
import { inspectWordMergePackage, mergeWordPackage, type WordMergePackageInfo } from '../../word-core/src/serialization/mail-merge';
import { coreWordOutputBytes, mountCoreWordReport, type CoreWordReport } from './core-output';

interface MergeView { panel: ReactNode; bar: ReactNode; canvas: ReactNode; previewOn: boolean; }
const EMPTY: MergeView = { panel: null, bar: null, canvas: null, previewOn: false };
interface Props { session: DocumentSession; open: boolean; onClose(): void; onInsertField(name: string): void; children(view: MergeView): ReactNode; }
/** 功能：只在原包配置完成读取后创建控制器，避免异步初值被空白面板状态覆盖；时间：2026-11-11；作者：lq。 */
export function WordCoreMailMerge(props: Props) {
  const [info, setInfo] = useState<WordMergePackageInfo>(); const [error, setError] = useState('');
  useEffect(() => {
    let active = true;
    void inspectWordMergePackage(props.session.wordCore!.parsed.internal.originalBytes).then(info => { if (active) setInfo(info); }).catch(error => { if (active) setError(String(error)); });
    return () => { active = false; };
  }, [props.session.wordCore]);
  if (!info) return props.children({ ...EMPTY, panel: props.open ? <aside aria-label="邮件合并面板" className="w-[350px] shrink-0 border-l border-[var(--mo-line)] bg-[var(--mo-surface)] p-4 text-sm">{error || '正在读取模板字段…'}<button className="ml-3 underline" onClick={props.onClose}>收起</button></aside> : null });
  return <NativeMailMergeContent {...props} info={info} />;
}
/** 功能：所有输出先捕获一个确定版本，取消/读取/预览不写源模板，分文件命名沿用现有规则；时间：2026-11-11；作者：lq。 */
async function generateNativeReports(session: DocumentSession, request: Parameters<NativeMailMergeTemplate['generate']>[0]): Promise<string> {
  const { records, namePattern, isCancelled, onProgress } = request; const check = () => { if (isCancelled()) throw new Error('已取消生成，本次结果未保存。'); };
  check(); const bytes = await coreWordOutputBytes(session.wordCore!); check();
  const info = await inspectWordMergePackage(bytes); const hasNext = info.nextBlocks > 0;
  const output = hasNext ? request.output === 'docx-zip' ? 'docx' : request.output === 'pdf-zip' ? 'pdf' : request.output : request.output;
  const name = `${stripExtension(session.displayName)}_合并报告`;
  const adapter = getPlatformAdapter(); const reports: CoreWordReport[] = [];
  try {
    if (output === 'docx') {
      const merged = await mergeWordPackage(bytes, records.map(record => record.values), { isCancelled, onProgress: (done, total) => onProgress(Math.round(done / total * 90), `正在合并 ${done} / ${total} 条记录…`) }); check();
      const id = await useWorkspaceStore.getState().openBytes({ name: `${name}.docx`, bytes: merged, generated: true });
      if (!id) throw new Error('合并文档打开失败，源模板保留。');
      return `已将 ${records.length} 条记录合并为新 Word 文档，可继续编辑并保存。`;
    }
    if (output === 'docx-zip' || output === 'pdf-zip') {
      const zip = new JSZip(); let size = 0;
      for (let index = 0; index < records.length; index++) {
        check(); onProgress(Math.round(index / records.length * 90), `正在生成 ${index + 1} / ${records.length} 份报告…`);
        const merged = await mergeWordPackage(bytes, [records[index].values], { isCancelled }); check();
        let result = merged;
        if (output === 'pdf-zip') {
          const report = await mountCoreWordReport(merged, isCancelled); reports.push(report);
          try { result = await exportElementsToPdf({ pages: report.pages, fileName: name, isCancelled }); } finally { report.dispose(); reports.pop(); }
        }
        check(); size += result.byteLength; if (size > 200 * 1024 * 1024) throw new Error('报告文件合计超过 200 MB，请减少本批记录数。');
        zip.file(`${mergeFileName(namePattern, records[index], index)}.${output === 'docx-zip' ? 'docx' : 'pdf'}`, result);
      }
      const archive = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }, () => check()); check();
      await adapter.saveBytes(archive, { suggestedName: `${name}.zip`, forcePicker: true }, 'application/zip');
      return `已保存 ${records.length} 份报告的 ZIP 文件。`;
    }
    const merged = await mergeWordPackage(bytes, records.map(record => record.values), { isCancelled, onProgress: (done, total) => onProgress(Math.round(done / total * 40), `正在合并 ${done} / ${total} 条记录…`) }); check();
    onProgress(45, '正在计算报告版式…'); const report = await mountCoreWordReport(merged, isCancelled); reports.push(report); check();
    if (output === 'print') {
      const ok = await printPages({ pages: report.pages, pageSizes: [...report.frames], widthPx: report.frames[0].widthPx, heightPx: report.frames[0].heightPx, title: name, fullPaper: true });
      if (!ok) throw new Error('当前平台不能打开打印，请选择合并 PDF。');
      return `已准备 ${records.length} 条记录、${report.pages.length} 页报告，请在系统打印面板中确认。`;
    }
    const pdf = await exportElementsToPdf({ pages: report.pages, fileName: name, isCancelled, onProgress: progress => onProgress(45 + Math.round(progress.current / Math.max(1, progress.total) * 50), progress.message) }); check();
    await adapter.saveBytes(pdf, { suggestedName: `${name}.pdf`, forcePicker: true }, 'application/pdf');
    return `已将 ${records.length} 条记录、${report.pages.length} 页报告保存为 PDF。`;
  } finally { reports.forEach(report => report.dispose()); }
}
/** 功能：正文只读取轻量字段摘要；数据源、筛选、命名和范围使用现有单一控制器；时间：2026-11-11；作者：lq。 */
function NativeMailMergeContent({ info, ...props }: Props & { info: WordMergePackageInfo }) {
  const core = props.session.wordCore!; const summary = wordMergeFields(core.editor.state.doc, core.parsed);
  const fields = useMemo(() => [...new Set([...info.staticAuxiliaryFields, ...summary.fields])], [info.staticAuxiliaryFields, summary.fields.join('\u0000')]);
  const native = useMemo<NativeMailMergeTemplate>(() => ({ fields, nextBlocks: summary.nextBlocks, panelState: info.panelState,
    persistPanel: (panelState, source) => core.updateMergeSettings({ panelState, source: source ? { xml: buildWordMailMergeXml(source), target: fileUriTarget(source.dataSourcePath) } : source }),
    generate: request => generateNativeReports(props.session, request),
  }), [core, fields, summary.nextBlocks, info, props.session.id, props.session.displayName]);
  const controller = useMailMergeController({ session: props.session, native });
  return props.children({
    panel: <MailMergePanelView open={props.open && !props.session.readonly} controller={controller} onClose={props.onClose} onInsertField={props.onInsertField} />,
    bar: controller.previewOn ? <MergePreviewBar controller={controller} /> : null,
    canvas: controller.previewOn ? <NativeMergePreview core={core} controller={controller} /> : null,
    previewOn: controller.previewOn,
  });
}
/** 功能：每条预览生成独立只读原生会话，切换记录/退出时清理；字段缓存和源模板完全隔离；时间：2026-11-11；作者：lq。 */
function NativeMergePreview({ core, controller: c }: { core: NonNullable<DocumentSession['wordCore']>; controller: MailMergeController }) {
  const host = useRef<HTMLDivElement>(null); const [error, setError] = useState(''); const [loading, setLoading] = useState(true);
  const records = c.hasNext ? c.previewRecords.slice(Math.floor(c.previewIndex / c.pageSize) * c.pageSize, Math.floor(c.previewIndex / c.pageSize) * c.pageSize + c.pageSize) : c.current ? [c.current] : [];
  const signature = JSON.stringify(records.map(record => [record.rowNumber, record.values]));
  // 功能：缩放/窗口尺寸变化只调整已排版页面，不重新生成快照或抢占保存队列；时间：2026-11-11；作者：lq。
  const zoom = useRef(c.previewZoom); zoom.current = c.previewZoom;
  const repaint = useRef<() => void>();
  useEffect(() => { const observer = new ResizeObserver(() => repaint.current?.()); if (host.current) observer.observe(host.current); return () => observer.disconnect(); }, []);
  useEffect(() => { repaint.current?.(); }, [c.previewZoom]);
  useEffect(() => {
    let cancelled = false, report: CoreWordReport | undefined; const target = host.current!; setError(''); setLoading(true); target.replaceChildren();
    const run = async () => {
      if (!records.length || c.issues.length) { setLoading(false); return; }
      const bytes = await coreWordOutputBytes(core); if (cancelled) return;
      const merged = await mergeWordPackage(bytes, records.map(record => record.values), { isCancelled: () => cancelled });
      report = await mountCoreWordReport(merged, () => cancelled); if (cancelled) { report.dispose(); return; }
      report.pages.forEach(page => { const wrapper = document.createElement('div'); page.style.transformOrigin = 'top left'; wrapper.append(page); target.append(wrapper); });
      repaint.current = () => {
        if (!report || cancelled) return;
        const scale = zoom.current === 'fit' ? Math.max(0.2, Math.min(1, (target.clientWidth - 32) / Math.max(...report.frames.map(frame => frame.widthPx)))) : Number(zoom.current);
        report.pages.forEach((page, index) => { const frame = report!.frames[index]; page.parentElement!.style.cssText = `width:${frame.widthPx * scale}px;height:${frame.heightPx * scale}px;margin:0 auto 20px;`; page.style.transform = `scale(${scale})`; });
      };
      repaint.current();
      setLoading(false);
    };
    void run().catch(error => { if (!cancelled) { setLoading(false); setError(error instanceof Error ? error.message : String(error)); } });
    return () => { cancelled = true; repaint.current = undefined; report?.dispose(); target.replaceChildren(); };
  }, [core, signature, c.issues.length]);
  return <div data-mail-merge-canvas className="h-full overflow-auto p-6"><p className="mb-3 text-center text-xs text-[var(--mo-text-muted)]">合并预览 · 数据第 {c.current?.rowNumber ?? '—'} 行</p>{loading ? <p role="status" className="text-center text-sm">正在计算报告分页…</p> : null}{error ? <p role="alert" className="text-sm text-[var(--mo-danger)]">{error}</p> : null}<div ref={host} data-core-merge-preview /></div>;
}
