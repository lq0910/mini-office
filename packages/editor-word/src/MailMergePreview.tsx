// 预览模式的主画布（真实合并文档）与工具栏下方控制条（记录导航、缩放、批量输出）
import { useEffect, useRef, useState } from 'react';
import type { DocModel } from '@mini-office/document-model';
import { Button, Select } from '@mini-office/ui';
import { MergeReport } from './MergeReport';
import type { MailMergeController } from './mail-merge-controller';

/** 预览模式的主画布——渲染当前选中记录的真实合并文档 */
export function MailMergeCanvas({ controller: c, doc }: { controller: MailMergeController; doc: DocModel }) {
  const ref = useRef<HTMLDivElement>(null);
  const [width, setWidth] = useState(800);
  useEffect(() => {
    const element = ref.current;
    if (!element) return;
    const observer = new ResizeObserver(() => setWidth(element.clientWidth));
    observer.observe(element);
    return () => observer.disconnect();
  }, []);
  const scale = c.previewZoom === 'fit' ? Math.max(0.2, Math.min(1, (width - 24) / doc.page.widthPx)) : Number(c.previewZoom);
  return (
    <div ref={ref} data-mail-merge-canvas className="mo-no-print relative min-w-0 flex-1 overflow-auto bg-[var(--mo-canvas)] p-6">
      <div className="mx-auto w-fit">
        <p className="mb-2 text-center text-xs text-[var(--mo-text-muted)]">合并预览 · 数据第 {c.current?.rowNumber ?? '—'} 行</p>
        {c.previewDoc && c.renderBlocks ? <MergeReport doc={c.previewDoc} renderBlocks={c.renderBlocks} scale={scale} /> : <p className="p-6 text-xs">请选择记录，并检查模板字段是否与数据列对应。</p>}
      </div>
    </div>
  );
}

/** 预览模式的工具栏下方控制条——记录导航、缩放、批量输出与退出 */
export function MergePreviewBar({ controller: c }: { controller: MailMergeController }) {
  const action = c.output === 'print' ? '打印' : '生成';
  // 功能：按钮文案跟随合并范围，避免提示的份数与实际输出不一致；时间：2026-10-30；作者：lq。
  const buttonLabel = c.mergeRange === 'all'
    ? `${action}选中 ${c.chosen.length} 份报告`
    : c.mergeRange === 'currentRecord'
      ? `${action}当前记录`
      : `${action}第 ${c.rangeStart}~${c.rangeEnd} 条（${c.rangedRecords.length} 条记录）`;
  // 功能：预览导航单位改为“页”——NEXT 模板一页就是一页模板（每页 N 条记录），普通模板一份报告即一页；时间：2026-10-30；作者：lq。
  const perPage = c.blocksPerPage > 0 ? c.blocksPerPage : 1;
  const total = c.previewRecords.length;
  const pageCount = c.blocksPerPage > 0 ? Math.max(1, Math.ceil(total / perPage)) : total;
  const pageIndex = Math.floor(c.previewIndex / perPage);
  const pageStart = pageIndex * perPage;
  const pageFirst = c.previewRecords[pageStart];
  const pageLast = Math.min(total, pageStart + perPage);
  return (
    <div data-mail-merge-bar className="mo-no-print flex flex-wrap items-center gap-2 border-b border-[var(--mo-line)] bg-[var(--mo-surface)] px-3 py-2 text-xs">
      <Button size="sm" disabled={c.busy || pageIndex <= 0} onClick={() => c.setPreviewIndex(() => Math.max(0, pageStart - perPage))}>上一页</Button>
      {/* 功能：NEXT 模板按页显示——页码 + 本页第几条，与实际出纸的一页对应；时间：2026-10-30；作者：lq。 */}
      {c.blocksPerPage > 0
        ? <span>第 {pageIndex + 1} / {pageCount} 页 · 本页第 {pageStart + 1}~{pageLast} 条 · 数据第 {pageFirst?.rowNumber ?? '—'} 行</span>
        : <span>第 {total ? Math.min(c.previewIndex + 1, total) : 0} / {total} 页 · 数据第 {c.current?.rowNumber ?? '—'} 行</span>}
      <Button size="sm" disabled={c.busy || pageIndex + 1 >= pageCount} onClick={() => c.setPreviewIndex(() => Math.min(Math.max(0, total - 1), pageStart + perPage))}>下一页</Button>
      <Select aria-label="报告预览缩放" className="!w-[110px]" value={c.previewZoom} onChange={(event) => c.setPreviewZoom(event.target.value)}><option value="fit">适合宽度</option><option value="0.5">50%</option><option value="0.75">75%</option><option value="1">100%</option></Select>
      <span className="flex-1" />
      {c.busy ? <>
        <p role="status">{c.status || '处理中…'}</p>
        <Button size="sm" onClick={c.cancelGeneration}>取消生成</Button>
      </> : <Button variant="primary" disabled={c.busy} onClick={() => void c.generate()}>{buttonLabel}</Button>}
      <Button size="sm" variant="ghost" disabled={c.busy} onClick={() => c.setPreviewOn(false)}>返回编辑模板</Button>
      {c.error ? <p role="alert" className="basis-full text-[var(--mo-danger)]">{c.error}</p> : null}
    </div>
  );
}
