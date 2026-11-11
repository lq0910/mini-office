// 邮件合并侧栏——纯 UI；状态与动作全部来自 useMailMergeController（外部传入或本地创建）；筛选浮层可拖动/缩放（居中打开）。
import { useEffect, useRef, useState, type PointerEvent as ReactPointerEvent } from 'react';
import { createPortal } from 'react-dom';
import { type DocModel, type WorkbookModel, type WordSession } from '@mini-office/document-model';
import { Button, Field, Input, ProgressBar, Select } from '@mini-office/ui';
import type { MergeBlockRenderer } from './MergeReport';
import { MAX_MERGE_BATCH, OPERATORS, useMailMergeController, type MailMergeController } from './mail-merge-controller';

/** 筛选浮层的位置与尺寸 */
type FilterBox = { x: number; y: number; w: number; h: number };
/** 缩放方向：四边与四角 */
type ResizeDir = 'n' | 's' | 'e' | 'w' | 'ne' | 'nw' | 'se' | 'sw';
/** 浮层最小尺寸：放宽到能缩到用户想要的任意大小 */
const MIN_FILTER_W = 380;
const MIN_FILTER_H = 220;

export function MailMergePanel({ open, session, doc, onClose, onInsertField, renderBlocks, controller }: {
  open: boolean; session: WordSession; doc: DocModel; onClose: () => void;
  onInsertField: (field: string) => void; renderBlocks: MergeBlockRenderer; controller?: MailMergeController;
}) {
  const own = useMailMergeController({
    session, doc, renderBlocks,
    // 外部已传入 controller 时，内部兜底实例只提供可用的动作，不持久化面板状态，避免与外部实例互相覆盖。
    persist: !controller,
  });
  const c = controller ?? own;
  return <MailMergePanelView open={open} controller={c} onClose={onClose} onInsertField={onInsertField} />;
}

/** 功能：纯面板可由原生 Word 会话直接驱动，不额外挂载旧模型控制器；时间：2026-11-11；作者：lq。 */
export function MailMergePanelView({ open, controller: c, onClose, onInsertField }: {
  open: boolean; controller: MailMergeController; onClose: () => void; onInsertField: (field: string) => void;
}) {
  const [filterModal, setFilterModal] = useState(false);
  // 功能：筛选浮层的位置与尺寸，标题栏可拖动、四边四角可缩放；时间：2026-10-28；作者：lq。
  const [filterBox, setFilterBox] = useState<FilterBox | null>(null);
  const dragRef = useRef<{ dx: number; dy: number } | null>(null);
  const resizeRef = useRef<{ start: FilterBox; dir: ResizeDir } | null>(null);
  // 功能：浮层默认在软件窗口正中打开；双击右下角抓手可恢复默认大小；时间：2026-10-28；作者：lq。
  const defaultFilterBox = (): FilterBox => {
    const w = Math.max(MIN_FILTER_W, Math.min(window.innerWidth - 80, 1080));
    const h = Math.max(MIN_FILTER_H, Math.min(window.innerHeight - 160, 540));
    return {
      x: Math.max(8, Math.round((window.innerWidth - w) / 2)),
      y: Math.max(8, Math.round((window.innerHeight - h) / 2)),
      w,
      h,
    };
  };
  useEffect(() => {
    if (!filterModal) return;
    setFilterBox((previous) => previous ?? defaultFilterBox());
  }, [filterModal]);
  // 功能：软件窗口尺寸变化后把浮层收回可视范围，避免拖到屏幕外找不到；时间：2026-10-30；作者：lq。
  useEffect(() => {
    if (!filterModal) return;
    const onViewportResize = () => setFilterBox((previous) => previous ? {
      ...previous,
      x: Math.min(previous.x, Math.max(8, window.innerWidth - previous.w - 8)),
      y: Math.min(previous.y, Math.max(8, window.innerHeight - previous.h - 8)),
      w: Math.min(previous.w, Math.max(MIN_FILTER_W, window.innerWidth - 16)),
      h: Math.min(previous.h, Math.max(MIN_FILTER_H, window.innerHeight - 16)),
    } : previous);
    window.addEventListener('resize', onViewportResize);
    return () => window.removeEventListener('resize', onViewportResize);
  }, [filterModal]);

  const onFilterTitlePointerDown = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (!filterBox || (event.target as HTMLElement).closest('button')) return;
    const rect = event.currentTarget.parentElement?.getBoundingClientRect();
    if (!rect) return;
    dragRef.current = { dx: event.clientX - rect.left, dy: event.clientY - rect.top };
    event.currentTarget.setPointerCapture(event.pointerId);
  };
  const onFilterTitlePointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    const drag = dragRef.current;
    if (!drag || !filterBox) return;
    const x = Math.min(Math.max(0, event.clientX - drag.dx), Math.max(0, window.innerWidth - filterBox.w - 8));
    const y = Math.min(Math.max(0, event.clientY - drag.dy), Math.max(0, window.innerHeight - filterBox.h - 8));
    setFilterBox({ ...filterBox, x, y });
  };
  const onFilterPointerEnd = (event: ReactPointerEvent<HTMLDivElement>) => {
    dragRef.current = null;
    resizeRef.current = null;
    event.currentTarget.releasePointerCapture?.(event.pointerId);
  };
  // 功能：四边四角都能拖动缩放——左边/上边拖动时同步移动位置，右边/下边只改尺寸；时间：2026-10-30；作者：lq。
  const onFilterResizePointerDown = (dir: ResizeDir) => (event: ReactPointerEvent<HTMLDivElement>) => {
    if (!filterBox) return;
    resizeRef.current = { start: { ...filterBox }, dir };
    event.currentTarget.setPointerCapture(event.pointerId);
    event.preventDefault();
  };
  const onFilterResizePointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    const state = resizeRef.current;
    if (!state) return;
    const { start, dir } = state;
    const dx = event.clientX - start.x;
    const dy = event.clientY - start.y;
    let { x, y, w, h } = start;
    if (dir.includes('e')) w = Math.min(Math.max(MIN_FILTER_W, start.w + dx), Math.max(MIN_FILTER_W, window.innerWidth - start.x - 8));
    if (dir.includes('s')) h = Math.min(Math.max(MIN_FILTER_H, start.h + dy), Math.max(MIN_FILTER_H, window.innerHeight - start.y - 8));
    if (dir.includes('w')) {
      const right = start.x + start.w;
      x = Math.min(Math.max(8, start.x + dx), right - MIN_FILTER_W);
      w = right - x;
    }
    if (dir.includes('n')) {
      const bottom = start.y + start.h;
      y = Math.min(Math.max(8, start.y + dy), bottom - MIN_FILTER_H);
      h = bottom - y;
    }
    setFilterBox({ x, y, w, h });
  };
  // 功能：窄窗口时筛选条件与记录表改为上下排布，保证小尺寸下仍能看清内容；时间：2026-10-30；作者：lq。
  const narrowFilter = !!filterBox && filterBox.w < 680;
  if (!open) return null;
  return (<>
    <aside className="mo-no-print flex w-[350px] shrink-0 flex-col overflow-y-auto border-l border-[var(--mo-line)] bg-[var(--mo-surface)]" aria-label="邮件合并面板">
      <div className="flex items-center justify-between border-b border-[var(--mo-line)] p-3"><strong className="text-sm">邮件合并 / 批量报告</strong><Button size="sm" variant="ghost" onClick={onClose}>收起</Button></div>
      <div className="space-y-5 p-4 text-xs">
        <p className="leading-5 text-[var(--mo-text-muted)]">一行数据生成一份报告。先准备模板，在正文或表格中插入字段，再挑选要打印的记录。</p>
        {c.hasNext ? <p className="leading-5 text-[var(--mo-text-muted)]">检测到 NEXT 模板：每页可连续输出多条记录，预览与输出均为合并后的单份文档。</p> : null}
        <section className="space-y-3"><h3 className="font-semibold">1. 选择数据</h3>
          <Button className="w-full" disabled={c.loading} onClick={() => void c.importDataFile()}>{c.loading ? '正在读取…' : '导入 Excel / CSV'}</Button>
          {c.workbooks.length ? <div className="flex gap-2"><Select aria-label="已打开的数据表" value={c.openWorkbook} onChange={(event) => c.setOpenWorkbook(event.target.value)}><option value="">选择已打开的 Excel</option>{c.workbooks.map((item) => <option key={item.id} value={item.id}>{item.displayName}</option>)}</Select><Button size="sm" disabled={!c.openWorkbook} onClick={() => { const value = c.workbooks.find((item) => item.id === c.openWorkbook); if (value) c.chooseSource(value.data as WorkbookModel, value.displayName, value.path); }}>使用</Button></div> : null}
          {c.source ? <><p className="break-all text-[var(--mo-text-muted)]">数据快照：{c.source.name}</p><Field label="工作表"><Select aria-label="合并工作表" value={c.sheetId} onChange={(event) => c.setSheetId(event.target.value)}>{c.source.workbook.sheets.filter((sheet) => sheet.state === 'visible').map((sheet) => <option key={sheet.id} value={sheet.id}>{sheet.name}</option>)}</Select></Field><Field label="表头所在行"><Input aria-label="表头所在行" type="number" min={1} value={Number.isNaN(c.headerRow) ? '' : c.headerRow} onChange={(event) => c.setHeaderRow(event.target.value === '' ? NaN : Number(event.target.value))} /></Field>{c.data ? <p>{c.data.records.length} 条记录 · {c.data.fields.length} 个字段</p> : null}{c.source.path ? null : <p className="text-[var(--mo-text-muted)]">数据源未保存在磁盘上，保存模板时不会写入 Word / WPS 的邮件合并信息。</p>}</> : <p className="leading-5 text-[var(--mo-text-muted)]">支持 XLS、XLSX 和 CSV，首行通常是字段名称。编号建议使用文本，日期与金额沿用 Excel 格式。</p>}
          {c.result.error ? <p role="alert" className="text-[var(--mo-danger)]">{c.result.error}</p> : null}
        </section>
        <section className="space-y-3 border-t border-[var(--mo-line)] pt-4"><h3 className="font-semibold">2. 在模板中插入字段</h3><p className="leading-5 text-[var(--mo-text-muted)]">将光标放在正文或表格单元格中，点击字段插入。生成时会替换为每行的真实数据。</p><div className="flex flex-wrap gap-2">{c.data?.fields.map((field) => <Button key={field} size="sm" onMouseDown={(event) => event.preventDefault()} onClick={() => onInsertField(field)}>{field}</Button>)}</div>{!c.data ? <p className="text-[var(--mo-text-muted)]">导入数据后显示可用字段。</p> : null}<p>模板已使用：{c.templateFields.length ? c.templateFields.join('、') : '尚未插入字段'}</p></section>
        <section className="space-y-3 border-t border-[var(--mo-line)] pt-4"><h3 className="font-semibold">3. 筛选、预览与批量输出</h3>
          <Button variant="primary" className="w-full" disabled={!c.data || !c.templateFields.length || c.multipleSections || c.busy} onClick={c.togglePreview}>{c.previewOn ? '退出预览，返回编辑模板' : '预览合并效果'}</Button>
          <p className="leading-5 text-[var(--mo-text-muted)]">点击「展开筛选窗口」筛选并勾选记录，主视图实时预览真实分页效果，每批最多 {MAX_MERGE_BATCH} 份。</p>
          <Button className="w-full" variant="secondary" onClick={() => setFilterModal(true)}>展开筛选窗口（可调整大小）</Button>
          <fieldset disabled={c.busy} className="space-y-3">
            <Field label="输出方式"><Select aria-label="合并输出方式" value={c.output} onChange={(event) => c.setOutput(event.target.value as typeof c.output)}><option value="print">直接批量打印</option><option value="pdf">合并为一个 PDF</option><option value="docx">合并为一个 Word 文档</option><option value="docx-zip">每条记录一个 DOCX，打包 ZIP</option><option value="pdf-zip">每条记录一个 PDF，打包 ZIP</option></Select></Field>
            {c.hasNext ? <p className="text-xs text-[var(--mo-text-muted)]">NEXT 模板输出为合并单文档；DOCX/PDF 打包选项会自动退化为单文件。</p> : null}
            <Field label="合并范围"><div className="space-y-2">
              <div className="flex flex-wrap items-center gap-x-4 gap-y-2">
                <label className="flex items-center gap-1.5 text-sm"><input type="radio" name="merge-range" checked={c.mergeRange === 'all'} onChange={() => c.setMergeRange('all')} />全部选中</label>
                <label className="flex items-center gap-1.5 text-sm"><input type="radio" name="merge-range" checked={c.mergeRange === 'currentRecord'} onChange={() => c.setMergeRange('currentRecord')} />当前记录</label>
                <label className="flex items-center gap-1.5 text-sm"><input type="radio" name="merge-range" checked={c.mergeRange === 'recordRange'} onChange={() => c.setMergeRange('recordRange')} />从第
                  <Input aria-label="起始页" type="number" min={1} max={c.totalPages} className="!w-[64px]" value={c.rangeStart} onChange={(event) => c.setRangeStart(Math.max(1, Number(event.target.value) || 1))} disabled={c.mergeRange !== 'recordRange'} />
                  页到第
                  <Input aria-label="结束页" type="number" min={1} max={c.totalPages} className="!w-[64px]" value={c.rangeEnd} onChange={(event) => c.setRangeEnd(Math.max(1, Number(event.target.value) || 1))} disabled={c.mergeRange !== 'recordRange'} />
                  页</label>
              </div>
              <p className="text-[var(--mo-text-muted)]">共 {c.totalPages} 页{c.pageSize > 1 ? ` · 每页 ${c.pageSize} 条记录` : ''} · 页号按勾选后的记录顺序排</p>
              {c.blocksPerPage > 0 && c.rangedRecords.length > 0 ? (() => {
                const perPage = c.blocksPerPage;
                const count = c.rangedRecords.length;
                const pages = Math.ceil(count / perPage);
                const last = count - (pages - 1) * perPage;
                return <p className="text-[var(--mo-text-muted)]">模板每页 {perPage} 个记录块，本次 {count} 条 → 共 {pages} 页{last < perPage ? `，最后一页只有 ${last} 个块有数据、其余 ${perPage - last} 个块为空` : ''}。记录块为空表示本次范围里没有更多记录可填，不会用其他记录补齐。</p>;
              })() : null}
            </div></Field>
          </fieldset>
          {c.busy || c.progress ? <div className="space-y-2">{c.busy ? <p role="status">{c.status || '处理中…'}</p> : null}<ProgressBar value={c.progress} /></div> : null}
          {c.status && !c.busy ? <p role="status" className="leading-5">{c.status}</p> : null}
          {c.error ? <p role="alert" className="text-[var(--mo-danger)]">{c.error}</p> : null}
        </section>
        {c.multipleSections ? <p role="alert" className="text-[var(--mo-danger)]">当前文档含多个节。请使用单份报告模板进行合并；批量成品可直接打开查看和打印。</p> : null}
      </div>
    </aside>
    {filterModal && filterBox ? createPortal(
      <div
        className="fixed z-[90] flex flex-col overflow-hidden rounded-lg border border-[var(--mo-line)] bg-[var(--mo-surface)] shadow-2xl"
        style={{ left: filterBox.x, top: filterBox.y, width: filterBox.w, height: filterBox.h }}
      >
        <div
          className="flex cursor-move select-none items-center justify-between border-b border-[var(--mo-line)] px-3 py-2"
          onPointerDown={onFilterTitlePointerDown}
          onPointerMove={onFilterTitlePointerMove}
          onPointerUp={onFilterPointerEnd}
          onPointerCancel={onFilterPointerEnd}
        >
          <strong className="text-sm">筛选记录（居中打开，可拖标题栏移动、拖四边四角调整大小）</strong>
          <Button size="sm" variant="ghost" onClick={() => setFilterModal(false)}>关闭</Button>
        </div>
        <div className="flex min-h-0 flex-1 gap-4 overflow-hidden p-4" style={narrowFilter ? { flexDirection: 'column' } : undefined}>
          <fieldset disabled={c.busy} className="flex shrink-0 flex-col gap-3 overflow-auto text-sm" style={narrowFilter ? { width: '100%', maxHeight: '45%' } : { width: 320 }}>
            <div className="flex gap-2"><Input aria-label="搜索合并记录" placeholder="搜索所有字段…" value={c.search} onChange={(event) => c.setSearch(event.target.value)} /><Button size="sm" disabled={!c.data?.fields.length || c.conditions.length >= 5} onClick={() => c.setConditions((previous) => [...previous, { field: c.data!.fields[0], op: 'contains', value: '' }])}>添加条件</Button>{c.conditions.length ? <Button size="sm" onClick={() => c.setConditions([])}>清空条件</Button> : null}</div>
            {c.conditions.map((condition, index) => <div key={index} className="grid grid-cols-[56px_minmax(0,1fr)_96px_minmax(0,1fr)] items-center gap-2">{index > 0 ? <Select aria-label={`条件 ${index + 1} 关系`} className="w-full" value={condition.conjunction ?? 'and'} onChange={(event) => c.setConditions((previous) => previous.map((item, i) => i === index ? { ...item, conjunction: event.target.value as 'and' | 'or' } : item))}><option value="and">且</option><option value="or">或</option></Select> : <span aria-hidden />}<Select aria-label={`条件 ${index + 1} 字段`} className="w-full min-w-0" value={condition.field} onChange={(event) => c.setConditions((previous) => previous.map((item, i) => i === index ? { ...item, field: event.target.value } : item))}>{c.data?.fields.map((field) => <option key={field}>{field}</option>)}</Select><Select aria-label={`条件 ${index + 1} 比较`} className="w-[96px]" value={condition.op} onChange={(event) => c.setConditions((previous) => previous.map((item, i) => i === index ? { ...item, op: event.target.value as typeof condition.op } : item))}>{OPERATORS.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}</Select><Input aria-label={`条件 ${index + 1} 值`} value={condition.value} disabled={condition.op === 'empty' || condition.op === 'notEmpty'} onChange={(event) => c.setConditions((previous) => previous.map((item, i) => i === index ? { ...item, value: event.target.value } : item))} /></div>)}
            {c.conditions.length ? <p className="text-[var(--mo-text-muted)]">条件之间可切换「且 / 或」；大于/小于按原始数值比较；包含/不包含不区分大小写。</p> : null}
            <div className="flex flex-wrap items-center gap-2"><Select aria-label="记录排序字段" className="!w-[120px]" value={c.sort} onChange={(event) => c.setSort(event.target.value)}><option value="">按源数据顺序</option>{c.data?.fields.map((field) => <option key={field}>{field}</option>)}</Select><Button size="sm" disabled={!c.sort} onClick={() => c.setDescending((value) => !value)}>{c.descending ? '降序' : '升序'}</Button><Button size="sm" onClick={() => c.setSelected(new Set(c.filtered.map((record) => record.rowNumber)))}>全选筛选结果</Button><Button size="sm" onClick={() => c.setSelected(new Set())}>清空选择</Button>{c.filtered.length > MAX_MERGE_BATCH ? <Button size="sm" onClick={() => c.setSelected(new Set(c.filtered.slice(0, MAX_MERGE_BATCH).map((record) => record.rowNumber)))}>选择前 {MAX_MERGE_BATCH} 条</Button> : null}</div>
            <p>当前选中 {c.chosen.length} / 筛选结果 {c.filtered.length} / 全部 {c.data?.records.length ?? 0} 条</p>
            {(c.output === 'docx-zip' || c.output === 'pdf-zip') ? <Field label="文件命名" hint="支持 {序号} 和 {字段名}。自动附加序号，避免重名覆盖。"><Input aria-label="报告文件命名" maxLength={150} value={c.namePattern} onChange={(event) => c.setNamePattern(event.target.value)} /></Field> : null}
            {c.invalidName ? <p role="alert" className="text-[var(--mo-danger)]">文件命名中的字段不存在：{c.badNameFields.join("、")}</p> : null}
            {c.issues.length ? <div role="alert" className="rounded-md bg-[var(--mo-surface-2)] p-3 text-[var(--mo-danger)]"><p>请修正数据或取消选择有问题的记录：</p>{c.issues.slice(0, 5).map((issue, index) => <p key={index}>{issue}</p>)}{c.issues.length > 5 ? <p>另有 {c.issues.length - 5} 项问题。</p> : null}</div> : null}
          </fieldset>
          <fieldset disabled={c.busy} className="flex min-w-0 flex-1 flex-col gap-2 text-sm">
            <div className="min-h-0 flex-1 overflow-auto rounded-md border border-[var(--mo-line)]"><table className="w-full whitespace-nowrap text-left text-sm"><thead className="sticky top-0 bg-[var(--mo-surface-2)]"><tr><th className="p-2.5">选择</th><th className="p-2.5">源行</th>{c.data?.fields.map((field) => <th key={field} className="p-2.5">{field}</th>)}</tr></thead><tbody>{c.filtered.slice(c.tablePage * 50, c.tablePage * 50 + 50).map((record) => <tr key={record.rowNumber} className="border-t border-[var(--mo-line)]"><td className="p-2.5"><input type="checkbox" aria-label={`选择第 ${record.rowNumber} 行`} checked={c.selected.has(record.rowNumber)} onChange={(event) => c.setSelected((previous) => { const next = new Set(previous); if (event.target.checked) next.add(record.rowNumber); else next.delete(record.rowNumber); return next; })} /></td><td className="p-2.5"><button className="underline" aria-label={`预览第 ${record.rowNumber} 行`} disabled={!c.selected.has(record.rowNumber)} onClick={() => { c.setPreviewIndex(Math.max(0, c.previewRecords.findIndex((item) => item.rowNumber === record.rowNumber))); c.setPreviewOn(true); }}>{record.rowNumber}</button></td>{c.data?.fields.map((field) => <td key={field} title={record.values[field]} className="max-w-[160px] truncate p-2.5">{record.values[field] || <span className="text-[var(--mo-text-faint)]">空</span>}</td>)}</tr>)}</tbody></table>{!c.filtered.length ? <p className="p-4 text-sm">没有符合条件的记录，请调整筛选。</p> : null}</div>
            <div className="flex items-center justify-between"><Button size="sm" disabled={!c.tablePage} onClick={() => c.setTablePage((value) => value - 1)}>上一批记录</Button><span>第 {c.tablePage + 1} / {Math.max(1, Math.ceil(c.filtered.length / 50))} 页</span><Button size="sm" disabled={(c.tablePage + 1) * 50 >= c.filtered.length} onClick={() => c.setTablePage((value) => value + 1)}>下一批记录</Button></div>
            {c.chosen.length > MAX_MERGE_BATCH ? <p role="alert" className="text-[var(--mo-danger)]">本批超过 {MAX_MERGE_BATCH} 份，请缩小筛选范围或减少勾选。</p> : null}
          </fieldset>
        </div>
        {/* 功能：四边与四角各一个拖拽热区，边与角均可调整到想要的尺寸；时间：2026-10-30；作者：lq。 */}
        {(['n', 's', 'e', 'w'] as ResizeDir[]).map((dir) => (
          <div
            key={dir}
            aria-label={`从${({ n: '上', s: '下', e: '右', w: '左' } as Record<string, string>)[dir]}边调整筛选窗口大小`}
            className={`absolute select-none ${({ n: 'left-5 right-5 top-0 h-1.5 cursor-n-resize', s: 'left-5 right-5 bottom-0 h-1.5 cursor-s-resize', e: 'top-5 bottom-5 right-0 w-1.5 cursor-e-resize', w: 'top-5 bottom-5 left-0 w-1.5 cursor-w-resize' } as Record<string, string>)[dir]}`}
            onPointerDown={onFilterResizePointerDown(dir)}
            onPointerMove={onFilterResizePointerMove}
            onPointerUp={onFilterPointerEnd}
            onPointerCancel={onFilterPointerEnd}
          />
        ))}
        {(['nw', 'ne', 'sw', 'se'] as ResizeDir[]).map((dir) => (
          <div
            key={dir}
            aria-label={`从${({ nw: '左上', ne: '右上', sw: '左下', se: '右下' } as Record<string, string>)[dir]}角调整筛选窗口大小`}
            className={`absolute h-5 w-5 select-none ${({ nw: 'left-0 top-0 cursor-nwse-resize', ne: 'right-0 top-0 cursor-nesw-resize', sw: 'bottom-0 left-0 cursor-nesw-resize', se: 'bottom-0 right-0 cursor-nwse-resize' } as Record<string, string>)[dir]}`}
            onPointerDown={onFilterResizePointerDown(dir)}
            onPointerMove={onFilterResizePointerMove}
            onPointerUp={onFilterPointerEnd}
            onPointerCancel={onFilterPointerEnd}
            onDoubleClick={dir === 'se' ? () => setFilterBox(defaultFilterBox()) : undefined}
          />
        ))}
        {/* 功能：右下角画可见抓手，提示此处可拖拽调整大小；时间：2026-10-30；作者：lq。 */}
        <div
          className="pointer-events-none absolute right-1 bottom-1 h-3.5 w-3.5 opacity-45"
          style={{ backgroundImage: 'repeating-linear-gradient(135deg, currentColor 0 1px, transparent 1px 4px)' }}
          aria-hidden
        />
      </div>,
      document.body,
    ) : null}
  </>);
}
