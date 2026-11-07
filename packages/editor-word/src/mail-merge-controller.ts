// 邮件合并全部状态与动作的单一来源，供侧栏、主画布预览与控制条共用
import { useEffect, useMemo, useRef, useState } from 'react';
import JSZip from 'jszip';
import { getSourceBytes, useWorkspaceStore, type DocBlock, type DocModel, type WordSession, type WorkbookModel } from '@mini-office/document-model';
import { getPlatformAdapter, printCurrentView } from '@mini-office/platform-adapter';
import { stripExtension } from '@mini-office/shared';
import { exportElementsToPdf } from '@mini-office/export-service';
import { serializeDocx } from './serialize';
import { mountMergeReport, type MergeBlockRenderer } from './MergeReport';
import { collectMergeFields, countNextBlocks, extractMergeData, filterMergeRecords, hasNextMergeField, MAX_MERGE_BATCH, mergeDocument, mergeFileName, mergeNextDocument, mergeReadProgressText, pageRangeToRecordRange, readMergeWorkbookOffThread, recordsByRange, type MergeCondition, type MergeData, type MergeRangeKind } from './mail-merge';
import { type WordMailMergeInfo } from './mail-merge-ooxml';

export { MAX_MERGE_BATCH } from './mail-merge';

export const OPERATORS: { value: MergeCondition['op']; label: string }[] = [
  { value: 'contains', label: '包含' }, { value: 'eq', label: '等于' }, { value: 'ne', label: '不等于' },
  { value: 'gt', label: '大于' }, { value: 'gte', label: '大于等于' }, { value: 'lt', label: '小于' }, { value: 'lte', label: '小于等于' },
  { value: 'empty', label: '为空' }, { value: 'notEmpty', label: '不为空' },
  { value: 'notContains', label: '不包含' },
];
export type OutputKind = 'print' | 'pdf' | 'docx' | 'docx-zip' | 'pdf-zip';
/** 功能：在同一文档切换标签时保留数据快照、筛选和勾选，不自动刷新源数据；时间：2026-10-24；作者：lq。 */
export type MergeSource = { workbook: WorkbookModel; name: string; path?: string };
interface MergeDraft { source?: MergeSource; sheetId: string; headerRow: number; conditions: MergeCondition[]; search: string; sort: string; descending: boolean; selected: number[]; allowEmpty: boolean; namePattern: string; output: OutputKind; mergeRange?: MergeRangeKind; rangeStart?: number; rangeEnd?: number }
const mergeDrafts = new Map<string, MergeDraft>();

/** 功能：从文档 meta 里读回保存的面板状态（筛选、排序、命名、输出方式、范围），首次打开模板时恢复；时间：2026-10-29；作者：lq。 */
function loadPersistedMailMergeState(doc: DocModel): MergeDraft | undefined {
  const raw = doc.meta?.mailMergeState;
  if (!raw) return undefined;
  try {
    const parsed = JSON.parse(raw) as Partial<MergeDraft>;
    return {
      source: undefined,
      sheetId: parsed.sheetId ?? '',
      headerRow: parsed.headerRow ?? 1,
      conditions: parsed.conditions ?? [],
      search: parsed.search ?? '',
      sort: parsed.sort ?? '',
      descending: parsed.descending ?? false,
      selected: parsed.selected ?? [],
      allowEmpty: parsed.allowEmpty ?? true,
      namePattern: parsed.namePattern ?? '报告_{序号}',
      output: parsed.output ?? 'print',
      mergeRange: parsed.mergeRange,
      rangeStart: parsed.rangeStart,
      rangeEnd: parsed.rangeEnd,
    };
  } catch {
    return undefined;
  }
}

/** 把各条记录生成的文档拼成一个可编辑文档：记录之间用分页符分隔 */
function mergeRecordsIntoDocument(documents: DocModel[]): DocModel {
  const [first, ...rest] = documents;
  if (!rest.length) return first;
  const blocks = [...first.blocks];
  rest.forEach((value, index) => {
    blocks.push({ type: 'pageBreak', id: `merge-record-${index}` });
    blocks.push(...value.blocks);
  });
  return { ...first, blocks };
}

export function useMailMergeController({ session, doc, renderBlocks, onOpenDocument, persist = true }: {
  session: WordSession; doc: DocModel; renderBlocks: MergeBlockRenderer;
  /** 功能：docx 输出时若提供，合并结果直接作为新文档在软件窗口打开，不弹保存框；时间：2026-10-30；作者：lq。 */
  onOpenDocument?: (doc: DocModel, name: string) => void | Promise<void>;
  /** 功能：外部已传入 controller 时内部兜底实例不持久化，避免多实例互相覆盖；时间：2026-10-29；作者：lq。 */
  persist?: boolean;
}) {
  const sessions = useWorkspaceStore((state) => state.sessions);
  const notify = useWorkspaceStore((state) => state.notify);
  const updateDocMeta = useWorkspaceStore((state) => state.updateDocMeta);
  const cached = useRef(mergeDrafts.get(session.id) ?? loadPersistedMailMergeState(doc)).current;
  // 功能：挂载面板只恢复界面，不把默认状态写入原文件或清除原邮件合并关系；时间：2026-11-07；作者：lq。
  const initialPanelState = useRef<string>();
  const panelWasEdited = useRef(false);
  const [source, setSource] = useState<MergeSource | undefined>(cached?.source);
  const [sheetId, setSheetId] = useState(cached?.sheetId ?? '');
  const [headerRow, setHeaderRow] = useState(cached?.headerRow ?? 1);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [openWorkbook, setOpenWorkbook] = useState('');
  // 主画布预览模式开关与缩放，控制条与画布共享
  const [previewOn, setPreviewOn] = useState(false);
  const [previewZoom, setPreviewZoom] = useState('fit');
  const [conditions, setConditions] = useState<MergeCondition[]>(cached?.conditions ?? []);
  const [search, setSearch] = useState(cached?.search ?? '');
  const [sort, setSort] = useState(cached?.sort ?? '');
  const [descending, setDescending] = useState(cached?.descending ?? false);
  const [selected, setSelected] = useState<Set<number>>(new Set(cached?.selected ?? []));
  const [tablePage, setTablePage] = useState(0);
  const [previewIndex, setPreviewIndex] = useState(0);
  // 功能：恢复已有模板的空值策略，避免挂载时更改原配置；时间：2026-11-07；作者：lq。
  const [allowEmpty, setAllowEmpty] = useState(cached?.allowEmpty ?? true);
  const [namePattern, setNamePattern] = useState(cached?.namePattern ?? '报告_{序号}');
  const [output, setOutput] = useState<OutputKind>(cached?.output ?? 'print');
  // 功能：合并范围与 0.1.4 一致，支持全部选中、当前记录和序号区间；时间：2026-10-30；作者：lq。
  const [mergeRange, setMergeRange] = useState<MergeRangeKind>(cached?.mergeRange ?? 'all');
  const [rangeStart, setRangeStart] = useState(cached?.rangeStart ?? 1);
  const [rangeEnd, setRangeEnd] = useState(cached?.rangeEnd ?? 1);
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState('');
  const cancel = useRef(false);
  const printCleanup = useRef<(() => void) | null>(null);
  const workbooks = Object.values(sessions).filter((item) => item.kind === 'xlsx');
  const templateFields = useMemo(() => collectMergeFields(doc), [doc]);
  const result = useMemo<{ data?: MergeData; error?: string }>(() => {
    if (!source) return {};
    try { return { data: extractMergeData(source.workbook, sheetId, headerRow) }; }
    catch (cause) { return { error: cause instanceof Error ? cause.message : String(cause) }; }
  }, [source, sheetId, headerRow]);
  const data = result.data;
  // 功能：按数据对象变化初始化选择，避免重复挂载清空已恢复的勾选；时间：2026-10-24；作者：lq。
  const selectionDataRef = useRef(data);
  const filtered = useMemo(() => {
    const records = filterMergeRecords(data?.records ?? [], conditions, search);
    if (sort) records.sort((a, b) => {
      const left = a.raw[sort], right = b.raw[sort];
      const numeric = left !== null && right !== null && String(left).trim() !== '' && String(right).trim() !== '' && Number.isFinite(Number(left)) && Number.isFinite(Number(right));
      const comparison = numeric ? Number(left) - Number(right) : (a.values[sort] ?? '').localeCompare(b.values[sort] ?? '', 'zh-CN', { numeric: true });
      return (descending ? -1 : 1) * comparison;
    });
    return records;
  }, [data, conditions, search, sort, descending]);
  // 功能：模板每页的记录块数（NEXT 块个数），用于说明输出页数与末尾空块来源；时间：2026-10-28；作者：lq。
  const hasNext = useMemo(() => hasNextMergeField(doc), [doc]);
  const blocksPerPage = useMemo(() => (hasNext ? countNextBlocks(doc) : 0), [hasNext, doc]);
  // 功能：一页容纳的记录数——NEXT 模板按记录块数，普通模板一份报告即一页；时间：2026-10-30；作者：lq。
  const pageSize = blocksPerPage > 0 ? blocksPerPage : 1;
  const chosen = useMemo(() => filtered.filter((record) => selected.has(record.rowNumber)), [filtered, selected]);
  // 功能：输出总页数，按勾选后的记录数与每页记录数换算；时间：2026-10-30；作者：lq。
  const totalPages = Math.max(1, Math.ceil(chosen.length / pageSize));
  // 功能：按合并范围从选中记录中取子集，作为校验、预览与输出的实际记录集合；范围按页表达，页号先换算成记录区间；时间：2026-10-30；作者：lq。
  const rangedRecords = useMemo(() => {
    if (mergeRange !== 'recordRange') return recordsByRange(chosen, mergeRange, previewIndex, rangeStart, rangeEnd);
    const [from, to] = pageRangeToRecordRange(rangeStart, rangeEnd, pageSize);
    return recordsByRange(chosen, 'recordRange', previewIndex, from, to);
  }, [chosen, mergeRange, previewIndex, rangeStart, rangeEnd, pageSize]);
  // 功能：预览导航的记录池——限定范围时只在范围内上下切换，避免预览跳到范围外的记录；时间：2026-10-30；作者：lq。
  const previewRecords = useMemo(() => mergeRange === 'all' ? chosen : rangedRecords, [chosen, rangedRecords, mergeRange]);
  const issues = useMemo(() => rangedRecords.flatMap((record) => [
    ...record.errors,
    ...templateFields.flatMap((field) => !data?.fields.includes(field) ? [`缺少数据列：${field}`] : !allowEmpty && !record.values[field]?.trim() ? [`字段为空：${field}`] : []),
  ].map((message) => `第 ${record.rowNumber} 行：${message}`)), [rangedRecords, templateFields, data, allowEmpty]);
  const current = previewRecords[Math.min(previewIndex, Math.max(0, previewRecords.length - 1))];
  // 功能：NEXT 预览按页取记录窗口——显示当前记录所在的那一页模板（每页几个块就填几条），而不是只填当前一条；时间：2026-10-30；作者：lq。
  const nextPageStart = blocksPerPage > 0 ? Math.floor(previewIndex / blocksPerPage) * blocksPerPage : 0;
  const nextWindow = useMemo(
    () => (blocksPerPage > 0 ? previewRecords.slice(nextPageStart, nextPageStart + blocksPerPage) : []),
    [previewRecords, nextPageStart, blocksPerPage],
  );
  const previewDoc = useMemo(() => {
    if (!templateFields.length) return undefined;
    // 功能：NEXT 模板(每页多条)的预览必须与正式合并一致，使用多块合并；否则仍按当前记录逐条预览；时间：2026-10-28；作者：lq。
    if (hasNextMergeField(doc)) {
      // 功能：NEXT 预览只渲染当前记录所在的那一页模板，不再一次性展开全部合并页；用「上一页 / 下一页」逐页切换；时间：2026-10-30；作者：lq。
      if (!nextWindow.length) return undefined;
      return mergeNextDocument(doc, nextWindow, { allowEmpty });
    }
    if (!current || templateFields.some((field) => !Object.hasOwn(current.values, field))) return undefined;
    return mergeDocument(doc, current);
  }, [current, doc, templateFields, nextWindow, allowEmpty]);
  const multipleSections = doc.blocks.some((block) => block.type === 'paragraph' && block.sectionEnd);
  const badNameFields = [...namePattern.matchAll(/\{([^{}]+)\}/g)].map((match) => match[1]).filter((field) => field !== '序号' && !data?.fields.includes(field));
  // NEXT 模板输出为合并单文档，docx-zip/pdf-zip 自动退化为单文件，不校验逐条命名
  const invalidName = !hasNext && (output === 'docx-zip' || output === 'pdf-zip') && badNameFields.length > 0;
  const canGenerate = !multipleSections && !invalidName && !!rangedRecords.length && rangedRecords.length <= MAX_MERGE_BATCH && !!templateFields.length && !issues.length;
  useEffect(() => {
    if (selectionDataRef.current === data) return;
    selectionDataRef.current = data;
    setSelected(new Set(data?.records.map((record) => record.rowNumber) ?? []));
    setConditions([]); setSearch(''); setSort(''); setPreviewIndex(0); setTablePage(0);
  }, [data]);
  useEffect(() => { setTablePage(0); setPreviewIndex(0); }, [search, conditions, sort, descending]);
  useEffect(() => () => { cancel.current = true; printCleanup.current?.(); }, []);
  // 功能：Word / WPS 兼容信息——只有从本机 xlsx 文件导入、且存在工作表名时才生成；时间：2026-10-29；作者：lq。
  const wordSheetName = useMemo(() => {
    if (!source) return undefined;
    return source.workbook.sheets.find((sheet) => sheet.id === sheetId)?.name
      ?? source.workbook.sheets.find((sheet) => sheet.state === 'visible')?.name
      ?? source.workbook.sheets[0]?.name;
  }, [source, sheetId]);
  const wordCompatible = useMemo<WordMailMergeInfo | undefined>(() => {
    if (!source?.path || !/\.xlsx$/i.test(source.name) || !wordSheetName) return undefined;
    return { dataSourcePath: source.path, sheetName: wordSheetName, fields: data?.fields ?? [], conditions };
  }, [source, wordSheetName, data]);

  useEffect(() => {
    if (!persist) return;
    const payload: MergeDraft = { source, sheetId, headerRow, conditions, search, sort, descending, selected: [...selected], allowEmpty, namePattern, output, mergeRange, rangeStart, rangeEnd };
    mergeDrafts.set(session.id, payload);
    for (const id of mergeDrafts.keys()) if (!sessions[id]) mergeDrafts.delete(id);
    // 同步模板持久化状态：通过 store 更新元数据而不标记脏，避免直接 mutate 丢失未保存标记；时间：2026-10-29；作者：lq。
    const mailMergeState = JSON.stringify({
      conditions, search, sort, descending, allowEmpty, output, mergeRange, rangeStart, rangeEnd,
      dataSource: source ? { name: source.name, path: source.path ?? '' } : undefined,
    });
    const currentMeta = (useWorkspaceStore.getState().sessions[session.id]?.data as DocModel | undefined)?.meta;
    // 功能：尚未重新选择数据源时保留导入模板关系，调整筛选不能把已有 Word 合并配置清空；时间：2026-11-07；作者：lq。
    const mailMergeTemplate = wordCompatible ?? (source ? undefined : currentMeta?.mailMergeTemplate);
    // 功能：仅在用户实际改变配置后持久化，恢复初值时也能同步，避免打开即污染 settings.xml；时间：2026-11-07；作者：lq。
    const signature = JSON.stringify({ mailMergeState, mailMergeTemplate });
    if (initialPanelState.current === undefined) initialPanelState.current = signature;
    if (signature !== initialPanelState.current) panelWasEdited.current = true;
    if (!panelWasEdited.current) return;
    // 以 store 中的最新元数据为准，避免组件持有旧 doc 引用时反复写入；时间：2026-10-29；作者：lq。
    if (currentMeta && (currentMeta.mailMergeState !== mailMergeState || JSON.stringify(currentMeta.mailMergeTemplate) !== JSON.stringify(mailMergeTemplate))) {
      updateDocMeta(session.id, { mailMergeState, mailMergeTemplate });
    }
  }, [session.id, sessions, source, sheetId, headerRow, conditions, search, sort, descending, selected, allowEmpty, namePattern, output, mergeRange, rangeStart, rangeEnd, wordCompatible, persist, updateDocMeta]);
  // 预览模式下按 ESC 快速返回编辑模板
  useEffect(() => {
    if (!previewOn) return;
    const onKey = (event: KeyboardEvent) => { if (event.key === 'Escape') setPreviewOn(false); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [previewOn]);
  const chooseSource = (workbook: WorkbookModel, name: string, path?: string) => {
    // 功能：冻结本次数据快照，后续 Excel 编辑不会悄悄改变预览与输出；时间：2026-10-24；作者：lq。
    setSource({ workbook: structuredClone(workbook), name, path });
    setSheetId(workbook.sheets.find((sheet) => sheet.state === 'visible')?.id ?? workbook.sheets[0]?.id ?? '');
    setHeaderRow(1); setError(''); setStatus('');
  };
  const importFile = async (file: File) => {
    setLoading(true); setError(''); setStatus('正在读取数据文件…');
    try {
      if (file.size > 50 * 1024 * 1024) throw new Error('数据文件不能超过 50 MB。');
      const bytes = new Uint8Array(await file.arrayBuffer());
      const workbook = await readMergeWorkbookOffThread(bytes, file.name, (stage, completed, total) => {
        setStatus(mergeReadProgressText(stage, completed, total));
      });
      chooseSource(workbook, file.name);
      setStatus('数据读取完成');
    } catch (cause) { setError(cause instanceof Error ? cause.message : String(cause)); setStatus(''); }
    finally { setLoading(false); }
  };
  /** 功能：通过系统文件对话框导入数据源，支持旧版 XLS（先经兼容桥转换）；时间：2026-10-29；作者：lq。 */
  const importDataFile = async () => {
    setLoading(true); setError(''); setStatus('正在读取数据文件…');
    try {
      const adapter = getPlatformAdapter();
      const opened = await adapter.openFile(['xlsx', 'xls', 'csv']);
      if (!opened) { setStatus(''); return; }
      if (opened.bytes.byteLength > 50 * 1024 * 1024) throw new Error('数据文件不能超过 50 MB。');
      // 功能：旧版 XLS 先经 Rust 兼容桥转换为 OOXML 再解析，与首页打开旧版文件的路径保持一致；时间：2026-10-30；作者：lq。
      const source = /\.xls$/i.test(opened.name) && opened.path && adapter.convertLegacy
        ? await adapter.convertLegacy(opened.path)
        : opened;
      if (!source) throw new Error('无法转换该 XLS 文件，请在 Excel 中另存为 XLSX 后再导入。');
      const workbook = await readMergeWorkbookOffThread(source.bytes, source.name, (stage, completed, total) => {
        setStatus(mergeReadProgressText(stage, completed, total));
      });
      chooseSource(workbook, opened.name, source === opened ? opened.path : undefined);
      setStatus('数据读取完成');
    } catch (cause) { setError(cause instanceof Error ? cause.message : String(cause)); setStatus(''); }
    finally { setLoading(false); }
  };
  /** 进入/退出主画布预览模式；生成中禁止切换 */
  const togglePreview = () => {
    if (busy) return;
    setError(''); setStatus(''); setProgress(0);
    setPreviewOn((value) => !value);
  };
  const cancelGeneration = () => { cancel.current = true; setStatus('正在取消，请稍候…'); };
  const generate = async () => {
    if (busy) return;
    if (!canGenerate) {
      if (multipleSections) setError('当前文档含多个节，无法合并，请改用单节模板。');
      else if (invalidName) setError(`文件命名中的字段不存在：${badNameFields.join('、')}`);
      else if (!rangedRecords.length) setError('当前范围没有可生成的记录。');
      else if (rangedRecords.length > MAX_MERGE_BATCH) setError(`本批 ${rangedRecords.length} 份超过单次上限 ${MAX_MERGE_BATCH} 份，请缩小范围。`);
      else if (!templateFields.length) setError('模板尚未插入合并字段，请先在正文或表格中插入字段。');
      else if (issues.length) setError('部分记录存在数据问题，请修正或取消勾选有问题的记录（见面板提示）。');
      return;
    }
    cancel.current = false; setBusy(true); setProgress(0); setError('');
    printCleanup.current?.(); printCleanup.current = null;
    const container = document.createElement('div');
    container.style.cssText = 'position:fixed;left:-100000px;top:0;background:white;color:black;';
    document.body.appendChild(container);
    const mounts: ReturnType<typeof mountMergeReport>[] = [];
    const cleanup = () => { mounts.forEach((mount) => mount.dispose()); container.remove(); };
    let keepForPrint = false;
    try {
      const records = rangedRecords;
      const template = structuredClone(doc);
      const sourceBytes = getSourceBytes(session.id)?.slice();
      const zip = new JSZip();
      const documents: DocModel[] = [];
      let archiveBytes = 0;
      const addArchiveFile = (name: string, bytes: Uint8Array) => { archiveBytes += bytes.byteLength; if (archiveBytes > 200 * 1024 * 1024) throw new Error('报告文件合计超过 200 MB，请减少本批记录数。'); zip.file(name, bytes); };
      const allPages: HTMLElement[] = [];
      const baseName = `${stripExtension(session.displayName)}_合并报告`;
      // NEXT 模板：单文档一次合并所有记录（一页多条）；普通模板保持每条记录独立文档
      const hasNext = hasNextMergeField(template);
      // 功能：NEXT 模板不支持按记录分文件，docx-zip/pdf-zip 退化为合并单文件输出；时间：2026-10-29；作者：lq。
      const effectiveOutput = hasNext && (output === 'docx-zip' || output === 'pdf-zip')
        ? (output === 'docx-zip' ? 'docx' : 'pdf')
        : output;
      if (hasNext && (output === 'docx-zip' || output === 'pdf-zip')) {
        setStatus('NEXT 模板输出为合并单文档，已自动切换为单文件输出。');
      }
      if (hasNext) {
        // 功能：NEXT 单文档：一次合并所有记录，一页连续多条；时间：2026-10-28；作者：lq。
        setStatus(`正在合并 ${records.length} 条记录为 1 份报告…`);
        const merged = mergeNextDocument(template, records, { allowEmpty });
        const mount = mountMergeReport(merged, renderBlocks, container); mounts.push(mount);
        let layout;
        try { layout = await mount.ready; } catch (cause) { throw new Error(`合并版式计算失败：${cause instanceof Error ? cause.message : String(cause)}`); }
        if (cancel.current) throw new Error('已取消生成，本次结果未保存。');
        // 功能：将实测的长表分页结果同步给 DOCX，成品与预览使用同一块序列；时间：2026-10-24；作者：lq。
        const pageBlocks: DocBlock[] = [];
        layout.slices.forEach((slice, pageIndex) => { if (pageIndex) pageBlocks.push({ type: 'pageBreak', id: `merge-break-${pageIndex}` }); pageBlocks.push(...slice.blocks); });
        const paged = { ...merged, blocks: pageBlocks };
        documents.push(paged);
        if (effectiveOutput === 'print' || effectiveOutput === 'pdf') {
          allPages.push(...layout.pages);
          if (allPages.length > 500) throw new Error('合并结果超过 500 页，请缩小筛选范围后分批输出。');
        } else { mount.dispose(); mounts.pop(); }
        setProgress(85);
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      } else {
        for (let index = 0; index < records.length; index++) {
          if (cancel.current) throw new Error('已取消生成，本次结果未保存。');
          const record = records[index];
          setStatus(`正在生成 ${index + 1} / ${records.length} 份（数据第 ${record.rowNumber} 行）…`);
          const merged = mergeDocument(template, record);
          const mount = mountMergeReport(merged, renderBlocks, container); mounts.push(mount);
          let layout;
          try { layout = await mount.ready; } catch (cause) { throw new Error(`第 ${record.rowNumber} 行：${cause instanceof Error ? cause.message : String(cause)}`); }
          if (cancel.current) throw new Error('已取消生成，本次结果未保存。');
          // 功能：将实测的长表分页结果同步给 DOCX，成品与预览使用同一块序列；时间：2026-10-24；作者：lq。
          const pageBlocks: DocBlock[] = [];
          layout.slices.forEach((slice, pageIndex) => { if (pageIndex) pageBlocks.push({ type: 'pageBreak', id: `merge-break-${pageIndex}` }); pageBlocks.push(...slice.blocks); });
          const paged = { ...merged, blocks: pageBlocks };
          documents.push(paged);
          const fileName = mergeFileName(namePattern, record, index);
          if (effectiveOutput === 'docx-zip') addArchiveFile(`${fileName}.docx`, await serializeDocx(paged, { sourceBytes, name: fileName, materialized: true }));
          if (effectiveOutput === 'pdf-zip') addArchiveFile(`${fileName}.pdf`, await exportElementsToPdf({ pages: layout.pages, fileName, pageWidthPx: doc.page.widthPx, pageHeightPx: doc.page.heightPx, isCancelled: () => cancel.current }));
          if (effectiveOutput === 'print' || effectiveOutput === 'pdf') {
            allPages.push(...layout.pages);
            if (allPages.length > 500) throw new Error('合并结果超过 500 页，请缩小筛选范围后分批输出。');
          } else { mount.dispose(); mounts.pop(); }
          setProgress(Math.round((index + 1) / records.length * 85));
          await new Promise<void>((resolve) => setTimeout(resolve, 0));
        }
      }
      if (cancel.current) throw new Error('已取消生成，本次结果未保存。');
      setStatus('正在整理输出…');
      if (effectiveOutput === 'print') {
        container.className = 'mo-merge-print-root';
        allPages.at(-1)?.setAttribute('data-merge-last', 'true');
        const style = document.createElement('style');
        // 功能：只打印选中的报告，隐藏模板和交互界面，逐份独立分页；时间：2026-10-24；作者：lq。
        style.textContent = `@media print { html, body { height:auto !important;overflow:visible !important; } body > :not(.mo-merge-print-root) { display:none !important; } body * { visibility:hidden !important; } .mo-word-pages { display:none !important; } .mo-merge-print-root, .mo-merge-print-root * { visibility:visible !important; } .mo-merge-print-root { position:absolute !important;left:0 !important;top:0 !important;width:100% !important; } .mo-merge-print-root .mo-merge-measure { display:none !important; } .mo-merge-print-root [data-merge-page] { break-after:page;page-break-after:always;transform:none !important; } .mo-merge-print-root [data-merge-last] { break-after:auto;page-break-after:auto; } .mo-merge-print-root .flex { display:block !important; } @page { size:${doc.page.widthPx}px ${doc.page.heightPx}px; margin:0; } }`;
        container.appendChild(style);
        keepForPrint = true;
        printCleanup.current = cleanup;
        if (!await printCurrentView()) throw new Error('当前平台不能打开打印，请选择合并 PDF 后打印。');
        setStatus(`已准备 ${records.length} 份报告、${allPages.length} 页。打印面板中确认打印机和份数；关闭本窗口会清理临时打印页面。`);
      } else {
        let bytes: Uint8Array | undefined;
        let extension: string;
        let mime: string;
        // 功能：合并为一个 Word 文档时，把各条记录直接拼成一个文档并在新窗口打开，不弹保存框；时间：2026-10-30；作者：lq。
        let mergedDocument: DocModel | undefined;
        if (effectiveOutput === 'docx') {
          mergedDocument = mergeRecordsIntoDocument(documents);
          extension = 'docx'; mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
        } else if (effectiveOutput === 'pdf') {
          bytes = await exportElementsToPdf({ pages: allPages, fileName: baseName, pageWidthPx: doc.page.widthPx, pageHeightPx: doc.page.heightPx, isCancelled: () => cancel.current, onProgress: (value) => { setStatus(value.message); setProgress(85 + Math.round(value.current / Math.max(1, value.total) * 12)); } });
          extension = 'pdf'; mime = 'application/pdf';
        } else {
          bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }, () => { if (cancel.current) throw new Error('已取消生成，本次结果未保存。'); });
          extension = 'zip'; mime = 'application/zip';
        }
        if (cancel.current) throw new Error('已取消生成，本次结果未保存。');
        if (effectiveOutput === 'docx' && mergedDocument) {
          // 功能：合并 Word 直接在软件窗口显示合并后的新文档，不先弹保存框；保存由用户在文档窗口手动进行；时间：2026-10-30；作者：lq。
          if (onOpenDocument) {
            await onOpenDocument(mergedDocument, baseName);
          } else {
            const id = await useWorkspaceStore.getState().createDocument('docx');
            if (id) useWorkspaceStore.getState().updateData(id, mergedDocument, { dirty: true });
          }
          setStatus(`已生成 ${records.length} 份报告，合并文档已在新窗口打开。`);
          notify('success', '邮件合并完成', `${records.length} 份报告已合并为新文档，可继续编辑或手动保存。`);
        } else {
          await getPlatformAdapter().saveBytes(bytes!, { suggestedName: `${baseName}.${extension}`, forcePicker: true }, mime);
          setStatus(`已生成并保存 ${records.length} 份报告。`);
          notify('success', '邮件合并完成', `${records.length} 份报告已保存，模板和源数据保留。`);
        }
      }
      setProgress(100);
    } catch (cause) {
      const message = cause instanceof Error ? cause.message : String(cause);
      setError(message); setStatus('');
      if (keepForPrint) { keepForPrint = false; printCleanup.current = null; }
    } finally { if (!keepForPrint) cleanup(); setBusy(false); }
  };
  return {
    session, doc, renderBlocks,
    source, sheetId, setSheetId, headerRow, setHeaderRow, loading, error, openWorkbook, setOpenWorkbook,
    previewOn, setPreviewOn, togglePreview, previewZoom, setPreviewZoom,
    conditions, setConditions, search, setSearch, sort, setSort, descending, setDescending,
    selected, setSelected, tablePage, setTablePage, previewIndex, setPreviewIndex,
    allowEmpty, setAllowEmpty, namePattern, setNamePattern, output, setOutput,
    mergeRange, setMergeRange, rangeStart, setRangeStart, rangeEnd, setRangeEnd, rangedRecords, previewRecords,
    busy, progress, status, workbooks, templateFields, data, filtered, chosen, issues, current, previewDoc,
    multipleSections, hasNext, blocksPerPage, pageSize, totalPages, badNameFields, invalidName, canGenerate, result,
    chooseSource, importFile, importDataFile, cancelGeneration, generate,
  };
}

export type MailMergeController = ReturnType<typeof useMailMergeController>;
