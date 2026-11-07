import { BrandIcon, FileIcon } from '@mini-office/ui';
import { useMemo, useState } from 'react';
import {
  ArrowRight,
  Clock3,
  ExternalLink,
  FilePlus2,
  FolderOpen,
  HardDrive,
  Search,
  ShieldCheck,
  Sparkles,
  Trash2,
  UploadCloud,
} from 'lucide-react';
import { useWorkspaceStore } from '@mini-office/document-model';
import { collectTemplates, readPluginFile, usePluginRegistry } from '@mini-office/plugin-host';
import { Button, Dialog, IconButton, Input } from '@mini-office/ui';
import { APP_VERSION, OFFICE_META, formatBytes, formatDateTime, type OfficeKind } from '@mini-office/shared';
import { isTauriRuntime, openExternalUrl } from '@mini-office/platform-adapter';
import { useLatestRelease } from './update-check';

const RELEASES_URL = 'https://github.com/lq0910/mini-office/releases';

// 功能：首页、模板和最近文件统一使用类型字标图标；时间：2026-11-07；作者：lq。
const KIND_TINTS: Record<OfficeKind, { bg: string; color: string }> = {
  docx: { bg: '#EEF3FF', color: '#2B579A' },
  xlsx: { bg: '#E9F7EF', color: '#217346' },
  pptx: { bg: '#FCEFE9', color: '#C43E1C' },
  pdf: { bg: '#FDECEC', color: '#E34F4F' },
  md: { bg: '#F0EEFF', color: '#6B5BD2' },
  html: { bg: '#FFF1E3', color: '#E67E22' },
};

export function HomePage() {
  const latestRelease = useLatestRelease();
  const recent = useWorkspaceStore((state) => state.recent);
  const createDocument = useWorkspaceStore((state) => state.createDocument);
  const openFileDialog = useWorkspaceStore((state) => state.openFileDialog);
  const openSnapshot = useWorkspaceStore((state) => state.openSnapshot);
  const deleteSnapshot = useWorkspaceStore((state) => state.deleteSnapshot);
  // 功能：复用工作区快照清理并同步最近文件列表；时间：2026-10-24；作者：lq。
  const clearSnapshots = useWorkspaceStore((state) => state.clearSnapshots);
  const updateData = useWorkspaceStore((state) => state.updateData);
  const notify = useWorkspaceStore((state) => state.notify);
  const openDialog = useWorkspaceStore((state) => state.openDialog);
  const openBytes = useWorkspaceStore((state) => state.openBytes);
  // 订阅 installed，安装/禁用插件后模板区立即刷新。
  const installedPlugins = usePluginRegistry((state) => state.installed);
  const pluginTemplates = useMemo(() => collectTemplates(installedPlugins), [installedPlugins]);
  const [query, setQuery] = useState('');
  const [busy, setBusy] = useState(false);
  // 功能：管理批量清空确认和处理中状态，防止误删与重复提交；时间：2026-10-24；作者：lq。
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [clearing, setClearing] = useState(false);

  const filtered = useMemo(() => {
    const keyword = query.trim().toLowerCase();
    if (!keyword) return recent;
    return recent.filter(
      (item) =>
        item.displayName.toLowerCase().includes(keyword) ||
        OFFICE_META[item.kind].label.toLowerCase().includes(keyword),
    );
  }, [query, recent]);

  const dirtyCount = recent.filter((item) => item.dirty).length;

  const createBlank = async (kind: OfficeKind) => {
    if (busy) return;
    setBusy(true);
    try {
      await createDocument(kind);
    } finally {
      setBusy(false);
    }
  };

  const createContract = async () => {
    if (busy) return;
    setBusy(true);
    try {
      const id = await createDocument('docx');
      if (!id) return;
      // 功能：选择模板后再加载模板与计算逻辑，首页启动不加载报销表公式引擎；时间：2026-10-24；作者：lq。
      const { contractTemplate } = await import('./templates');
      updateData(id, contractTemplate(), { dirty: true });
      notify('success', '已创建合同模板', '可按需修改条款并另存为正式合同。');
    } finally {
      setBusy(false);
    }
  };

  const createExpense = async () => {
    if (busy) return;
    setBusy(true);
    try {
      const id = await createDocument('xlsx');
      if (!id) return;
      // 功能：费用模板按需加载，保留公式预计算行为；时间：2026-10-24；作者：lq。
      const { expenseTemplate } = await import('./templates');
      updateData(id, expenseTemplate(), { dirty: true });
      notify('success', '已创建费用报销模板', '公式已自动计算，可直接填写金额。');
    } finally {
      setBusy(false);
    }
  };

  const openFile = async () => {
    if (busy) return;
    setBusy(true);
    try {
      await openFileDialog();
    } finally {
      setBusy(false);
    }
  };

  /** 从插件包里取出模板字节，按普通文件打开——模板永远不会被覆盖，用户保存时走另存为。 */
  const createFromPluginTemplate = async (pluginId: string, file: string, label: string, kind: OfficeKind) => {
    if (busy) return;
    setBusy(true);
    try {
      const bytes = await readPluginFile(pluginId, file);
      if (!bytes) {
        notify('error', '模板文件缺失', '插件可能未完整安装，请重新安装该插件。');
        return;
      }
      const name = `${label}.${kind}`;
      await openBytes({ name, bytes, restored: { dirty: true } });
      notify('success', `已从模板新建「${label}」`, '模板创建为未保存副本，不会覆盖插件文件。');
    } catch (error) {
      notify('error', '打开模板失败', error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  };

  const openRecent = async (id: string) => {
    if (busy) return;
    setBusy(true);
    try {
      await openSnapshot(id);
    } finally {
      setBusy(false);
    }
  };

  // 功能：清空所有最近快照并重置搜索，失败时保留确认框供重试；时间：2026-10-24；作者：lq。
  const clearRecent = async () => {
    if (busy || clearing) return;
    setBusy(true);
    setClearing(true);
    try {
      await clearSnapshots();
      if (useWorkspaceStore.getState().recent.length > 0) {
        throw new Error('部分恢复快照未能清除，请重试。');
      }
      setQuery('');
      setShowClearConfirm(false);
      notify('success', '已清空最近文件与恢复快照', '磁盘上的原始 Office 文件不受影响。');
    } catch (error) {
      notify('error', '清空失败', error instanceof Error ? error.message : String(error));
    } finally {
      setClearing(false);
      setBusy(false);
    }
  };

  return (
    <div className="mo-home mo-scroll h-full overflow-y-auto bg-[var(--mo-bg)]">
      <div className="mx-auto w-full max-w-[1120px] px-6 pb-16 pt-8 md:px-10">
        <section className="mo-glass relative overflow-hidden rounded-2xl border border-[var(--mo-line)] bg-[var(--mo-surface)] px-7 py-8 shadow-panel md:px-10 md:py-10">
          <div
            aria-hidden
            className="pointer-events-none absolute -right-20 -top-28 h-72 w-72 rounded-full opacity-60 blur-3xl"
            style={{
              background:
                'radial-gradient(circle, color-mix(in srgb, var(--mo-accent) 20%, transparent), transparent 70%)',
            }}
          />
          <div
            aria-hidden
            className="pointer-events-none absolute -bottom-32 right-40 h-64 w-64 rounded-full opacity-50 blur-3xl"
            style={{ background: 'radial-gradient(circle, rgba(18,183,106,0.16), transparent 70%)' }}
          />
          <div className="relative flex flex-col gap-7 md:flex-row md:items-end md:justify-between">
            <div className="max-w-[620px]">
              <div className="mb-4 flex items-center gap-2.5">
                <LogoMark />
                <span className="text-[11px] font-semibold uppercase tracking-[0.18em] text-[var(--mo-text-muted)]">
                  Mini Office
                </span>
              </div>
              <h1 className="text-[30px] font-semibold leading-tight tracking-tight text-[var(--mo-text-primary)] md:text-[36px]">
                只保留真正需要的
                <span className="text-[var(--mo-accent)]"> Office</span> 功能
              </h1>
              <p className="mt-3 max-w-[540px] text-[13px] leading-6 text-[var(--mo-text-muted)]">
                离线优先、轻量快速的跨平台办公套件。Word、Excel、PPT 在同一处编辑，
                PDF、Markdown、HTML 在同一处预览，重视真实文件兼容与文档安全。
              </p>
              <div className="mt-5 flex flex-wrap items-center gap-2">
                <Button variant="primary" icon={<FolderOpen size={15} />} onClick={() => void openFile()} loading={busy}>
                  打开文件
                </Button>
                <Button variant="secondary" icon={<FilePlus2 size={15} />} onClick={() => void createBlank('docx')}>
                  新建文档
                </Button>
                <button
                  type="button"
                  onClick={() => void openFile()}
                  className="ml-1 inline-flex items-center gap-1.5 rounded-md px-2 py-1.5 text-[12px] text-[var(--mo-text-muted)] hover:bg-[var(--mo-surface-2)]"
                >
                  <UploadCloud size={14} /> 或将 DOCX / XLSX / PPTX / PDF / MD / HTML 拖入窗口
                </button>
              </div>
            </div>
            <div className="flex shrink-0 flex-wrap gap-2 md:max-w-[280px] md:justify-end">
              {[
                { icon: ShieldCheck, label: '离线优先' },
                { icon: HardDrive, label: '本地保存' },
                { icon: Sparkles, label: '中文友好' },
              ].map(({ icon: Icon, label }) => (
                <span
                  key={label}
                  className="inline-flex items-center gap-1.5 rounded-full border border-[var(--mo-line)] bg-[var(--mo-surface-2)]/60 px-2.5 py-1 text-[11px] text-[var(--mo-text-muted)]"
                >
                  <Icon size={12} /> {label}
                </span>
              ))}
            </div>
          </div>
        </section>

        <section className="mt-6 grid gap-4 lg:grid-cols-[1.4fr_1fr]">
          <div className="grid gap-3 sm:grid-cols-3">
            <CreateCard
              kind="docx"
              title="Word 文档"
              description="合同、公文与常用文档编辑"
              actionLabel="新建文档"
              onAction={() => void createBlank('docx')}
            />
            <CreateCard
              kind="xlsx"
              title="Excel 表格"
              description="数据录入、公式计算与打印"
              actionLabel="新建表格"
              onAction={() => void createBlank('xlsx')}
            />
            <CreateCard
              kind="pptx"
              title="PowerPoint"
              // 功能：启用 PPT 新建与日常编辑入口；时间：2026-10-24；作者：lq。
              description="内容编辑、页面排版与全屏演示"
              actionLabel="新建演示"
              onAction={() => void createBlank('pptx')}
            />
          </div>

          <div className="rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface)] p-4 shadow-panel">
            <div className="mb-3 flex items-center justify-between">
              <h2 className="text-[13px] font-semibold">从模板开始</h2>
              <span className="text-[10px] text-[var(--mo-text-faint)]">
                {pluginTemplates.length > 0 ? `内置 + ${pluginTemplates.length} 个插件模板` : '内置模板'}
              </span>
            </div>
            <div className="space-y-2">
              <TemplateRow
                icon={<FileIcon kind="docx" size={25} />}
                tint="#2B579A"
                title="房屋租赁合同"
                description="标准条款、首行缩进与签章区"
                onClick={() => void createContract()}
              />
              <TemplateRow
                icon={<FileIcon kind="xlsx" size={25} />}
                tint="#217346"
                title="费用报销单"
                description="含 SUM 公式、日期与金额格式"
                onClick={() => void createExpense()}
              />
              {pluginTemplates.map(({ key, pluginId, pluginName, template }) => {
                const tint = KIND_TINTS[template.kind];
                return (
                  <TemplateRow
                    key={key}
                    icon={<FileIcon kind={template.kind} size={22} />}
                    tint={tint.color}
                    title={template.label}
                    description={template.description ?? `来自插件 ${pluginName}`}
                    onClick={() =>
                      void createFromPluginTemplate(pluginId, template.file, template.label, template.kind)
                    }
                  />
                );
              })}
            </div>
            <div className="mt-3 rounded-lg bg-[var(--mo-surface-2)] px-3 py-2 text-[10px] leading-4 text-[var(--mo-text-muted)]">
              模板创建为未保存副本，保存时会提示选择位置，不会覆盖任何模板文件。
            </div>
          </div>
        </section>

        {dirtyCount > 0 ? (
          <section className="mt-5 flex items-center gap-3 rounded-xl border border-[color-mix(in_srgb,var(--mo-warning)_35%,transparent)] bg-[var(--mo-warning-soft)] px-4 py-3">
            <Clock3 size={16} className="shrink-0 text-[var(--mo-warning)]" />
            <div className="min-w-0 flex-1">
              <div className="text-[12px] font-medium text-[var(--mo-text-primary)]">
                发现 {dirtyCount} 个未保存的自动恢复快照
              </div>
              <div className="text-[10px] text-[var(--mo-text-muted)]">
                快照保存在本机；恢复后可继续编辑并另存为文件。
              </div>
            </div>
            <Button
              size="sm"
              variant="soft"
              onClick={() => {
                const first = recent.find((item) => item.dirty);
                if (first) void openRecent(first.id);
              }}
            >
              恢复最近一个
            </Button>
          </section>
        ) : null}

        <section className="mt-8">
          <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
            <div className="flex items-baseline gap-2">
              <h2 className="text-[15px] font-semibold">最近文件</h2>
              <span className="text-[10px] text-[var(--mo-text-faint)]">
                本机保留 {recent.length} 个恢复快照
              </span>
            </div>
            {/* 功能：将快照清空入口置于搜索框旁，窄窗口允许换行；时间：2026-10-24；作者：lq。 */}
            <div className="flex flex-wrap items-center gap-2">
              <div className="relative w-[240px] max-w-full">
                <Search aria-hidden size={13} className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-[var(--mo-text-faint)]" />
                {/* 功能：提高左侧留白优先级，避免全局 mo-field 内边距覆盖后图文重叠；时间：2026-10-24；作者：lq。 */}
                <Input
                  aria-label="搜索最近文件"
                  value={query}
                  onChange={(event) => setQuery(event.target.value)}
                  placeholder="搜索最近文件…"
                  className="h-8 !pl-8 text-[12px]"
                />
              </div>
              <Button
                size="sm"
                variant="ghost"
                icon={<Trash2 size={13} />}
                disabled={recent.length === 0 || busy}
                loading={clearing}
                onClick={() => setShowClearConfirm(true)}
                title="删除本机所有最近文件记录与恢复快照"
                className="text-[var(--mo-text-muted)] hover:text-[var(--mo-danger)]"
              >
                一键清空
              </Button>
            </div>
          </div>
          <div className="overflow-hidden rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface)] shadow-panel">
            {filtered.length === 0 ? (
              <div className="flex flex-col items-center justify-center gap-3 px-6 py-14 text-center">
                <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-[var(--mo-surface-2)] text-[var(--mo-text-muted)]">
                  <FolderOpen size={22} />
                </div>
                <div>
                  <div className="text-[13px] font-medium">{recent.length === 0 ? '还没有最近文件' : '没有匹配的文件'}</div>
                  <div className="mt-1 text-[11px] text-[var(--mo-text-muted)]">
                    {recent.length === 0
                      ? '打开一个 Office 文件，或从上方模板开始。'
                      : '尝试更换搜索关键词。'}
                  </div>
                </div>
                {recent.length === 0 ? (
                  <Button variant="primary" size="sm" icon={<FolderOpen size={14} />} onClick={() => void openFile()}>
                    打开文件
                  </Button>
                ) : null}
              </div>
            ) : (
              <ul className="divide-y divide-[var(--mo-line)]">
                {filtered.map((item) => {
                  const tint = KIND_TINTS[item.kind];
                  return (
                    <li key={item.id}>
                      <div className="group flex items-center gap-3 px-4 py-3 transition-colors hover:bg-[var(--mo-surface-2)]">
                        <button
                          type="button"
                          onClick={() => void openRecent(item.id)}
                          className="flex min-w-0 flex-1 items-center gap-3 text-left"
                        >
                          <span
                            className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg"
                            style={{ background: tint.bg, color: tint.color }}
                          >
                            <FileIcon kind={item.kind} name={item.displayName} size={30} />
                          </span>
                          <span className="min-w-0 flex-1">
                            <span className="flex items-center gap-2">
                              <span className="truncate text-[12.5px] font-medium text-[var(--mo-text-primary)]">
                                {item.displayName}
                              </span>
                              {item.dirty ? (
                                <span className="shrink-0 rounded bg-[var(--mo-warning-soft)] px-1 py-0.5 text-[9px] text-[var(--mo-warning)]">
                                  未保存
                                </span>
                              ) : null}
                            </span>
                            <span className="mt-0.5 flex items-center gap-2 text-[10px] text-[var(--mo-text-muted)]">
                              <span>{OFFICE_META[item.kind].label}</span>
                              <span className="text-[var(--mo-text-faint)]">·</span>
                              <span>{formatBytes(item.size)}</span>
                              <span className="text-[var(--mo-text-faint)]">·</span>
                              <span>{formatDateTime(item.updatedAt)}</span>
                              {item.path ? (
                                <>
                                  <span className="text-[var(--mo-text-faint)]">·</span>
                                  <span className="max-w-[260px] truncate">{item.path}</span>
                                </>
                              ) : null}
                            </span>
                          </span>
                          <ArrowRight
                            size={14}
                            className="shrink-0 text-[var(--mo-text-faint)] opacity-0 transition-opacity group-hover:opacity-100"
                          />
                        </button>
                        <IconButton
                          label="从最近文件移除"
                          // 功能：批量清空期间避免并发移除快照；时间：2026-10-24；作者：lq。
                          disabled={clearing}
                          onClick={() => void deleteSnapshot(item.id)}
                          className="opacity-0 transition-opacity group-hover:opacity-100"
                        >
                          <Trash2 size={13} />
                        </IconButton>
                      </div>
                    </li>
                  );
                })}
              </ul>
            )}
          </div>
        </section>

        {/* 功能：清空前确认全部快照范围及未保存内容风险，兼容桌面 WebView；时间：2026-10-24；作者：lq。 */}
        <Dialog
          open={showClearConfirm}
          onClose={() => { if (!clearing) setShowClearConfirm(false); }}
          title="清空最近文件与恢复快照？"
          icon={<Trash2 size={16} />}
          width="sm"
          closeOnBackdrop={!clearing}
          footer={
            <>
              <Button disabled={clearing} onClick={() => setShowClearConfirm(false)}>取消</Button>
              <Button variant="danger" loading={clearing} onClick={() => void clearRecent()}>确认清空</Button>
            </>
          }
        >
          <p className="text-[12px] leading-6 text-[var(--mo-text-muted)]">
            将删除本机全部 {recent.length} 个恢复快照和最近文件记录，包括搜索结果以外的快照。磁盘上的原始 Office 文件不会被删除，已打开的文档也不会关闭。
          </p>
          {dirtyCount > 0 ? (
            <p className="mt-3 rounded-lg bg-[var(--mo-warning-soft)] px-3 py-2 text-[12px] leading-5 text-[var(--mo-warning)]">
              其中 {dirtyCount} 个快照包含未保存内容，删除后将无法通过这些快照恢复。请确认需要的内容已另行保存。
            </p>
          ) : null}
        </Dialog>

        <section className="mt-8 flex flex-wrap items-center justify-between gap-3 border-t border-[var(--mo-line)] pt-5 text-[10px] text-[var(--mo-text-faint)]">
          <div className="flex flex-wrap items-center gap-4">
            {/* 功能：展示统一版本号；时间：2026-10-24；作者：lq。 */}
            <span>Mini Office {APP_VERSION}</span>
            <span>离线优先 · 不上传文件</span>
            <span>DOCX / XLSX / PPTX</span>
            <a
              href={latestRelease?.url ?? RELEASES_URL}
              target="_blank"
              rel="noreferrer"
              // 功能：提供新版本发布页入口；时间：2026-10-24；作者：lq。
              onClick={(event) => {
                if (!isTauriRuntime()) return;
                event.preventDefault();
                void openExternalUrl(RELEASES_URL);
              }}
              className="inline-flex items-center gap-1 rounded px-1.5 py-1 text-[var(--mo-accent)] hover:bg-[var(--mo-surface-2)] hover:underline"
            >
              {latestRelease ? `发现新版本 ${latestRelease.tagName}` : '新版本更新'}
              <ExternalLink size={11} />
            </a>
          </div>
          <button
            type="button"
            onClick={() => openDialog({ type: 'shortcuts' })}
            className="rounded px-1.5 py-1 hover:bg-[var(--mo-surface-2)] hover:text-[var(--mo-text-muted)]"
          >
            键盘快捷键
          </button>
        </section>
      </div>
    </div>
  );
}

/** 功能：首页复用正式品牌图标；时间：2026-10-18；作者：lq。 */
function LogoMark({ size = 34 }: { size?: number }) {
  return <BrandIcon size={size} />;
}

function CreateCard({
  kind,
  title,
  description,
  actionLabel,
  onAction,
}: {
  kind: OfficeKind;
  title: string;
  description: string;
  actionLabel: string;
  onAction: () => void;
}) {
  const tint = KIND_TINTS[kind];
  return (
    <button
      type="button"
      onClick={onAction}
      className="group flex flex-col items-start rounded-xl border border-[var(--mo-line)] bg-[var(--mo-surface)] p-4 text-left shadow-panel transition-all duration-fast hover:-translate-y-0.5 hover:border-[var(--mo-line-strong)] hover:shadow-pop"
    >
      <span
        className="mb-4 flex h-10 w-10 items-center justify-center rounded-lg"
        style={{ background: tint.bg, color: tint.color }}
      >
        <FileIcon kind={kind} size={36} />
      </span>
      <span className="text-[13px] font-semibold text-[var(--mo-text-primary)]">{title}</span>
      <span className="mt-1 text-[10.5px] leading-4 text-[var(--mo-text-muted)]">{description}</span>
      <span className="mt-4 inline-flex items-center gap-1 text-[11px] font-medium text-[var(--mo-accent)]">
        {actionLabel}
        <ArrowRight size={12} className="transition-transform duration-fast group-hover:translate-x-0.5" />
      </span>
    </button>
  );
}

function TemplateRow({
  icon,
  tint,
  title,
  description,
  onClick,
}: {
  icon: React.ReactNode;
  tint: string;
  title: string;
  description: string;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="flex w-full items-center gap-3 rounded-lg border border-transparent px-2 py-2 text-left transition-colors hover:border-[var(--mo-line)] hover:bg-[var(--mo-surface-2)]"
    >
      <span
        className="flex h-8 w-8 items-center justify-center rounded-md"
        style={{ background: `${tint}14`, color: tint }}
      >
        {icon}
      </span>
      <span className="min-w-0 flex-1">
        <span className="block truncate text-[11.5px] font-medium">{title}</span>
        <span className="block truncate text-[10px] text-[var(--mo-text-muted)]">{description}</span>
      </span>
      <ArrowRight size={13} className="text-[var(--mo-text-faint)]" />
    </button>
  );
}
