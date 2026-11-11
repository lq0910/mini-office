/** 功能：沿用文档属性面板的页面控件，按分节提交原生设置草稿；时间：2026-11-10；作者：lq。 */
import { useState } from 'react';
import { X, Pilcrow, FileWarning } from 'lucide-react';
import type { DocumentSession } from '@mini-office/document-model';
import { formatBytes, MOD_KEY } from '@mini-office/shared';
import { Button, Field, IconButton, Input, Select, SegmentedControl } from '@mini-office/ui';
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import type { WordPagePatch } from '../../word-core/src/editing/page-layout';

/** 功能：以 OOXML twips 保存标准纸张尺寸，避免 px 往返改变原始边距；时间：2026-11-10；作者：lq。 */
const SIZES = { A4: [11906, 16838], A3: [16838, 23811], A5: [8391, 11906], B5: [10318, 14570], Letter: [12240, 15840], Legal: [12240, 20160] } as const;
const MM = 1440 / 25.4;
/** 功能：页面范围明确为选中的节，草稿取消不修改文档；时间：2026-11-10；作者：lq。 */
export function WordCorePagePanel({ core, readonly, session, onClose, onApplied, onOpenParagraph, onCompatibility }: { core: WordCoreSession; readonly: boolean; session: DocumentSession; onClose: () => void; onApplied: () => void; onOpenParagraph: () => void; onCompatibility: () => void }) {
  const [section, setSection] = useState(0);
  const [patch, setPatch] = useState<WordPagePatch>({});
  const [error, setError] = useState('');
  const sections = core.sections();
  const stats = core.statistics();
  const current = sections[section] ?? sections[0];
  const settings = { ...current.settings, ...patch };
  const size = Object.entries(SIZES).find(([, [w, h]]) => Math.abs(Math.min(settings.pageWidth, settings.pageHeight) - w) <= 2 && Math.abs(Math.max(settings.pageWidth, settings.pageHeight) - h) <= 2)?.[0] ?? '自定义';
  /** 功能：每次修改只合并用户字段，不能把其他页面选项写回默认值；时间：2026-11-10；作者：lq。 */
  const update = (value: WordPagePatch) => { setPatch(previous => ({ ...previous, ...value })); setError(''); };
  const apply = () => {
    try {
      const catalog = core.pageLayout();
      core.commitPageLayout({ ...catalog, sections: { ...catalog.sections, [current.lastBlockIndex]: { ...catalog.sections[current.lastBlockIndex], ...patch } } });
      setPatch({}); setError(''); onApplied();
    } catch (caught) { setError(caught instanceof Error ? caught.message : String(caught)); }
  };
  return <aside aria-label="文档属性" className="mo-glass mo-no-print flex w-[268px] shrink-0 flex-col overflow-y-auto border-l border-[var(--mo-line)] bg-[var(--mo-surface)]">
    <div className="flex h-10 items-center justify-between border-b border-[var(--mo-line)] px-3"><span className="text-[12px] font-semibold">文档属性</span><IconButton label="关闭属性面板" onClick={onClose}><X size={14} /></IconButton></div>
    <div className="space-y-4 p-3">
      <Field label="应用范围"><Select aria-label="页面设置分节" value={section} onChange={event => { setSection(Number(event.target.value)); setPatch({}); setError(''); }}>{sections.map((item, index) => <option key={item.lastBlockIndex} value={index}>第 {index + 1} 节{sections.length === 1 ? '（整篇文档）' : ''}</option>)}</Select></Field>
      <Field label="纸张尺寸"><Select aria-label="纸张尺寸" disabled={readonly} value={size} onChange={event => {
        const value = SIZES[event.target.value as keyof typeof SIZES]; if (!value) return;
        const [w, h] = value; update({ pageWidth: settings.orientation === 'landscape' ? h : w, pageHeight: settings.orientation === 'landscape' ? w : h });
      }}>{Object.keys(SIZES).map(name => <option key={name}>{name}</option>)}<option value="自定义" disabled>自定义</option></Select></Field>
      <Field label="方向"><fieldset disabled={readonly}><SegmentedControl value={settings.orientation} className="w-full" onChange={orientation => { if (orientation !== settings.orientation) update({ orientation, pageWidth: settings.pageHeight, pageHeight: settings.pageWidth }); }} options={[{ value: 'portrait', label: '纵向' }, { value: 'landscape', label: '横向' }]} /></fieldset></Field>
      <div className="grid grid-cols-2 gap-2">{(['Top', 'Right', 'Bottom', 'Left'] as const).map((side, index) => {
        const key = `margin${side}` as const;
        const label = `页边距 ${['上', '右', '下', '左'][index]} (mm)`;
        return <Field key={key} label={label}><Input aria-label={label} disabled={readonly} type="number" min={0} step={0.1} value={Number.isNaN(settings[key]) ? '' : Math.round(settings[key] / MM * 10) / 10} onChange={event => update({ [key]: event.target.value === '' ? NaN : Math.round(Number(event.target.value) * MM) })} /></Field>;
      })}</div>
      <p className="text-[11px] leading-5 text-[var(--mo-text-muted)]">设置应用于所选分节，其他节保留原来的纸张和边距。</p>
      {error ? <p role="alert" className="text-xs text-[var(--mo-danger)]">{error}</p> : null}
      <Button variant="primary" disabled={readonly || !Object.keys(patch).length} onClick={apply}>应用页面设置</Button>
      {/* 功能：复用原属性面板的段落入口、统计和兼容报告；时间：2026-11-11；作者：lq。 */}
      <section className="space-y-2 border-t border-[var(--mo-line)] pt-3">
        <h3 className="text-[11px] font-semibold text-[var(--mo-text-muted)]">段落格式</h3>
        <p className="text-[10px] leading-4 text-[var(--mo-text-muted)]">缩进、行距、段前段后间距和分页选项应用于当前选区。</p>
        <Button size="sm" disabled={readonly} onClick={onOpenParagraph}><Pilcrow size={13} />打开段落设置…<span className="text-[var(--mo-text-faint)]">{MOD_KEY}⌥M</span></Button>
      </section>
      <section className="space-y-1.5 border-t border-[var(--mo-line)] pt-3" aria-label="文档统计">
        <h3 className="text-[11px] font-semibold text-[var(--mo-text-muted)]">统计</h3>
        {Object.entries({ 页数: session.view.pageCount ?? 1, 字数: stats.words, 字符数: stats.characters, 文件大小: formatBytes(session.sourceSize), 格式级别: `L${session.capability.level}` }).map(([label, value]) => <div key={label} className="flex justify-between text-[11px]"><span className="text-[var(--mo-text-muted)]">{label}</span><span className="tabular-nums">{value}</span></div>)}
      </section>
      <section className="space-y-2 border-t border-[var(--mo-line)] pt-3">
        <Button size="sm" onClick={onCompatibility}><FileWarning size={14} />查看兼容性报告<span>{session.capability.warnings.length}</span></Button>
        {session.capability.warnings.slice(0, 3).map(warning => <p key={warning.code} className="text-[10px] text-[var(--mo-text-muted)]">· {warning.message}</p>)}
      </section>
    </div>
  </aside>;
}
