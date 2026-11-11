/** 功能：复用原页眉页脚对话框交互，以原生部件草稿编辑文字与动态页码；时间：2026-11-10；作者：lq。 */
import { useRef, useState } from 'react';
import { Button, Dialog, Field, Input, Select } from '@mini-office/ui';
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import { wordHeaderFooterState, type WordPageLayout, type WordPagePatch } from '../../word-core/src/editing/page-layout';
import { resolveHf, withHfEdit, withHfLink, hfLinked } from '../../word-core/src/upstream/apps/docs/src/renderer/hf-sections';
import { PAGE_MARK, TOTAL_PAGES_MARK, SECTION_PAGES_MARK, type HeaderFooter, type HfParagraph } from '../../word-core/src/index';
import { PAGE_NUMBER_TEMPLATES } from './header-footer';
// 功能：预览沿用原文档页码格式，不把罗马数字强制显示为十进制；时间：2026-11-10；作者：lq。
import { formatPageNumber } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-hf';

const MM = 1440 / 25.4;
/** 功能：动态域转为编辑标记，普通井号仍是普通字符；时间：2026-11-10；作者：lq。 */
export function coreHeaderFooterText(value: HeaderFooter | null): string {
  const text = value?.paras ? value.paras.map(para => para.runs.map(run => run.text).join('')).join('\n') : value?.text ?? '';
  const result = text.replaceAll(PAGE_MARK, '{PAGE}').replaceAll(TOTAL_PAGES_MARK, '{NUMPAGES}').replaceAll(SECTION_PAGES_MARK, '{SECTIONPAGES}');
  return value?.pageNumber && !text.includes(PAGE_MARK) ? result + '{PAGE}' : result;
}
/** 功能：文字和页码构造原生段落，预览与保存都使用同一动态域标记；时间：2026-11-10；作者：lq。 */
export function coreMakeHeaderFooter(text: string, align: HfParagraph['align'], sizeHalfPoints: number): HeaderFooter {
  const native = text.replaceAll('{PAGE}', PAGE_MARK).replaceAll('{NUMPAGES}', TOTAL_PAGES_MARK).replaceAll('{SECTIONPAGES}', SECTION_PAGES_MARK);
  return { text: native, pageNumber: native.includes(PAGE_MARK), paras: native.split('\n').map(line => ({ align, runs: [{ text: line, sizeHalfPoints }] })) };
}
/** 功能：草稿不写入会话；应用一次进入正文历史，保留未改区域及分节变体；时间：2026-11-10；作者：lq。 */
export function WordCoreHeaderFooterDialog({ core, pageCount, initialRegion = 'footer', onClose }: { core: WordCoreSession; pageCount: number; initialRegion?: 'header' | 'footer'; onClose: () => void }) {
  const [draft, setDraft] = useState<WordPageLayout>(() => structuredClone(core.pageLayout()));
  const [section, setSection] = useState(0);
  const [kind, setKind] = useState(initialRegion);
  const [variant, setVariant] = useState<'default' | 'first' | 'even'>('default');
  const [previewPage, setPreviewPage] = useState(1);
  const [error, setError] = useState('');
  const input = useRef<HTMLTextAreaElement>(null);
  // 功能：只构造目录摘要，不转换正文或复制旧 DocModel；时间：2026-11-10；作者：lq。
  const state = wordHeaderFooterState(core.editor.state.doc, core.parsed);
  state.edits = draft.edits; state.links = draft.links;
  state.sections = state.sections.map(item => ({ ...item, titlePg: draft.sections[item.lastBlockIndex]?.titlePg ?? item.titlePg, settings: { ...item.settings, ...draft.sections[item.lastBlockIndex] }, pageNumberStart: draft.sections[item.lastBlockIndex]?.pageNumberStart ?? item.pageNumberStart }));
  const evenOdd = draft.evenAndOddHeaders ?? core.evenAndOddHeaders();
  const selected = state.sections[section];
  const resolved = resolveHf(state, section, kind, variant);
  const value = resolved.value;
  const text = coreHeaderFooterText(value);
  const para = value?.paras?.[0];
  const align = para?.align ?? 'center';
  const size = para?.runs[0]?.sizeHalfPoints ?? 21;
  const distanceKey = kind === 'header' ? 'headerDist' : 'footerDist';
  const distance = selected.settings[distanceKey] ?? 720;
  const start = selected.pageNumberStart ?? 1;
  const complex = !!resolved.images?.length || !!value?.paras?.some(p => p.cells || p.box || p.boxAnchored || p.lineOnly || p.runs.some(run => run.link || run.instrField || run.refField || run.fldBeginXml));
  /** 功能：仅用户修改的节字段进入草稿，关联部件仍按来源继承；时间：2026-11-10；作者：lq。 */
  const patchSection = (patch: WordPagePatch) => { setError(''); setDraft(previous => ({ ...previous, sections: { ...previous.sections, [selected.lastBlockIndex]: { ...previous.sections[selected.lastBlockIndex], ...patch } } })); };
  const update = (changes: { text?: string; align?: HfParagraph['align']; size?: number }) => {
    const next = withHfEdit(state, section, kind, variant, coreMakeHeaderFooter(changes.text ?? text, changes.align ?? align, changes.size ?? size));
    if (next) { setError(''); setDraft(previous => ({ ...previous, ...next })); }
  };
  const insert = (token: string) => {
    const from = input.current?.selectionStart ?? text.length; const to = input.current?.selectionEnd ?? from;
    update({ text: text.slice(0, from) + token + text.slice(to) });
    requestAnimationFrame(() => { input.current?.focus(); input.current?.setSelectionRange(from + token.length, from + token.length); });
  };
  const apply = () => { try { core.commitPageLayout(draft); onClose(); } catch (caught) { setError(caught instanceof Error ? caught.message : String(caught)); } };
  const renderPreview = (region: 'header' | 'footer') => {
    const part = resolveHf(state, section, region, variant).value;
    return <div style={{ textAlign: part?.paras?.[0]?.align === 'distribute' ? 'justify' : part?.paras?.[0]?.align ?? 'center', fontSize: `${(part?.paras?.[0]?.runs[0]?.sizeHalfPoints ?? 21) / 2}pt`, whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' }}>{coreHeaderFooterText(part).replaceAll('{PAGE}', formatPageNumber(start + previewPage - 1, selected.pageNumberFmt)).replaceAll('{NUMPAGES}', String(pageCount))}</div>;
  };
  return <Dialog open onClose={onClose} width="xl" title="页眉、页脚与页码" description="编辑所选分节的页眉页脚；沿用同一页眉页脚的分节同步更新。"
    footer={<><Button onClick={onClose}>取消</Button><Button variant="primary" onClick={apply}>应用到文档</Button></>}>
    <div className="grid gap-6 sm:grid-cols-[1fr_240px]">
      <div className="space-y-4">
        <div className="grid grid-cols-2 gap-3"><Field label="分节"><Select aria-label="页眉页脚分节" value={section} onChange={event => { setSection(Number(event.target.value)); setVariant('default'); }}>{state.sections.map((item, index) => <option key={item.lastBlockIndex} value={index}>第 {index + 1} 节</option>)}</Select></Field>
        <Field label="页面类型"><Select aria-label="页眉页脚页面类型" value={variant} onChange={event => setVariant(event.target.value as typeof variant)}><option value="default">默认页面</option>{selected.titlePg ? <option value="first">首页</option> : null}{evenOdd ? <option value="even">偶数页</option> : null}</Select></Field></div>
        {/* 功能：变体开关仅控制显示规则，关闭时保留隐藏变体和其来源部件；时间：2026-11-11；作者：lq。 */}
        <div className="flex flex-wrap gap-4 text-sm"><label className="flex items-center gap-2"><input type="checkbox" aria-label="首页不同" checked={selected.titlePg} onChange={event => { patchSection({ titlePg: event.target.checked }); if (!event.target.checked && variant === 'first') setVariant('default'); }} />首页不同（当前节）</label><label className="flex items-center gap-2"><input type="checkbox" aria-label="奇偶页不同" checked={evenOdd} onChange={event => { setDraft(previous => ({ ...previous, evenAndOddHeaders: event.target.checked })); if (!event.target.checked && variant === 'even') setVariant('default'); }} />奇偶页不同（整篇）</label></div>
        <div className="flex gap-2" role="group" aria-label="编辑区域">{(['header', 'footer'] as const).map(region => <Button key={region} variant={kind === region ? 'primary' : 'secondary'} onClick={() => setKind(region)}>{region === 'header' ? '页眉' : '页脚'}</Button>)}</div>
        {resolved.owner !== section ? <p role="status" className="text-xs text-[var(--mo-text-muted)]">此区域沿用第 {resolved.owner + 1} 节；编辑会同步到沿用此页眉页脚的分节。</p> : null}
        {section > 0 ? <label className="flex items-center gap-2 text-sm"><input type="checkbox" aria-label="链接到前一节" checked={!!hfLinked(state, section, kind, variant)} onChange={event => { const next = withHfLink(state, section, kind, variant, event.target.checked); if (next) setDraft(previous => ({ ...previous, ...next })); }} />链接到前一节</label> : null}
        {complex ? <p role="status" className="text-xs text-[var(--mo-warning)]">此区域含图片、表格或特殊内容；请双击页面上的页眉或页脚编辑，设置面板保留原内容。</p> : null}
        <Field label="常用形式"><Select aria-label="常用页码形式" disabled={complex} value="choose" onChange={event => update({ text: event.target.value })}><option value="choose" disabled>选择文字或页码样式</option>{PAGE_NUMBER_TEMPLATES.map(item => <option key={item.value} value={item.value}>{item.label}</option>)}</Select></Field>
        <Field label={kind === 'header' ? '页眉内容' : '页脚内容'} hint="可换行。{PAGE} 为当前页码，{NUMPAGES} 为文档总页数。"><textarea ref={input} aria-label={kind === 'header' ? '页眉内容' : '页脚内容'} rows={3} maxLength={600} disabled={complex} value={text} onChange={event => update({ text: event.target.value })} className="w-full resize-none rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface)] p-2 text-sm outline-none focus:border-[var(--mo-accent)]" /></Field>
        <div className="flex flex-wrap gap-2"><Button size="sm" disabled={complex} onClick={() => insert('{PAGE}')}>插入当前页码</Button><Button size="sm" disabled={complex} onClick={() => insert('{NUMPAGES}')}>插入总页数</Button><Button size="sm" disabled={complex} onClick={() => insert('{SECTIONPAGES}')}>插入本节页数</Button><Button size="sm" disabled={complex} onClick={() => update({ text: '' })}>清空此区域</Button></div>
        <div className="grid grid-cols-3 gap-3">
          <Field label="对齐"><Select aria-label="页眉页脚对齐" disabled={complex} value={align} onChange={event => update({ align: event.target.value as HfParagraph['align'] })}><option value="left">左对齐</option><option value="center">居中</option><option value="right">右对齐</option><option value="justify">两端对齐</option></Select></Field>
          <Field label="字号"><Select aria-label="页眉页脚字号" disabled={complex} value={size} onChange={event => update({ size: Number(event.target.value) })}>{Array.from(new Set([18, 21, 24, 28, 32, size])).sort((a,b) => a-b).map(value => <option key={value} value={value}>{value / 2} pt</option>)}</Select></Field>
          <Field label="距纸边"><Select aria-label="距纸边" value={distance} onChange={event => patchSection({ [distanceKey]: Number(event.target.value) })}>{Array.from(new Set([5, 10, 12.7, 15, 20, 25].map(mm => Math.round(mm * MM)).concat(distance))).sort((a,b) => a-b).map(value => <option key={value} value={value}>{Math.round(value / MM * 10) / 10} mm</option>)}</Select></Field>
        </div>
        <Field label="起始页码" hint="只修改所选分节的起始页码；总页数仍是实际页数。"><Input aria-label="起始页码" type="number" min={0} max={32767} value={Number.isNaN(start) ? '' : start} onChange={event => patchSection({ pageNumberStart: event.target.value === '' ? NaN : Number(event.target.value) })} /></Field>
        {error ? <p role="alert" className="text-xs text-[var(--mo-danger)]">{error}</p> : null}
      </div>
      <div className="space-y-3"><div className="flex justify-between text-xs"><span>内容预览</span><span>{previewPage} / {pageCount} 页</span></div>
        <div className="flex min-h-[300px] flex-col justify-between rounded border border-gray-200 bg-white p-5 text-black shadow-sm"><div className="min-h-10 border-b border-dashed border-gray-200 pb-3">{renderPreview('header')}</div><div className="space-y-3 py-8" aria-hidden>{[100,92,100,85,95].map((width,index) => <div key={index} className="h-1 rounded bg-gray-100" style={{ width: `${width}%` }} />)}</div><div className="min-h-10 border-t border-dashed border-gray-200 pt-3">{renderPreview('footer')}</div></div>
        <div className="flex justify-between"><Button size="sm" disabled={previewPage <= 1} onClick={() => setPreviewPage(page => page - 1)}>上一页</Button><Button size="sm" disabled={previewPage >= pageCount} onClick={() => setPreviewPage(page => page + 1)}>下一页</Button></div>
        <p className="text-[11px] leading-5 text-[var(--mo-text-muted)]">此处预览文字与对齐；实际位置与复杂内容以应用后的页面为准。支持撤销。</p>
      </div>
    </div>
  </Dialog>;
}
