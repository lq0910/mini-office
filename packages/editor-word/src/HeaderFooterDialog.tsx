import { useRef, useState } from 'react';
import type { DocModel, HeaderFooter, WordAlignment } from '@mini-office/document-model';
import { Button, Dialog, Field, Input, Select } from '@mini-office/ui';
import { mmToPx, pxToMm } from '@mini-office/ooxml';
import { headerFooterText, makeHeaderFooter, PAGE_NUMBER_TEMPLATES } from './header-footer';
import { renderParagraphHtml } from './html';

/** 功能：页眉页脚独立草稿、常用页码、光标插域及实时预览，应用后统一进入撤销历史；时间：2026-10-24；作者：lq。 */
export function HeaderFooterDialog({ doc, pageCount, initialRegion, onClose, onApply }: {
  doc: DocModel; pageCount: number; initialRegion: 'header' | 'footer'; onClose: () => void;
  onApply: (patch: Pick<DocModel, 'header' | 'footer' | 'page'>) => void;
}) {
  const [region, setRegion] = useState(initialRegion);
  const [draft, setDraft] = useState({ header: doc.header, footer: doc.footer });
  const [start, setStart] = useState(doc.page.pageNumberStart ?? 1);
  const [previewPage, setPreviewPage] = useState(1);
  const [layoutError, setLayoutError] = useState('');
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const current = draft[region];
  const text = headerFooterText(current);
  const paragraph = current?.blocks.find((block) => block.type === 'paragraph');
  const alignment = paragraph?.style.alignment ?? 'center';
  const fontSize = paragraph?.runs[0]?.style.fontSizePt ?? 10.5;
  const distance = current?.distancePx ?? doc.page.margins[region];
  // 功能：锚定文本框属于需保留的复杂内容，不能通过普通文字表单无提示覆盖；清空操作仍可显式移除；时间：2026-11-05；作者：lq。
  const complex = !!current?.textBoxes?.length || !!current?.blocks.some((block) => block.type !== 'paragraph' || block.runs.some((run) => run.image || (run.field && !['PAGE', 'NUMPAGES'].includes(run.field))));
  const validStart = Number.isInteger(start) && start >= 0 && start <= 32767;
  const update = (changes: { text?: string; alignment?: WordAlignment; fontSize?: number; distance?: number }) => {
    setLayoutError('');
    setDraft((previous) => ({ ...previous, [region]: makeHeaderFooter(changes.text ?? text, changes.alignment ?? alignment, changes.fontSize ?? fontSize, changes.distance ?? distance) }));
  };
  const insert = (token: string) => {
    const element = inputRef.current;
    const from = element?.selectionStart ?? text.length;
    const to = element?.selectionEnd ?? from;
    update({ text: text.slice(0, from) + token + text.slice(to) });
    requestAnimationFrame(() => { element?.focus(); element?.setSelectionRange(from + token.length, from + token.length); });
  };
  const apply = () => {
    if (!validStart) return;
    const margins = { ...doc.page.margins };
    // 功能：自动为新增内容留足空间，避免页眉页脚与正文相互遮挡；时间：2026-10-24；作者：lq。
    for (const kind of ['header', 'footer'] as const) {
      const value = draft[kind];
      if (!value?.edited) continue;
      margins[kind] = value.distancePx;
      const width = Math.max(120, doc.page.widthPx - margins.left - margins.right);
      // 功能：按正文宽度保守估算自动换行高度，防止长文字被页边距裁切；时间：2026-10-24；作者：lq。
      const height = headerFooterText(value) ? value.blocks.reduce((sum, block) => {
        if (block.type !== 'paragraph') return sum;
        const fontPx = Math.max(10.5, ...block.runs.map((run) => run.style.fontSizePt ?? 10.5)) * 4 / 3;
        const length = block.runs.reduce((count, run) => count + (run.field ? 5 : (run.text ?? '').replace(/\t/g, '    ').length), 0);
        return sum + Math.max(1, Math.ceil(length * fontPx / width)) * fontPx * 1.2;
      }, 0) : 0;
      const edge = kind === 'header' ? 'top' : 'bottom';
      if (height > 0) margins[edge] = Math.max(margins[edge], value.distancePx + height + 12);
    }
    if (margins.top + margins.bottom > doc.page.heightPx - 96) {
      setLayoutError('页眉页脚占用空间过大，请减少文字、字号或距纸边距离，为正文保留足够空间。');
      return;
    }
    onApply({ ...draft, page: { ...doc.page, margins, pageNumberStart: start } });
  };
  const preview = (value?: HeaderFooter) => value?.blocks.map((block, index) => block.type === 'paragraph'
    ? <div key={index} style={{ textAlign: block.style.alignment, fontSize: `${block.runs[0]?.style.fontSizePt ?? 10.5}pt`, lineHeight: 1.2 }} dangerouslySetInnerHTML={{ __html: renderParagraphHtml(block, { pageNumber: start + previewPage - 1, totalPages: pageCount }) }} />
    : <div key={index} className="text-center text-xs">已有表格内容（保留）</div>);
  return <Dialog open onClose={onClose} width="xl" title="页眉、页脚与页码" description="设置默认页眉页脚。页码随分页自动更新，保存后可在 Word 中继续编辑。"
    footer={<><Button onClick={onClose}>取消</Button><Button variant="primary" disabled={!validStart} onClick={apply}>应用到文档</Button></>}>
    <div className="grid gap-6 sm:grid-cols-[1fr_240px]">
      <div className="space-y-4">
        <div className="flex gap-2" role="group" aria-label="编辑区域">{(['header', 'footer'] as const).map((kind) => <Button key={kind} variant={region === kind ? 'primary' : 'secondary'} onClick={() => setRegion(kind)}>{kind === 'header' ? '页眉' : '页脚'}</Button>)}</div>
        {complex ? <p role="status" className="text-xs text-[var(--mo-warning)]">此区域含图片、表格或特殊域。原内容会保留；如需改为文字和页码，请先点击“清空此区域”。</p> : null}
        <Field label="常用形式" hint="选择样式会替换当前区域内容；也可在文字中插入页码。"><Select aria-label="常用页码形式" disabled={complex} value="choose" onChange={(event) => update({ text: event.target.value })}><option value="choose" disabled>选择文字或页码样式</option>{PAGE_NUMBER_TEMPLATES.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}</Select></Field>
        <Field label={region === 'header' ? '页眉内容' : '页脚内容'} hint="可换行。{PAGE} 表示当前页码，{NUMPAGES} 表示文档总页数。"><textarea ref={inputRef} aria-label={region === 'header' ? '页眉内容' : '页脚内容'} rows={3} maxLength={600} disabled={complex} value={text} onChange={(event) => update({ text: event.target.value })} placeholder="输入文字，或选择上方的页码样式" className="w-full resize-none rounded-md border border-[var(--mo-line)] bg-[var(--mo-surface)] p-2 text-sm outline-none focus:border-[var(--mo-accent)]" /></Field>
        <div className="flex flex-wrap gap-2"><Button size="sm" disabled={complex} onClick={() => insert('{PAGE}')}>插入当前页码</Button><Button size="sm" disabled={complex} onClick={() => insert('{NUMPAGES}')}>插入总页数</Button><Button size="sm" onClick={() => update({ text: '' })}>清空此区域</Button></div>
        <div className="grid grid-cols-3 gap-3">
          <Field label="对齐"><Select aria-label="页眉页脚对齐" disabled={complex} value={alignment} onChange={(event) => update({ alignment: event.target.value as WordAlignment })}><option value="left">左对齐</option><option value="center">居中</option><option value="right">右对齐</option><option value="justify">两端对齐</option></Select></Field>
          <Field label="字号"><Select aria-label="页眉页脚字号" disabled={complex} value={fontSize} onChange={(event) => update({ fontSize: Number(event.target.value) })}>{Array.from(new Set([9, 10.5, 12, 14, 16, fontSize])).sort((a, b) => a - b).map((size) => <option key={size} value={size}>{size} pt</option>)}</Select></Field>
          <Field label="距纸边"><Select aria-label="距纸边" disabled={complex} value={Math.round(pxToMm(distance) * 10) / 10} onChange={(event) => update({ distance: mmToPx(Number(event.target.value)) })}>{Array.from(new Set([5, 10, 12.7, 15, 20, 25, Math.round(pxToMm(distance) * 10) / 10])).sort((a, b) => a - b).map((mm) => <option key={mm} value={mm}>{mm} mm</option>)}</Select></Field>
        </div>
        <Field label="起始页码" hint={validStart ? '应用于整篇文档；总页数始终表示实际页数。空间不足时会自动增大正文页边距。' : '请输入 0–32767 的整数。'}><Input aria-label="起始页码" aria-invalid={!validStart} type="number" min={0} max={32767} value={Number.isNaN(start) ? '' : start} onChange={(event) => setStart(event.target.value === '' ? NaN : Number(event.target.value))} /></Field>
        {layoutError ? <p role="alert" className="text-xs text-[var(--mo-danger)]">{layoutError}</p> : null}
      </div>
      <div className="space-y-3">
        <div className="flex items-center justify-between text-xs"><span>内容预览</span><span>{previewPage} / {pageCount} 页</span></div>
        <div className="flex min-h-[300px] flex-col justify-between overflow-hidden rounded border border-gray-200 bg-white p-5 text-black shadow-sm">
          <div className="min-h-10 border-b border-dashed border-gray-200 pb-3 break-words">{preview(draft.header)}</div>
          <div className="space-y-3 py-8" aria-hidden>{[100, 92, 100, 85, 95].map((width, index) => <div key={index} className="h-1 rounded bg-gray-100" style={{ width: `${width}%` }} />)}</div>
          <div className="min-h-10 border-t border-dashed border-gray-200 pt-3 break-words">{preview(draft.footer)}</div>
        </div>
        <div className="flex justify-between"><Button size="sm" disabled={previewPage <= 1} onClick={() => setPreviewPage((page) => page - 1)}>上一页</Button><Button size="sm" disabled={previewPage >= pageCount} onClick={() => setPreviewPage((page) => page + 1)}>下一页</Button></div>
        <p className="text-[11px] leading-5 text-[var(--mo-text-muted)]">此处预览内容与对齐。实际位置请在应用后的文档页面查看；支持撤销。</p>
      </div>
    </div>
  </Dialog>;
}
