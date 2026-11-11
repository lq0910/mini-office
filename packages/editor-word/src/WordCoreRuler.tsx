/** 功能：复用现有标尺，分节几何和段落草稿直接来自新核心；时间：2026-11-11；作者：lq。 */
import { useLayoutEffect, useRef, useState } from 'react';
import type { ParagraphStyle, TabStop } from '@mini-office/document-model';
import type { WordCoreSession } from '../../word-core/src/session/word-session';
import { selectedParagraphs } from '../../word-core/src/editing/formatting';
import { applyWordRuler, wordRulerTabs } from '../../word-core/src/editing/ruler';
import { sectionVertical, paperWidthPx } from '../../word-core/src/upstream/apps/docs/src/renderer/pagination-sections';
import { WordRuler } from './WordRuler';
import { coreToolbarFormatting } from './core-formatting';

/** 功能：正文、同格及文本框共享缩进/制表位，多栏和竖排保护原有布局；时间：2026-11-11；作者：lq。 */
export function WordCoreRuler({ core, zoom, readonly }: { core: WordCoreSession; zoom: number; readonly: boolean }) {
  const anchor = useRef<HTMLDivElement>(null);
  const [left, setLeft] = useState(0);
  // 功能：采用纸张真实原点，兼容内层滚动槽、缩放和窄窗口横向滚动；时间：2026-11-11；作者：lq。
  useLayoutEffect(() => {
    const align = () => {
      const paper = core.editor.view.dom.closest<HTMLElement>('.page-wrap');
      if (!paper || !anchor.current || paper.getBoundingClientRect().width <= 0) return;
      const next = paper.getBoundingClientRect().left-anchor.current.getBoundingClientRect().left;
      setLeft(previous => Math.abs(previous-next) < 0.1 ? previous : next);
    };
    align(); const observer = new ResizeObserver(align);
    observer.observe(core.editor.view.dom); if (anchor.current) observer.observe(anchor.current);
    return () => observer.disconnect();
  }, [core, zoom]);
  const [preview, setPreview] = useState<Partial<ParagraphStyle>>({});
  const gesture = useRef<{ editor: typeof core.editor; doc: typeof core.editor.state.doc; bookmark: ReturnType<typeof core.editor.state.selection.getBookmark>; patch: Partial<ParagraphStyle> }>();
  // 功能：标尺跟随当前已登记的编辑区域，拖动书签与文档版本也属于该区域；时间：2026-11-11；作者：lq。
  const editor = core.commandEditor();
  const regionSection = editor.view.dom.closest<HTMLElement>('[data-word-region-section]')?.dataset.wordRegionSection;
  const paragraphs = selectedParagraphs(editor);
  const sections = core.sections();
  // 功能：按仍在文档中的分节锚段定位插入段落，避免新段落无来源索引时误归第一节；时间：2026-11-11；作者：lq。
  const sectionAt = (pos: number) => {
    let index = 0;
    core.editor.state.doc.forEach((node, offset) => {
      if (offset >= pos) return;
      const boundary = sections.findIndex(section => section.lastBlockIndex === node.attrs.docxIndex);
      if (boundary >= 0 && boundary < sections.length - 1) index = boundary + 1;
    });
    return index;
  };
  const section = regionSection !== undefined ? Math.max(0, Math.min(sections.length - 1, Number(regionSection))) : sectionAt(editor.state.selection.$from.depth ? editor.state.selection.$from.before(1) : editor.state.selection.from);
  const settings = sections[section].settings;
  // 功能：格内和文本框以真实容器净宽为局部标尺原点，不借用正文边距；时间：2026-11-11；作者：lq。
  const containerAt = (pos: number) => {
    const dom = editor.view.nodeDOM(pos);
    return (dom instanceof HTMLElement ? dom : dom?.parentElement)?.closest<HTMLElement>('td,th,.page-hf-cell,.doc-textbox');
  };
  const container = paragraphs[0] && containerAt(paragraphs[0].pos);
  const paper = core.editor.view.dom.closest<HTMLElement>('.page-wrap');
  const factor = paper ? paper.getBoundingClientRect().width / paper.offsetWidth || zoom : zoom;
  const bounds = container?.getBoundingClientRect(), paperBounds = paper?.getBoundingClientRect();
  const css = container ? getComputedStyle(container) : undefined;
  const contentLeft = bounds && paperBounds ? (bounds.left-paperBounds.left)/factor + parseFloat(css?.paddingLeft ?? '0') : (paperWidthPx(sections)-settings.pageWidth/15)/2 + settings.marginLeft/15;
  const contentWidth = bounds ? bounds.width/factor-parseFloat(css?.paddingLeft ?? '0')-parseFloat(css?.paddingRight ?? '0') : (settings.pageWidth-settings.marginLeft-settings.marginRight)/15;
  const enabled = !readonly && !editor.view.composing && paragraphs.length > 0
    && paragraphs.every(({node, pos}) => ['docParagraph', 'docHeading', 'docListItem'].includes(node.type.name) && (container ? containerAt(pos) === container : editor.state.doc.resolve(pos).depth === 0 && (regionSection !== undefined || sectionAt(pos) === section)))
    && !sectionVertical(settings) && (!!container || (settings.columns ?? 1) === 1 && !core.parsed.mirrorMargins);
  const source = core.formattingSource();
  const tabs = paragraphs[0] ? wordRulerTabs(paragraphs[0].node.attrs, source).filter(stop => stop.val !== 'bar').map(stop => ({ posPx: stop.pos / 15, alignment: stop.val as TabStop['alignment'], leader: stop.leader })) : [];
  const mixedTabs = paragraphs.some(({node}) => JSON.stringify(wordRulerTabs(node.attrs, source).filter(stop => stop.val !== 'bar')) !== JSON.stringify(paragraphs[0] && wordRulerTabs(paragraphs[0].node.attrs, source).filter(stop => stop.val !== 'bar')));
  const current = coreToolbarFormatting(core).paragraph;
  /** 功能：恢复原选区并仅提交用户改变的字段，过期拖动不覆盖期间其他编辑；时间：2026-11-11；作者：lq。 */
  const commit = (patch: Partial<ParagraphStyle>) => {
    if (!enabled || gesture.current && (gesture.current.editor !== editor || gesture.current.doc !== editor.state.doc)) return;
    if (gesture.current) editor.view.dispatch(editor.state.tr.setSelection(gesture.current.bookmark.resolve(editor.state.doc)));
    applyWordRuler(editor, source, {
      ...('indentLeftPx' in patch ? { indentLeft: patch.indentLeftPx! * 15 } : {}),
      ...('indentRightPx' in patch ? { indentRight: patch.indentRightPx! * 15 } : {}),
      ...('firstLineIndentPx' in patch ? { indentFirstLine: patch.firstLineIndentPx! * 15 } : {}),
      ...(patch.tabStops ? { tabs: patch.tabStops.map(stop => ({ pos: Math.round(stop.posPx * 15), val: stop.alignment, leader: stop.leader })) } : {}),
    });
  };
  const update = (patch: Partial<ParagraphStyle>) => {
    if (gesture.current) { gesture.current.patch = { ...gesture.current.patch, ...patch }; setPreview(gesture.current.patch); }
    else commit(patch);
  };
  return <div ref={anchor} className="sticky top-0 z-20 mb-3" title={!enabled ? '当前区域仅显示标尺；请在单栏横排段落或页眉页脚中调整缩进。' : mixedTabs ? '选区制表位不同，请缩小选区后调整标尺。' : undefined}>
    <div style={{ width: paperWidthPx(sections) * zoom, marginLeft: left }}>
      <WordRuler pageWidthPx={paperWidthPx(sections)} contentLeftPx={contentLeft} contentWidthPx={contentWidth}
        zoom={zoom} readonly={!enabled || mixedTabs} style={{ ...current, tabStops: tabs, ...preview }} defaultTabStopPx={(core.parsed.defaultTabStopTwips ?? 720)/15}
        onDragStart={() => { gesture.current = { editor, doc: editor.state.doc, bookmark: editor.state.selection.getBookmark(), patch: {} }; }}
        onDragEnd={cancelled => { if (!cancelled && gesture.current && Object.keys(gesture.current.patch).length) commit(gesture.current.patch); gesture.current = undefined; setPreview({}); }}
        onChangeIndent={patch => update(patch)} onChangeTabStops={tabStops => update({ tabStops })} />
    </div>
  </div>;
}
