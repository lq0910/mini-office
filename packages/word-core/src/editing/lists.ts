/** 功能：同源事务切换列表，新增编号定义仅在快照引用时写入 DOCX；时间：2026-11-10；作者：lq。 */
import type { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import type { Node } from '@tiptap/pm/model';
import type { NumberingDef, SaveOptions, CustomNumberingLevel } from '../upstream/packages/docx-engine/src/index';
import { selectedParagraphs } from './formatting';

/** 功能：工具栏只显示选区一致的列表类别及层级；时间：2026-11-10；作者：lq。 */
export function readWordList(editor: Editor): { kind?: 'bullet' | 'ordered'; level?: number; mixed: boolean } {
  const paragraphs = selectedParagraphs(editor);
  const listOf = (node: Node) => node.type.name === 'docListItem' ? { kind: node.attrs.kind, level: Number(node.attrs.ilvl) || 0 } : node.attrs.listMarker?.kind ? { kind: node.attrs.listMarker.kind, level: Number(node.attrs.listMarker.ilvl) || 0 } : undefined;
  const first = paragraphs[0] && listOf(paragraphs[0].node);
  const mixed = paragraphs.some(({ node }) => JSON.stringify(listOf(node)) !== JSON.stringify(first));
  const sameKind = paragraphs.every(({ node }) => listOf(node)?.kind === first?.kind);
  return { kind: sameKind ? first?.kind : undefined, level: mixed ? undefined : first?.level, mixed };

}

/** 功能：工具栏增减缩进调整每个选中列表项的真实层级，保留编号身份且不残留旧直接缩进；时间：2026-11-10；作者：lq。 */
export function changeWordListLevel(editor: Editor, delta: number): boolean {
  if (!editor.isEditable || editor.view.composing) return false;
  const paragraphs = selectedParagraphs(editor);
  if (!paragraphs.length || paragraphs.some(({ node }) => node.type.name !== 'docListItem' && !node.attrs.listMarker?.kind)) return false;
  const tr = closeHistory(editor.state.tr);
  for (const { node, pos } of paragraphs) {
    const current = node.attrs.listMarker?.ilvl ?? node.attrs.ilvl;
    const ilvl = Math.max(0, Math.min(8, (Number(current) || 0) + delta));
    if (ilvl !== current) tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...(node.attrs.listMarker ? { listMarker: { ...node.attrs.listMarker, ilvl, indentLeft: 720 * (ilvl + 1), hanging: 360 } } : { ilvl }), indentLeft: null, indentFirstLine: null, indentFirstLineChars: null });
  }
  if (!tr.docChanged) return false;
  editor.view.dispatch(tr);
  return true;
}

/** 功能：编号资源归当前会话所有；撤销后未引用的定义不污染保存，重做可恢复原引用；时间：2026-11-10；作者：lq。 */
export class WordListDefinitions {
  private readonly pending = new Map<string, { numId: string; kind: 'bullet' | 'ordered'; levels: CustomNumberingLevel[] }>();
  constructor(private readonly original: Map<string, NumberingDef>) {}
  /** 功能：列表样式和正文列表共用分配器，定义先分配、随实际引用快照保存；时间：2026-11-11；作者：lq。 */
  forStyle(editor: Editor, kind: 'bullet' | 'ordered', numId?: string): string {
    return numId && (this.original.has(numId) || this.pending.has(numId)) ? numId : this.allocate(editor, kind);
  }

  /** 功能：分配与导入文档及空模板均不冲突的编号，九层缩进与保存定义保持一致；时间：2026-11-10；作者：lq。 */
  private allocate(editor: Editor, kind: 'bullet' | 'ordered'): string {
    const ids = [...this.original.keys(), ...this.pending.keys(), ...editor.storage.listNumbering.defs.keys()];
    const numId = String(Math.max(2, ...ids.map(Number).filter(Number.isFinite)) + 1);
    const levels = Array.from({ length: 9 }, (_, level) => ({ numFmt: kind === 'bullet' ? 'bullet' : 'decimal', lvlText: kind === 'bullet' ? '•' : `%${level + 1}.`, indentLeft: 720 * (level + 1), hanging: 360, start: 1 }));
    this.pending.set(numId, { numId, kind, levels });
    editor.storage.listNumbering.defs = new Map(editor.storage.listNumbering.defs).set(numId, {
      numId, abstractNumId: `pending-${numId}`, startOverrides: {}, levels: Object.fromEntries(levels.map((level, i) => [i, level])),
    });
    return numId;
  }

  /** 功能：同类再次点击取消列表；转换保留来源锚点、段落及字符属性，邻接同类列表沿用编号；时间：2026-11-10；作者：lq。 */
  toggle(editor: Editor, kind: 'bullet' | 'ordered', owner = editor): boolean {
    // 功能：精简区域 schema 使用原生 listMarker，资源仍由唯一主会话管理；时间：2026-11-11；作者：lq。
    if (!editor.schema.nodes.docListItem && editor.schema.nodes.docParagraph?.spec.attrs?.listMarker) return this.toggleRegion(editor, owner, kind);
    // 功能：精简的文本框 schema 没有正文列表节点，迁移前拒绝执行，避免误写正文或丢失标题级别；时间：2026-11-10；作者：lq。
    if (!editor.isEditable || editor.view.composing || !editor.schema.nodes.docListItem || !editor.storage.listNumbering) return false;
    const paragraphs = selectedParagraphs(editor).filter(({ node }) => ['docParagraph', 'docListItem'].includes(node.type.name));
    if (!paragraphs.length) return false;
    const remove = paragraphs.every(({ node }) => node.type.name === 'docListItem' && node.attrs.kind === kind);
    const tr = closeHistory(editor.state.tr);
    let numId: string | undefined;
    if (!remove) {
      // 功能：先继承选区内或前后相邻列表；不采用导入文件中碰巧同号但不同格式的模板编号；时间：2026-11-10；作者：lq。
      const existing = paragraphs.find(({ node }) => node.type.name === 'docListItem' && node.attrs.kind === kind)?.node;
      const before = editor.state.doc.resolve(paragraphs[0].pos).nodeBefore;
      const last = paragraphs[paragraphs.length - 1];
      const after = editor.state.doc.resolve(last.pos + last.node.nodeSize).nodeAfter;
      const candidate = [existing, before, after].find(node => node?.type.name === 'docListItem' && node.attrs.kind === kind && node.attrs.numId && editor.storage.listNumbering.defs.has(String(node.attrs.numId)));
      numId = candidate ? String(candidate.attrs.numId) : this.allocate(editor, kind);
    }
    for (const { node, pos } of paragraphs) {
      const attrs = { ...node.attrs };
      if (remove) {
        // 功能：取消列表清除编号字段，原段落直接缩进和样式继续保留；时间：2026-11-10；作者：lq。
        delete attrs.kind; delete attrs.numId; delete attrs.ilvl;
      } else Object.assign(attrs, { kind, numId, ilvl: node.type.name === 'docListItem' ? Number(node.attrs.ilvl) || 0 : 0 });
      tr.setNodeMarkup(pos, remove ? editor.schema.nodes.docParagraph : editor.schema.nodes.docListItem, attrs);
    }
    editor.view.dispatch(tr);
    return true;
  }

  /** 功能：文本框和页眉页脚不切换节点类型，只修改原生编号身份及可撤销属性；时间：2026-11-11；作者：lq。 */
  private toggleRegion(editor: Editor, owner: Editor, kind: 'bullet' | 'ordered'): boolean {
    if (!editor.isEditable || editor.view.composing) return false;
    const paragraphs = selectedParagraphs(editor); if (!paragraphs.length) return false;
    const remove = paragraphs.every(({ node }) => node.attrs.listMarker?.kind === kind);
    const existing = paragraphs.find(({ node }) => node.attrs.listMarker?.kind === kind)?.node.attrs.listMarker;
    const numId = remove ? undefined : this.forStyle(owner, kind, existing?.numId);
    const tr = closeHistory(editor.state.tr);
    for (const { node, pos } of paragraphs) {
      const ilvl = Number(node.attrs.listMarker?.ilvl) || 0;
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, listMarker: remove ? null : { kind, numId, ilvl, text: kind === 'bullet' ? '•' : '1.', indentLeft: 720 * (ilvl + 1), hanging: 360 } });
    }
    editor.view.dispatch(tr); return true;
  }

  /** 功能：从不可变保存快照收集实际使用的定义，避免保存、撤销和重做产生悬空编号；时间：2026-11-10；作者：lq。 */
  saveOptions(doc: Node): NonNullable<SaveOptions['numbering']> {
    const used = new Set<string>();
    doc.descendants(node => { if (node.type.name === 'docListItem' && node.attrs.numId) used.add(String(node.attrs.numId)); });
    // 功能：区域内容保存在主文档不可变属性中，同样收集它实际引用的编号；时间：2026-11-11；作者：lq。
    const collect = (value: unknown): void => {
      if (!value || typeof value !== 'object') return;
      if (Array.isArray(value)) { value.forEach(collect); return; }
      const record = value as Record<string, unknown>;
      const marker = record.listMarker as { numId?: string } | undefined; if (marker?.numId) used.add(String(marker.numId));
      for (const [key, child] of Object.entries(record)) if (key !== 'runs' && key !== 'miniOfficeCellSource') collect(child);
    };
    collect(doc.attrs.miniOfficePageLayout?.edits);
    doc.descendants(node => { collect(node.attrs.textboxes); });
    const styles = doc.attrs.miniOfficeStyles?.definitions ?? [];
    for (const style of styles) if (style.upsert.pPr?.numPr?.numId) used.add(String(style.upsert.pPr.numPr.numId));
    return { newDefs: [...this.pending.values()].filter(def => used.has(def.numId)) };
  }
}
