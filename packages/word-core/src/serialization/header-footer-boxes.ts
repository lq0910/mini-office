/** 功能：在原页眉页脚的 DrawingML/VML 文本框内局部写回富文本与页码，原定位和备用分支同步保留；时间：2026-11-11；作者：lq。 */
import { inlineRunsXml, mergePPrFormat } from '../upstream/packages/docx-engine/src/generate';
import { PAGE_MARK, TOTAL_PAGES_MARK, SECTION_PAGES_MARK, type HfParagraph, type Run, type HeaderFooter } from '../upstream/packages/docx-engine/src/index';
import type { WordBoxFooter } from '../editing/header-footer-boxes';
import { xmlSpans, wordXmlNode, applyXmlEdits, type XmlSpan } from './xml-spans';

/** 功能：域按 OOXML 写为可更新的复杂域，各组成运行继承原字号字体；时间：2026-11-11；作者：lq。 */
function boxRuns(runs: Run[], total: number): string {
  const fields: Record<string, [string, number]> = { [PAGE_MARK]: ['PAGE', 1], [TOTAL_PAGES_MARK]: ['NUMPAGES', total], [SECTION_PAGES_MARK]: ['SECTIONPAGES', 1] };
  return runs.map(run => run.text.split(new RegExp(`([${PAGE_MARK}${TOTAL_PAGES_MARK}${SECTION_PAGES_MARK}])`)).map(text => {
    const field = fields[text];
    if (!field) return text ? inlineRunsXml([{ ...run, text }]) : '';
    const rPr = /<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>|<w:rPr\b[^>]*\/>/.exec(inlineRunsXml([{ ...run, text: 'x' }]))?.[0] ?? '';
    return `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rPr}<w:instrText xml:space="preserve"> ${field[0]} </w:instrText></w:r><w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rPr}<w:t>${Math.max(1, Math.round(field[1]))}</w:t></w:r><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`;
  }).join('')).join('');
}

/** 功能：按命名空间和原形状 ID 精确找到文本框，含空框和同页多个框时不依赖段落序号猜测；时间：2026-11-11；作者：lq。 */
function shapeId(node: XmlSpan): string | undefined {
  for (let p = node.parent; p; p = p.parent) {
    if (p.local === 'anchor' || p.local === 'inline') { const id = p.children.find(c => c.local === 'docPr')?.attrs.id; return id ? `drawing:${id}` : undefined; }
    if (p.prefix === 'v' && p.attrs.id) return `vml:${p.attrs.id}`;
  }
}

/** 功能：只改已编辑文本框的内部段落，保留原根、绘图、关系、未知属性和未改形状字节；时间：2026-11-11；作者：lq。 */
export function patchWordHfTextboxes(xml: string, hf: HeaderFooter, total = 1): string {
  const ids = (hf as WordBoxFooter).miniOfficeBoxEdits;
  if (!ids?.length) return xml;
  const nodes = xmlSpans(xml), edits: Array<{ start: number; end: number; text: string }> = [];
  for (const id of ids) {
    const paras = hf.paras?.filter(p => p.box?.id === id) ?? [];
    const sourceId = paras[0]?.box?.miniOfficeShapeId;
    const source = nodes.find(n => wordXmlNode(n, 'txbxContent') && shapeId(n) === sourceId);
    if (!source || !paras.length) throw new Error('无法定位原页眉页脚文本框，已取消有损保存');
    let alternate: XmlSpan | undefined;
    for (let p = source.parent; p; p = p.parent) if (p.local === 'AlternateContent') { alternate = p; break; }
    // 功能：一个兼容分支可能有多个文本框，按同一分支内次序配对，不能把目标文字写入其他框；时间：2026-11-11；作者：lq。
    const targets = [source];
    if (alternate) {
      let branch = source;
      while (branch.parent && branch.parent !== alternate) branch = branch.parent;
      const boxesIn = (part: XmlSpan) => nodes.filter(n => wordXmlNode(n, 'txbxContent') && n.start > part.start && n.end < part.end);
      const siblings = boxesIn(branch), index = siblings.indexOf(source);
      for (const other of alternate.children.filter(n => n !== branch && (n.local === 'Choice' || n.local === 'Fallback'))) {
        const boxes = boxesIn(other);
        if (!boxes.length) continue;
        if (boxes.length !== siblings.length || index < 0) throw new Error('页眉页脚文本框的兼容分支无法对应，已取消有损保存');
        targets.push(boxes[index]);
      }
    }
    for (const target of targets) {
      const originals = target.children.filter(p => wordXmlNode(p, 'p'));
      const body = paras.map((para: HfParagraph, index) => {
        const original = originals[Math.min(index, originals.length - 1)];
        const pPr = original?.children.find(p => wordXmlNode(p, 'pPr'));
        const sourcePr = pPr ? xml.slice(pPr.start, pPr.end) : '<w:pPr/>';
        return `<w:p>${mergePPrFormat(sourcePr, para)}${boxRuns(para.runs, total)}</w:p>`;
      }).join('');
      edits.push({ start: target.openEnd, end: target.closeStart, text: body });
    }
  }
  return applyXmlEdits(xml, edits);
}
