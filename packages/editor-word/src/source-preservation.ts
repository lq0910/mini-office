/** 功能：按 GenOffice 的原始 XML 区间策略保存未修改正文，纯文字修改保留原 run 属性和未知节点；时间：2026-11-03；作者：lq。 */
import type { DocBlock, DocModel, WordSourceSnapshot } from '@mini-office/document-model';
import { scanBody } from '../../platform-adapter/src/vendor/genoffice/xml-scan.js';
import { patchParagraphTexts } from '../../platform-adapter/src/vendor/genoffice/xml-text-patch.js';
// 功能：普通段落与嵌套表格共用属性组局部补丁；时间：2026-11-07；作者：lq。
import { patchSourceBlock } from './source-block-patch';

/** 功能：排除分页统计和来源数据，稳定比较实际可保存内容；时间：2026-11-03；作者：lq。 */
export function documentSignature(doc: DocModel): string {
  return JSON.stringify({ ...doc, source: undefined, stats: undefined, sectionPageCount: undefined });
}

/** 功能：把顶层原始元素与解析出的块绑定，未识别节点也保留在原文区间；时间：2026-11-03；作者：lq。 */
export function captureDocumentSource(doc: DocModel, xml: string, groups: string[][]): WordSourceSnapshot | undefined {
  try {
    const scan = scanBody(xml);
    if (scan.elements.length !== groups.length) return undefined;
    const byId = new Map(doc.blocks.map((block) => [block.id, block]));
    return {
      xml, bodyStart: scan.bodyContentStart, bodyEnd: scan.bodyContentEnd,
      signature: documentSignature(doc), page: JSON.stringify(doc.page), styles: styleSignature(doc),
      groups: scan.elements.map((element, index) => ({
        start: element.start, end: element.end, ids: groups[index],
        baseline: JSON.stringify(groups[index].map((id) => byId.get(id))),
      })),
    };
  } catch {
    // 功能：非标准命名空间或无法定位的源 XML 回退现有生成流程，不猜测区间写坏文件；时间：2026-11-03；作者：lq。
    return undefined;
  }
}

/** 功能：纯文字变化时仅替换 w:t 内容；字段、图片、换行等复杂内容走块生成；时间：2026-11-03；作者：lq。 */
function patchText(raw: string, baseline: string, blocks: DocBlock[]): string | null {
  const previous = JSON.parse(baseline) as DocBlock[];
  if (previous.length !== 1 || blocks.length !== 1) return null;
  const before = previous[0];
  const after = blocks[0];
  if (before.type !== 'paragraph' || after.type !== 'paragraph') return null;
  if (![...before.runs, ...after.runs].every((run) => Object.keys(run).every((key) => key === 'text' || key === 'style'))) return null;
  const withoutText = (paragraph: typeof after) => ({ ...paragraph, runs: paragraph.runs.map((run) => ({ ...run, text: '' })) });
  if (JSON.stringify(withoutText(before)) !== JSON.stringify(withoutText(after))) return null;
  // 功能：只有解析文字与原始 w:t 内容一致时才使用文字补丁，避免隐藏文本被错误消费；时间：2026-11-03；作者：lq。
  const text = before.runs.map((run) => run.text ?? '').join('');
  if (patchParagraphTexts(raw, text) !== raw) return null;
  return patchParagraphTexts(raw, after.runs.map((run) => run.text ?? '').join(''));
}

/** 功能：未改的顶层块逐字保留，插入/删除仍保留相邻未知元素；页面改变时回退完整节重建；时间：2026-11-03；作者：lq。 */
export function patchSourceDocument(doc: DocModel, generated: Map<string, string>): string | undefined {
  const source = doc.source;
  if (!source || JSON.stringify(doc.page) !== source.page) return undefined;
  // 功能：目录仍由目录写入器刷新，但保留来源快照供表格保存与无操作透传使用；时间：2026-11-07；作者：lq。
  if (doc.blocks.some(block => block.type === 'paragraph' && block.toc)) return undefined;
  // 功能：显式修改节或页眉页脚时交回节写入器，避免保留旧引用而丢掉用户修改；时间：2026-11-03；作者：lq。
  if (doc.header?.edited || doc.footer?.edited || doc.blocks.some((block) => block.type === 'paragraph' && (block.sectionEnd?.header?.edited || block.sectionEnd?.footer?.edited))) return undefined;
  const current = new Map(doc.blocks.map((block) => [block.id, block]));
  for (const group of source.groups) {
    const before = JSON.parse(group.baseline) as DocBlock[];
    if (before.some((block) => {
      const after = current.get(block.id);
      return block.type === 'paragraph' && after?.type === 'paragraph' && JSON.stringify(block.sectionEnd) !== JSON.stringify(after.sectionEnd);
    })) return undefined;
  }
  const present = new Set(doc.blocks.map((block) => block.id));
  const byId = new Map<string, WordSourceSnapshot['groups'][number]>();
  const prefixes = new Map<WordSourceSnapshot['groups'][number], string>();
  let cursor = source.bodyStart;
  let pending = '';
  for (const group of source.groups) {
    pending += source.xml.slice(cursor, group.start);
    cursor = group.end;
    if (group.ids.length === 0) {
      pending += source.xml.slice(group.start, group.end);
      continue;
    }
    if (group.ids.some((id) => present.has(id))) {
      prefixes.set(group, pending);
      pending = '';
      for (const id of group.ids) byId.set(id, group);
    }
  }
  pending += source.xml.slice(cursor, source.bodyEnd);
  const emitted = new Set<WordSourceSnapshot['groups'][number]>();
  let body = '';
  for (let i = 0; i < doc.blocks.length; i += 1) {
    const block = doc.blocks[i];
    const group = byId.get(block.id);
    if (group && !emitted.has(group)) {
      body += prefixes.get(group) ?? '';
      emitted.add(group);
      const candidates = doc.blocks.slice(i, i + group.ids.length);
      if (candidates.length === group.ids.length && candidates.every((candidate, index) => candidate.id === group.ids[index])) {
        const raw = source.xml.slice(group.start, group.end);
        if (JSON.stringify(candidates) === group.baseline) {
          body += raw;
          i += group.ids.length - 1;
          continue;
        }
        const previous = JSON.parse(group.baseline) as DocBlock[];
        const patched = previous.length === 1 && candidates.length === 1
          ? patchSourceBlock(raw, previous[0], candidates[0], generated.get(block.id) ?? '') ?? patchText(raw, group.baseline, candidates)
          : patchText(raw, group.baseline, candidates);
        if (patched !== null) {
          body += patched;
          i += group.ids.length - 1;
          continue;
        }
      }
    }
    let xml = generated.get(block.id) ?? '';
    // 功能：编辑分节段落后仍保留其原始 sectPr，避免保存时把连续分节变成普通段落；时间：2026-11-03；作者：lq。
    if (group && block.type === 'paragraph' && block.sectionEnd) {
      const section = source.xml.slice(group.start, group.end).match(/<w:sectPr\b[^>]*\/>|<w:sectPr\b[\s\S]*?<\/w:sectPr>/)?.[0];
      if (section) xml = xml.includes('</w:pPr>') ? xml.replace('</w:pPr>', `${section}</w:pPr>`) : xml.replace('<w:p>', `<w:p><w:pPr>${section}</w:pPr>`);
    }
    body += xml;
  }
  let xml = source.xml.slice(0, source.bodyStart) + body + pending + source.xml.slice(source.bodyEnd);
  // 功能：原件未使用绘图而本次新增图片时，仅补充生成片段需要的命名空间；时间：2026-11-03；作者：lq。
  const namespaces: Record<string, string> = { r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', a: 'http://schemas.openxmlformats.org/drawingml/2006/main', pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture' };
  xml = xml.replace(/<w:document\b[^>]*>/, (root) => {
    for (const [prefix, uri] of Object.entries(namespaces)) {
      if (new RegExp(`(?:<|\\s)${prefix}:`).test(xml) && !root.includes(`xmlns:${prefix}=`)) root = root.slice(0, -1) + ` xmlns:${prefix}="${uri}">`;
    }
    return root;
  });
  return xml;
}

/** 功能：默认字体和样式有变化才更新 styles.xml，普通正文修改保留原样式部件；时间：2026-11-03；作者：lq。 */
export function styleSignature(doc: DocModel): string {
  return JSON.stringify({ styles: doc.styles, defaultFont: doc.defaultFont, defaultFontSizePt: doc.defaultFontSizePt, themeFonts: doc.themeFonts });
}
