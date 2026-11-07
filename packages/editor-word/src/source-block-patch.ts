/** 功能：按原始物理行格及段落定位补丁，保留未建模的 OOXML；时间：2026-11-07；作者：lq。 */
import type { DocBlock, DocParagraph, DocRun, DocTable } from '@mini-office/document-model';
import { scanBody } from '../../platform-adapter/src/vendor/genoffice/xml-scan.js';
import { patchParagraphTexts } from '../../platform-adapter/src/vendor/genoffice/xml-text-patch.js';

/** 功能：记录原始字符串区间，所有替换均直接拼接源字节对应文本；时间：2026-11-07；作者：lq。 */
interface Fragment { name: string; start: number; end: number; xml: string }

/** 功能：跳过引号中的大于号，定位 XML 起始标签；时间：2026-11-07；作者：lq。 */
function openingTag(xml: string): string { return /^<(?:[\w:.-]+)(?:"[^"]*"|'[^']*'|[^'">])*>/.exec(xml)?.[0] ?? ''; }

/** 功能：扫描直接子节点，不把嵌套表格、文本框中的同名元素误当兄弟节点；时间：2026-11-07；作者：lq。 */
function children(xml: string): Fragment[] {
  const open = openingTag(xml);
  if (!open || open.endsWith('/>')) return [];
  const end = xml.lastIndexOf('</');
  if (end < open.length) throw new Error('无法定位 XML 结束标签');
  const prefix = '<w:body>';
  const scan = scanBody(prefix + xml.slice(open.length, end) + '</w:body>');
  return scan.elements.map(element => {
    const start = element.start - prefix.length + open.length;
    const finish = element.end - prefix.length + open.length;
    return { name: element.name, start, end: finish, xml: xml.slice(start, finish) };
  });
}

/** 功能：按降序应用区间替换，避免前面的编辑令后续源锚点偏移；时间：2026-11-07；作者：lq。 */
function splice(xml: string, patches: Array<{ start: number; end: number; xml: string }>): string {
  return patches.sort((a, b) => b.start - a.start).reduce((result, patch) => result.slice(0, patch.start) + patch.xml + result.slice(patch.end), xml);
}

/** 功能：比较可序列化模型属性，忽略 undefined 的键排列差异；时间：2026-11-07；作者：lq。 */
function equal(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (!a || !b || typeof a !== 'object' || typeof b !== 'object') return false;
  const left = a as Record<string, unknown>;
  const right = b as Record<string, unknown>;
  const keys = new Set([...Object.keys(left), ...Object.keys(right)]);
  return [...keys].every(key => equal(left[key], right[key]));
}

/** 功能：已建模属性到 OOXML 属性组的映射，只有实际变化的组允许替换；时间：2026-11-07；作者：lq。 */
const RUN_GROUPS: Record<string, string[]> = {
  fontFamily: ['w:rFonts'], bold: ['w:b'], italic: ['w:i'], allCaps: ['w:caps'], smallCaps: ['w:smallCaps'],
  strike: ['w:strike'], color: ['w:color'], letterSpacingPx: ['w:spacing'], fontSizePt: ['w:sz', 'w:szCs'],
  underline: ['w:u'], highlight: ['w:shd'], verticalAlign: ['w:vertAlign'],
};
/** 功能：段落样式只更新对应属性组，保留未知兼容与节属性；时间：2026-11-07；作者：lq。 */
const PARA_GROUPS: Record<string, string[]> = {
  keepNext: ['w:keepNext'], keepLines: ['w:keepLines'], pageBreakBefore: ['w:pageBreakBefore'], widowControl: ['w:widowControl'],
  snapToGrid: ['w:snapToGrid'], borders: ['w:pBdr'], backgroundColor: ['w:shd'], tabStops: ['w:tabs'],
  spaceBeforePx: ['w:spacing'], spaceAfterPx: ['w:spacing'], lineHeight: ['w:spacing'], lineHeightExactPx: ['w:spacing'], lineHeightAtLeastPx: ['w:spacing'],
  indentLeftPx: ['w:ind'], indentRightPx: ['w:ind'], firstLineIndentPx: ['w:ind'], contextualSpacing: ['w:contextualSpacing'],
  alignment: ['w:jc'], outlineLevel: ['w:outlineLvl'],
};
/** 功能：用规范属性次序插入新组，避免将新增格式附在不合法位置；时间：2026-11-07；作者：lq。 */
const ORDERS: Record<string, string[]> = {
  'w:pPr': 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange'.split(' '),
  'w:rPr': 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange'.split(' '),
  'w:tblPr': 'tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription tblPrChange'.split(' '),
  'w:trPr': 'cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader tblCellSpacing jc hidden ins del trPrChange'.split(' '),
  'w:tcPr': 'cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark headers cellIns cellDel cellMerge tcPrChange'.split(' '),
};

/** 功能：收集发生变化的属性组；无法表达的属性变化返回失败；时间：2026-11-07；作者：lq。 */
function changedGroups(before: object, after: object, mapping: Record<string, string[]>): Set<string> | null {
  const result = new Set<string>();
  const left = before as Record<string, unknown>;
  const right = after as Record<string, unknown>;
  for (const key of new Set([...Object.keys(left), ...Object.keys(right)])) {
    if (equal(left[key], right[key])) continue;
    if (!mapping[key]) return null;
    mapping[key].forEach(name => result.add(name));
  }
  return result;
}

/** 功能：替换或插入单个属性容器中的目标组，原始未知兄弟节点逐字保留；时间：2026-11-07；作者：lq。 */
function patchProperties(raw: string, generated: string, container: string, names: Set<string>, attributes = new Map<string, Set<string>>()): string {
  if (!names.size) return raw;
  const old = children(raw).find(child => child.name === container);
  const fresh = children(generated).find(child => child.name === container);
  let props = old?.xml ?? `<${container}/>`;
  const replacement = fresh ? children(fresh.xml) : [];
  for (const name of names) {
    const matches = children(props).filter(child => child.name === name);
    const value = replacement.filter(child => child.name === name).map(child => child.xml).join('');
    const selected = attributes.get(name);
    if (selected && matches.length === 1) {
      const source = matches[0].xml;
      // 功能：行距和缩进共享一个属性组，仅改用户选择的属性，保留字符缩进及自动间距等其他值；时间：2026-11-07；作者：lq。
      const attrPattern = /\s+([\w:.-]+)\s*=\s*("[^"]*"|'[^']*')/g;
      let updated = source.replace(attrPattern, (whole, key: string) => selected.has(key) ? '' : whole);
      const freshAttrs = [...value.matchAll(attrPattern)].filter(match => selected.has(match[1])).map(match => match[0]).join('');
      const end = openingTag(updated).length;
      const at = updated.slice(0, end).endsWith('/>') ? end - 2 : end - 1;
      updated = updated.slice(0, at) + freshAttrs + updated.slice(at);
      props = splice(props, [{ ...matches[0], xml: updated }]);
      continue;
    }
    if (matches.length) {
      props = splice(props, matches.map((child, index) => ({ ...child, xml: index === 0 ? value : '' })));
    } else if (value) {
      if (props.endsWith('/>')) props = props.slice(0, -2) + `></${container}>`;
      const order = ORDERS[container] ?? [];
      const rank = order.indexOf(name.slice(2));
      const next = children(props).find(child => order.indexOf(child.name.slice(2)) > rank);
      const at = next?.start ?? props.lastIndexOf('</');
      props = props.slice(0, at) + value + props.slice(at);
    }
  }
  if (old) return splice(raw, [{ ...old, xml: props }]);
  const at = openingTag(raw).length;
  // 功能：空段落可能是自闭合标签，需要先展开后再插入属性；时间：2026-11-07；作者：lq。
  if (raw.slice(0, at).endsWith('/>')) return raw.slice(0, at - 2) + `>${props}</${/^<([\w:.-]+)/.exec(raw)![1]}>`;
  return raw.slice(0, at) + props + raw.slice(at);
}

/** 功能：纯文字 run 经浏览器合并后仍比较有效格式，不把分片差异当作格式编辑；时间：2026-11-07；作者：lq。 */
function normalizedRuns(runs: DocRun[]): DocRun[] | null {
  const result: DocRun[] = [];
  for (const run of runs) {
    if (Object.keys(run).some(key => key !== 'text' && key !== 'style')) return null;
    const previous = result.at(-1);
    if (previous && equal(previous.style, run.style)) previous.text = (previous.text ?? '') + (run.text ?? '');
    else result.push({ ...run });
  }
  return result;
}

/** 功能：单一源 run 的粘贴、换行、Tab 和局部格式可拆成多个 run，沿用未知原字符属性；时间：2026-11-07；作者：lq。 */
function patchSimpleRunStream(raw: string, before: DocRun[], after: DocRun[], generated: string): string | null {
  const allowed = new Set(['text', 'style', 'lineBreak', 'tab']);
  if ([...before, ...after].some(run => Object.keys(run).some(key => !allowed.has(key)))) return null;
  if (!before.length || before.some(run => !equal(run.style, before[0].style))) return null;
  const originals = children(raw).filter(child => child.name === 'w:r');
  const fresh = children(generated).filter(child => child.name === 'w:r');
  if (originals.length > 1 || fresh.length !== after.length) return null;
  // 功能：超链接、域及修订中的 run 不参与直接定位；不能将复杂内容当成空段落；时间：2026-11-07；作者：lq。
  const other = children(raw).filter(child => !['w:pPr', 'w:r', 'w:bookmarkStart', 'w:bookmarkEnd'].includes(child.name));
  if (other.length) return null;
  const contentNames = new Set(['w:t', 'w:br', 'w:cr', 'w:tab']);
  const original = originals[0]?.xml ?? '<w:r/>';
  if (children(original).some(child => child.name !== 'w:rPr' && !contentNames.has(child.name))) return null;
  const content = children(original).filter(child => contentNames.has(child.name));
  // 功能：原始文本必须与模型一致，隐藏文本、非文本换行和不识别的属性不允许被展开；时间：2026-11-07；作者：lq。
  const actual = content.map(child => {
    if (child.name === 'w:t') return new DOMParser().parseFromString(`<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${child.xml}</root>`, 'application/xml').documentElement.textContent ?? '';
    if (child.name === 'w:tab') return '\t';
    return /\bw:type\s*=\s*["'](?:page|column)["']/.test(child.xml) ? '\u0000' : '\n';
  }).join('');
  const expected = before.map(run => run.tab ? '\t' : run.lineBreak ? '\n' : run.text ?? '').join('');
  if (actual !== expected) return null;
  let replacement = '';
  for (let index = 0; index < after.length; index++) {
    const groups = changedGroups(before[0].style, after[index].style, RUN_GROUPS);
    if (!groups) return null;
    let run = patchProperties(original, fresh[index].xml, 'w:rPr', groups);
    run = splice(run, children(run).filter(child => contentNames.has(child.name)).map(child => ({ ...child, xml: '' })));
    if (run.endsWith('/>')) run = run.slice(0, -2) + '></w:r>';
    const payload = children(fresh[index].xml).filter(child => child.name !== 'w:rPr').map(child => child.xml).join('');
    const at = run.lastIndexOf('</');
    replacement += run.slice(0, at) + payload + run.slice(at);
  }
  if (originals.length) return splice(raw, [{ ...originals[0], xml: replacement }]);
  if (raw.endsWith('/>')) raw = raw.slice(0, -2) + '></w:p>';
  const at = raw.lastIndexOf('</');
  return raw.slice(0, at) + replacement + raw.slice(at);
}

/** 功能：在结构简单且来源文本完全吻合时保留 run 原属性执行文字补丁；时间：2026-11-07；作者：lq。 */
function patchParagraph(raw: string, before: DocParagraph, after: DocParagraph, generated: string): string | null {
  const { runs: oldRuns, style: oldStyle, markStyle: oldMark, ...oldRest } = before;
  const { runs: newRuns, style: newStyle, markStyle: newMark, ...newRest } = after;
  const groups = changedGroups(oldStyle, newStyle, PARA_GROUPS);
  const restGroups = changedGroups(oldRest, newRest, { styleId: ['w:pStyle'], list: ['w:numPr'] });
  if (!groups || !restGroups) return null;
  // 功能：无编号关系的列表依赖正文前缀生成，不能仅改 numPr 冒充已写入；时间：2026-11-07；作者：lq。
  if (!equal(before.list, after.list) && after.list && !after.list.numId) return null;
  restGroups.forEach(name => groups.add(name));
  let result = raw;
  const markGroups = changedGroups(oldMark ?? {}, newMark ?? {}, RUN_GROUPS);
  if (!markGroups) return null;
  if (!equal(oldRuns, newRuns)) {
    const oldNormalized = normalizedRuns(oldRuns);
    const newNormalized = normalizedRuns(newRuns);
    // 功能：只在全部文字采用同一格式时跨 run 合并补丁，避免跨样式编辑把后段文字变成首段格式；时间：2026-11-07；作者：lq。
    if (oldNormalized?.length === 1 && newNormalized?.length === 1 && equal(oldNormalized[0].style, newNormalized[0].style)) {
      const oldText = oldRuns.map(run => run.text ?? '').join('');
      if (patchParagraphTexts(raw, oldText) !== raw) return null;
      const patched = patchParagraphTexts(raw, newRuns.map(run => run.text ?? '').join(''));
      if (patched === null) return null;
      result = patched;
    } else {
      const simple = patchSimpleRunStream(raw, oldRuns, newRuns, generated);
      if (simple !== null) result = simple;
      else {
        // 功能：多格式段落只允许一对一直接 run 补丁；复杂域、超链接等结构保持保护；时间：2026-11-07；作者：lq。
        const originals = children(raw).filter(child => child.name === 'w:r');
        const fresh = children(generated).filter(child => child.name === 'w:r');
        if (originals.length !== oldRuns.length || oldRuns.length !== newRuns.length || fresh.length !== newRuns.length) return null;
        const patches = [];
        for (let index = 0; index < oldRuns.length; index++) {
          if (equal(oldRuns[index], newRuns[index])) continue;
          const { style: a, ...oldContent } = oldRuns[index];
          const { style: b, ...newContent } = newRuns[index];
          let run = originals[index].xml;
          const runGroups = changedGroups(a, b, RUN_GROUPS);
          if (!runGroups) return null;
          if (!equal(oldContent, newContent)) {
            if (!normalizedRuns([oldRuns[index]]) || !normalizedRuns([newRuns[index]])) return null;
            const wrapper = `<w:p>${run}</w:p>`;
            if (patchParagraphTexts(wrapper, oldRuns[index].text ?? '') !== wrapper) return null;
            const patched = patchParagraphTexts(wrapper, newRuns[index].text ?? '');
            if (patched === null) return null;
            run = patched.slice(5, -6);
          }
          run = patchProperties(run, fresh[index].xml, 'w:rPr', runGroups);
          patches.push({ ...originals[index], xml: run });
        }
        result = splice(raw, patches);
      }
    }
  }
  // 功能：段落结束标记单独补丁，避免空段字号修改丢失或覆盖原标记语言属性；时间：2026-11-07；作者：lq。
  const emptyBefore = oldRuns.length === 1 && !oldRuns[0].text && normalizedRuns(oldRuns);
  const emptyAfter = newRuns.length === 1 && !newRuns[0].text && normalizedRuns(newRuns);
  if (oldMark && emptyBefore && emptyAfter) {
    const emptyGroups = changedGroups({ ...oldMark, ...oldRuns[0].style }, { ...newMark, ...newRuns[0].style }, RUN_GROUPS);
    if (!emptyGroups) return null;
    emptyGroups.forEach(name => markGroups.add(name));
  }
  if (markGroups.size) {
    const old = children(result).find(child => child.name === 'w:pPr');
    const fresh = children(generated).find(child => child.name === 'w:pPr');
    const props = patchProperties(old?.xml ?? '<w:pPr/>', fresh?.xml ?? '<w:pPr/>', 'w:rPr', markGroups);
    if (old) result = splice(result, [{ ...old, xml: props }]);
    else result = patchProperties(result, `<w:p>${props}</w:p>`, 'w:pPr', new Set(['w:rPr']));
  }
  const attributes = new Map<string, Set<string>>();
  const attrGroups: Record<string, [string, string[]]> = {
    spaceBeforePx: ['w:spacing', ['w:before', 'w:beforeLines', 'w:beforeAutospacing']], spaceAfterPx: ['w:spacing', ['w:after', 'w:afterLines', 'w:afterAutospacing']],
    lineHeight: ['w:spacing', ['w:line', 'w:lineRule']], lineHeightExactPx: ['w:spacing', ['w:line', 'w:lineRule']], lineHeightAtLeastPx: ['w:spacing', ['w:line', 'w:lineRule']],
    indentLeftPx: ['w:ind', ['w:left', 'w:start', 'w:leftChars', 'w:startChars']], indentRightPx: ['w:ind', ['w:right', 'w:end', 'w:rightChars', 'w:endChars']],
    firstLineIndentPx: ['w:ind', ['w:firstLine', 'w:hanging', 'w:firstLineChars', 'w:hangingChars']],
  };
  for (const [key, [tag, attrs]] of Object.entries(attrGroups)) {
    if (equal(oldStyle[key as keyof typeof oldStyle], newStyle[key as keyof typeof newStyle])) continue;
    const selected = attributes.get(tag) ?? new Set<string>();
    attrs.forEach(attr => selected.add(attr));
    attributes.set(tag, selected);
  }
  return patchProperties(result, generated, 'w:pPr', groups, attributes);
}

/** 功能：对单格内容按源段落顺序定位，嵌套表格递归补丁；结构不匹配时拒绝猜测；时间：2026-11-07；作者：lq。 */
function patchContents(raw: string, before: DocBlock[], after: DocBlock[], generated: string): string | null {
  const originals = children(raw).filter(child => child.name === 'w:p' || child.name === 'w:tbl');
  const fresh = children(generated).filter(child => child.name === 'w:p' || child.name === 'w:tbl');
  if (originals.length !== before.length || fresh.length !== after.length || before.some(block => block.type === 'pageBreak') || after.some(block => block.type === 'pageBreak')) return null;
  const byId = new Map(before.map((block, index) => [block.id, { block, fragment: originals[index] }]));
  if (new Set(after.map(block => block.id)).size !== after.length) return null;
  const prefixes = new Map<string, string>();
  const present = new Set(after.map(block => block.id));
  // 功能：tcPr 固定保留在所有内容之前，不能随首段锚点移动到新增段落后面；时间：2026-11-07；作者：lq。
  const contentStart = originals[0]?.start ?? raw.lastIndexOf('</');
  let cursor = contentStart;
  let pending = '';
  for (let index = 0; index < before.length; index++) {
    pending += raw.slice(cursor, originals[index].start);
    cursor = originals[index].end;
    if (present.has(before[index].id)) {
      prefixes.set(before[index].id, pending);
      pending = '';
    }
  }
  pending += raw.slice(cursor, raw.lastIndexOf('</'));
  let body = '';
  for (let index = 0; index < after.length; index++) {
    const source = byId.get(after[index].id);
    body += prefixes.get(after[index].id) ?? '';
    if (!source) body += fresh[index].xml;
    else {
      const result = patchSourceBlock(source.fragment.xml, source.block, after[index], fresh[index].xml);
      if (result === null) return null;
      body += result;
    }
  }
  return raw.slice(0, contentStart) + body + pending + raw.slice(raw.lastIndexOf('</'));
}

/** 功能：按物理行格跳过原始纵向合并续格，保留续格的全部 tcPr 和未知内容；时间：2026-11-07；作者：lq。 */
function patchTable(raw: string, before: DocTable, after: DocTable, generated: string): string | null {
  const { rows: oldRows, ...oldProps } = before;
  const { rows: newRows, ...newProps } = after;
  const tableGroups = changedGroups(oldProps, newProps, { widthPct: ['w:tblW'], alignment: ['w:jc'], indentPx: ['w:tblInd'], borders: ['w:tblBorders'], layout: ['w:tblLayout'], cellMarginPx: ['w:tblCellMar'], styleId: ['w:tblStyle'] });
  if (!tableGroups || oldRows.length !== newRows.length) return null;
  const originals = children(raw).filter(child => child.name === 'w:tr');
  const fresh = children(generated).filter(child => child.name === 'w:tr');
  if (originals.length !== oldRows.length || fresh.length !== newRows.length) return null;
  const rowPatches = [];
  for (let r = 0; r < oldRows.length; r++) {
    const { cells: oldCells, ...oldRow } = oldRows[r];
    const { cells: newCells, ...newRow } = newRows[r];
    const rowGroups = changedGroups(oldRow, newRow, { heightPx: ['w:trHeight'], heightRule: ['w:trHeight'], isHeader: ['w:tblHeader'] });
    if (!rowGroups || oldCells.length !== newCells.length) return null;
    const cells = children(originals[r].xml).filter(child => child.name === 'w:tc');
    const generatedCells = children(fresh[r].xml).filter(child => child.name === 'w:tc');
    if (cells.length !== generatedCells.length) return null;
    let logical = 0;
    const cellPatches = [];
    for (let c = 0; c < cells.length; c++) {
      const properties = children(cells[c].xml).find(child => child.name === 'w:tcPr');
      const merge = properties && children(properties.xml).find(child => child.name === 'w:vMerge');
      if (merge && !/\bw:val\s*=\s*["']restart["']/.test(merge.xml)) continue;
      const oldCell = oldCells[logical];
      const newCell = newCells[logical++];
      if (!oldCell || !newCell) return null;
      const { blocks: oldBlocks, ...oldCellProps } = oldCell;
      const { blocks: newBlocks, ...newCellProps } = newCell;
      const cellGroups = changedGroups(oldCellProps, newCellProps, { backgroundColor: ['w:shd'], borders: ['w:tcBorders'], verticalAlign: ['w:vAlign'], marginPx: ['w:tcMar'] });
      if (!cellGroups) return null;
      const content = patchContents(cells[c].xml, oldBlocks, newBlocks, generatedCells[c].xml);
      if (content === null) return null;
      cellPatches.push({ ...cells[c], xml: patchProperties(content, generatedCells[c].xml, 'w:tcPr', cellGroups) });
    }
    if (logical !== oldCells.length) return null;
    const row = splice(originals[r].xml, cellPatches);
    rowPatches.push({ ...originals[r], xml: patchProperties(row, fresh[r].xml, 'w:trPr', rowGroups) });
  }
  return patchProperties(splice(raw, rowPatches), generated, 'w:tblPr', tableGroups);
}

/** 功能：只返回可证明来源匹配的块补丁，null 交由调用者显式拒绝有损保存；时间：2026-11-07；作者：lq。 */
export function patchSourceBlock(raw: string, before: DocBlock, after: DocBlock, generated: string): string | null {
  if (equal(before, after)) return raw;
  if (before.id !== after.id || before.type !== after.type) return null;
  if (before.type === 'paragraph' && after.type === 'paragraph') return patchParagraph(raw, before, after, generated);
  if (before.type === 'table' && after.type === 'table') return patchTable(raw, before, after, generated);
  return null;
}

/** 功能：更新最后段落的节属性时保留其原始起始标签，避免带 rsid 属性的段落定位失败；时间：2026-11-07；作者：lq。 */
export function patchParagraphSection(raw: string, section: string): string {
  return patchProperties(raw, `<w:p><w:pPr>${section}</w:pPr></w:p>`, 'w:pPr', new Set(['w:sectPr']));
}
