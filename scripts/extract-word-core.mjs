/** 功能：从 M0 固定来源提取 Word 运行时闭包并记录每项本地修改；时间：2026-11-08；作者：lq。 */
import { readFileSync, writeFileSync, existsSync, mkdirSync, readdirSync } from 'node:fs';
import { dirname, resolve, relative, join, posix } from 'node:path';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';
import postcss from 'postcss';
import { execFileSync } from 'node:child_process';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const origin = join(root, 'output/word-core-m0.local/upstream');
const target = join(root, 'packages/word-core/src/upstream');
const baseline = JSON.parse(readFileSync(join(root, 'docs/word-core-m0/upstream.json'), 'utf8'));
const archive = join(root, baseline.archive.path);
if (digest(readFileSync(archive)) !== baseline.archive.sha256) throw new Error('固定上游归档校验失败');
const expected = new Map(baseline.files.map(file => [file.path, file.sha256]));
const banner = '/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */\n';
const mappings = [];
const manifestPath = join(root, 'packages/word-core/SOURCE-MANIFEST.json');
const previousFiles = new Map(existsSync(manifestPath) ? JSON.parse(readFileSync(manifestPath, 'utf8')).extractedFiles.map(file => [file.localPath, file.localSha256]) : []);
const externals = new Set();
const aliases = {
  '@genoffice/docx-engine': 'packages/docx-engine/src/index.ts',
  '@genoffice/zip-gate': 'packages/zip-gate/src/index.ts',
  '@genoffice/agent-core/redact-range': 'packages/agent-core/src/redact-range.ts',
  '@genoffice/agent-core/redact-core': 'packages/agent-core/src/redact-core.ts',
  '@genoffice/pptx-engine/custgeom': 'packages/pptx-engine/src/custgeom.ts',
  '@genoffice/pptx-engine/chart-workbook': 'packages/pptx-engine/src/chart-workbook.ts',
  '@genoffice/font-metrics/resolve': 'packages/font-metrics/src/resolve.ts',
  '@genoffice/font-metrics/word-faces.json': 'packages/font-metrics/word-faces.json',
  '@genoffice/font-metrics': 'packages/font-metrics/src/metrics.ts',
  '@genoffice/i18n': 'packages/i18n/src/index.ts',
  '@genoffice/ui': 'packages/ui/src/word-core-shared.ts',
};
const seeds = [
  'packages/docx-engine/src/index.ts',
  'packages/docx-engine/src/vendor/emf-converter/index.d.mts',
  'packages/docx-engine/src/vendor/emf-converter/LICENSE',
  'packages/docx-engine/src/vendor/emf-converter/README.md',
  'packages/i18n/src/font-names.ts',
  'apps/docs/src/renderer/editor/extensions.ts', 'apps/docs/src/renderer/editor/convert.ts',
  'apps/docs/src/renderer/pagination.ts', 'apps/docs/src/renderer/pagination-loop.ts',
  'apps/docs/src/renderer/pagination-plan.ts', 'apps/docs/src/renderer/pagination-index.ts',
  'apps/docs/src/renderer/doc-style-css.ts', 'apps/docs/src/renderer/doc-state.ts',
  'apps/docs/src/renderer/embedded-fonts.ts', 'apps/docs/src/renderer/hf-sections.ts',
  'apps/docs/src/renderer/editor/paste-text.ts', 'apps/docs/src/renderer/editor/paste-web-html.ts',
  'apps/docs/src/renderer/editor/paste-options.ts', 'apps/docs/src/renderer/editor/trailing-table-exit.ts',
  'apps/docs/src/renderer/document-settings.ts',
  // 功能：M2 完整分页接入页底注释的同源 DOM 渲染，不复制上游产品界面；时间：2026-11-09；作者：lq。
  'apps/docs/src/renderer/editor/note-dom.ts',
  // 功能：M2 继承页眉页脚变体真实度量与页面底色的额外段落规则；时间：2026-11-09；作者：lq。
  'apps/docs/src/renderer/hf-heights.ts', 'apps/docs/src/renderer/hf-phantom.ts',
  // 功能：M3 提取同源表格结构、属性、边框和图片插入规则，不引入上游菜单；时间：2026-11-11；作者：lq。
  'apps/docs/src/renderer/editor/table-ops.ts', 'apps/docs/src/renderer/editor/table-properties.ts',
  'apps/docs/src/renderer/editor/table-borders.ts', 'apps/docs/src/renderer/editor/insert-picture.ts',
  'apps/docs/src/renderer/editor/new-table.ts',
  // 功能：页眉页脚直接编辑复用同源富文本转换和布局表格隔离规则；时间：2026-11-11；作者：lq。
  'apps/docs/src/renderer/editor/hf-editor.ts',
  'apps/docs/src/renderer/parse-off-thread.ts', 'apps/docs/src/renderer/parse-worker.ts',
];

// 功能：从文件动作中提取实际文档初始化规则，去掉产品上下文与 I/O，保留同一函数体；时间：2026-11-08；作者：lq。
const actionsPath = 'apps/docs/src/renderer/file-actions.ts';
const actionsText = original(actionsPath).toString('utf8');
const actionsAst = ts.createSourceFile(actionsPath, actionsText, ts.ScriptTarget.Latest, true);
const settingsFunction = actionsAst.statements.find(node => ts.isFunctionDeclaration(node) && node.name?.text === 'applyDocLayoutSettings');
if (!settingsFunction) throw new Error('上游文档初始化函数已变化');
const derived = new Map([['apps/docs/src/renderer/document-settings.ts', {
  source: actionsPath,
  text: `import type { Editor } from '@tiptap/core';
import type { ParsedDocFull } from '@genoffice/docx-engine';
import { setNoteNumFmts } from './note-format';
import { hangulSpaceWideningFor, setAdditiveParagraphSpacing, setHangulSpaceWidening, setLatinViaEastAsiaSlot, latinViaEastAsiaSlotFor } from './line-metrics';
import { docBodyEaFont } from './doc-style-css';
import { setStaticAsciiResolver } from './editor/protected-render';
export ${settingsFunction.getText(actionsAst)}\n`,
}]]);

/** 功能：校验固定来源的原始字节，禁止从漂移源码继续提取；时间：2026-11-08；作者：lq。 */
function original(path) {
  const bytes = readFileSync(join(origin, path));
  // 功能：Word 绘图可能引用 M0 宽范围外的底层几何数据，必须与同一固定归档逐字核对；时间：2026-11-08；作者：lq。
  if (!expected.has(path)) {
    const archived = execFileSync('tar', ['-xOf', archive, `genoffice-${baseline.commit}/${path}`], { maxBuffer: 32 * 1024 * 1024 });
    if (!bytes.equals(archived)) throw new Error(`新增引用与固定归档不一致：${path}`);
    expected.set(path, digest(archived));
  }
  if (expected.get(path) !== digest(bytes)) throw new Error(`来源未记录或已变化：${path}`);
  return bytes;
}
/** 功能：计算源码和本地文件指纹；时间：2026-11-08；作者：lq。 */
function digest(value) { return createHash('sha256').update(value).digest('hex'); }
/** 功能：生成相对路径，不在运行时保留上游 workspace 别名；时间：2026-11-08；作者：lq。 */
function specifier(from, to) {
  let value = posix.relative(posix.dirname(from), to).replace(/\.(?:tsx?|mts)$/, '');
  return value.startsWith('.') ? value : `./${value}`;
}
/** 功能：写入可重复提取文件，拒绝覆盖已有本地手工修改；时间：2026-11-08；作者：lq。 */
function save(path, bytes, source, changes) {
  const output = join(target, path);
  mkdirSync(dirname(output), { recursive: true });
  if (existsSync(output) && !readFileSync(output).equals(Buffer.from(bytes)) && previousFiles.get(relative(root, output)) !== digest(readFileSync(output))) throw new Error(`存在未包含在提取规则中的本地修改：${path}`);
  writeFileSync(output, bytes);
  mappings.push({ upstreamPath: source, localPath: relative(root, output), upstreamSha256: source ? expected.get(source) : null, localSha256: digest(bytes), modifications: changes });
}
/** 功能：解析相对导入，覆盖 TypeScript 模块和运行资源；时间：2026-11-08；作者：lq。 */
function resolveModule(path) {
  const found = [path, path.replace(/\.js$/, '.ts'), `${path}.ts`, `${path}.tsx`, `${path}.json`, `${path}/index.ts`].find(candidate => existsSync(join(origin, candidate)) && !readdirDirectory(candidate));
  if (!found) throw new Error(`无法定位上游导入：${path}`);
  return found;
}
/** 功能：区分模块文件和目录，避免把目录当作资源复制；时间：2026-11-08；作者：lq。 */
function readdirDirectory(path) { try { readdirSync(join(origin, path)); return true; } catch { return false; } }

// 功能：仅保留 Word 绘制与字体候选的数据函数，排除上游 React UI 总入口；时间：2026-11-08；作者：lq。
const shared = 'packages/ui/src/word-core-shared.ts';
const sharedText = banner + "export * from './wordart-presets';\nexport * from './font-list';\n";
save(shared, sharedText, null, ['新增 Word 使用的纯数据子入口，未引入 GenOffice UI 组件']);
const queue = [...seeds, 'packages/ui/src/wordart-presets.ts', 'packages/ui/src/font-list.ts'];
const seen = new Set([shared]);
while (queue.length) {
  const path = queue.shift();
  if (seen.has(path)) continue;
  seen.add(path);
  const sourcePath = derived.get(path)?.source ?? path;
  let bytes = derived.has(path) ? Buffer.from(derived.get(path).text) : original(path);
  const changes = derived.has(path) ? ['从 file-actions.ts 提取 applyDocLayoutSettings 原函数体，显式列出依赖，不复制 UI 和 Electron I/O'] : [];
  if (/\.(?:ts|tsx|mts|mjs)$/.test(path)) {
    let text = bytes.toString('utf8');
    // 功能：压缩扩展共享一次自然流窗口，避免逐段隐藏/恢复页缝导致长文档数十次强制重排；时间：2026-11-11；作者：lq。
    if (path === 'apps/docs/src/renderer/editor/settled-measure.ts') {
      text = text.replace('  const gaps = el.querySelectorAll<HTMLElement>', "  if (el.closest('.measuring-natural')) return fn()\n  const gaps = el.querySelectorAll<HTMLElement>");
      text = text.replace('    return map\n', `    // 功能：长表格内的段落也一次建立节点索引，避免逐段 posAtDOM 在整表中重复遍历；时间：2026-11-11；作者：lq。
    for (const element of view.dom.querySelectorAll<HTMLElement>('p, div.doc-textbox-para')) {
      const desc = (element as { pmViewDesc?: { node?: ProseMirrorNode } }).pmViewDesc
      if (desc?.node?.isTextblock) map.set(desc.node, element)
    }
    return map
`);
      changes.push('已进入全篇自然流时不逐段改写页缝 DOM，避免布局抖动');
    }
    if (/editor\/(?:cjk-punct-shrink|justify-shrink)\.ts$/.test(path)) {
      // 功能：字符测量保留完整正文，按帧预算处理；失效、编辑和销毁取消旧任务，完成后一次提交装饰；时间：2026-11-11；作者：lq。
      text = text.replace('  private lastSig =', `  private workRaf = 0
  private destroyed = false
  private work: { doc: ProseMirrorNode; paras: Array<{ node: ProseMirrorNode; pos: number }>; shrinks: MeasuredShrink[]; next: number; measurable: boolean; topLevel: Map<ProseMirrorNode, HTMLElement> } | null = null
  private lastSig =`);
      text = text.replaceAll('this.measure()', 'this.queueMeasure()');
      text = text.replace('  private restartConvergence() {', '  private restartConvergence() {\n    this.work = null');
      text = text.replace('  destroy() {', '  destroy() {\n    this.destroyed = true\n    this.work = null\n    if (this.workRaf) cancelAnimationFrame(this.workRaf)');
      text = text.replace('  private measure() {', `  /** 功能：同一帧的字体、样式和文档变更仅触发一次测量；时间：2026-11-11；作者：lq。 */
  private queueMeasure() {
    if (this.destroyed || this.workRaf) return
    this.workRaf = requestAnimationFrame(() => { this.workRaf = 0; if (!this.destroyed) this.measure() })
  }

  private measure() {`);
      const start = text.indexOf('    const paras: Array<{ node: ProseMirrorNode; pos: number }> = []\n', text.indexOf('  private measure() {'));
      const end = text.indexOf('    if (!measurable) {', start);
      if (start < 0 || end < 0) throw new Error(`字符分帧提取位置失效：${path}`);
      const candidate = path.includes('cjk-punct') ? 'isCandidate' : 'isShrinkCandidate';
      text = text.slice(0, start) + `    // 功能：一轮中共享自然流几何；所有文档超过 8ms 让出帧，缓存已测段落且不提交不完整结果；时间：2026-11-11；作者：lq。
    let shrinks: MeasuredShrink[] = []
    let measurable = false
    const natural = view.dom.classList.contains('measuring-natural')
    view.dom.classList.add('measuring-natural')
    try {
      let work = this.work
      if (!work || work.doc !== view.state.doc) {
        const paras: Array<{ node: ProseMirrorNode; pos: number }> = []
        view.state.doc.descendants((node, pos) => {
          if (!node.isTextblock) return true
          if (${candidate}(node)) paras.push({ node, pos })
          return false
        })
        this.results.beginPass(view)
        work = this.work = { doc: view.state.doc, paras, shrinks: [], next: 0, measurable: paras.length === 0, topLevel: SettledParagraphCache.topLevelDom(view) }
      }
      const started = performance.now()
      while (work.next < work.paras.length) {
        const para = work.paras[work.next++]
        const measured = this.results.measure(view, para.node, para.pos, el => this.measureParagraph(para.node, para.pos, el), work.topLevel.get(para.node))
        if (measured) { work.measurable = true; work.shrinks.push(...measured) }
        if (work.next < work.paras.length && performance.now() - started >= 8) { this.queueMeasure(); return }
      }
      shrinks = work.shrinks; measurable = work.measurable; this.work = null
    } finally { if (!natural) view.dom.classList.remove('measuring-natural') }
` + text.slice(end);
      changes.push('字符压缩合并到帧调度并按 8ms 预算分批，原模型完整，原子提交装饰且失效可取消');
    }
    // 功能：只保留表格原生命令，错误提示交给 Mini Office 现有界面；时间：2026-11-11；作者：lq。
    if (path === 'apps/docs/src/renderer/editor/table-ops.ts') {
      text = text.replace("import { showToast } from '../components/toast-bus'\n", '')
        .replace("import { t } from '../i18n/locale'\n", '')
        .replace("  showToast(t('ribbonReadOnlyEdit'), 'error')\n", '');
      text = text.replace('{ ...first.attrs, docxIndex: null }', '{ ...first.attrs, docxIndex: null, miniOfficeCellSource: null }');
      changes.push('移除上游 Toast/UI 依赖，只读命令返回 false，由本项目 UI 反馈');
    }
    // 功能：不同纸张/分节的表格尺寸按其所在正文净宽，避免整页含边距宽度泄入网格；时间：2026-11-11；作者：lq。
    if (path === 'apps/docs/src/renderer/editor/table-column-drag.ts' || path === 'apps/docs/src/renderer/editor/table-handle.ts') {
      text = text.replace('function sectionContentWidth(view: EditorView, fallback: number): number {', 'function sectionContentWidth(view: EditorView, fallback: number, table: HTMLElement): number {')
        .replace("getComputedStyle(view.dom).getPropertyValue('--section-content-w')", "getComputedStyle(table).getPropertyValue('--section-content-w')")
        .replace('sectionContentWidth(view, table.parentElement?.clientWidth || 624)', 'sectionContentWidth(view, table.parentElement?.clientWidth || 624, table)');
      changes.push('表格拖动尺寸采用所在分节的正文净宽，不能采用含页边距的整页宽');
    }
    // 功能：表格结构/属性改变时保留未修改段落的书签与未知 XML，来源随段落移动；时间：2026-11-11；作者：lq。
    if (path === 'packages/docx-engine/src/types.ts') {
      // 功能：页脚框的原形状身份用于局部保存，不能按显示序号重建文本框；时间：2026-11-11；作者：lq。
      text = text.replace('  id: number\n  /** wps:bodyPr', '  /** 功能：页眉页脚框的来源形状标识，保留 DrawingML/VML 原定位；时间：2026-11-11；作者：lq。 */\n  miniOfficeShapeId?: string\n  id: number\n  /** wps:bodyPr');
      text = text.replace("  | 'align'\n  | 'indentLeft'", "  | 'listMarker'\n  | 'align'\n  | 'indentLeft'");
      text = text.replace('export interface ParaFormat {', 'export interface ParaFormat {\n  /** 功能：区域原生样式引用参与保存；时间：2026-11-11；作者：lq。 */\n  styleId?: string\n  listMarker?: TextboxListMarker\n');

      text = text.replace('export interface ParaBorderLine {', `export interface ParaBorderLine {
  /** 功能：段落样式/格式刷保留双线、虚线和点线类型；时间：2026-11-11；作者：lq。 */
  val?: string`);
      text = text.replace('export interface HeaderFooter {', `export interface HeaderFooter {
  /** 功能：解除继承时从来源部件克隆 XML/关系，不重建图片或未知内容；时间：2026-11-11；作者：lq。 */
  miniOfficeSourceRid?: string`);
      text = text.replace('export interface TableParagraph extends ParaFormat {', `export interface TableParagraph extends ParaFormat {
  /** 功能：结构回写时复用未编辑的格内段落 XML；时间：2026-11-11；作者：lq。 */
  sourceXml?: string`);
      // 功能：布局表格段落使用完整原生格式，右缩进、分页选项和边框也能往返；时间：2026-11-11；作者：lq。
      text = text.replace('export type HfCellParaProps = Pick<', 'export type HfCellParaProps = ParaFormat & { indentFirstLineChars?: number } & Pick<');
      changes.push('TableParagraph 增加非目标格内段落 XML 来源，供结构回写保留');
    }
    // 功能：边框样式在解析、显示、保存和未改 XML 对比使用同一字段；时间：2026-11-11；作者：lq。
    if (path === 'packages/docx-engine/src/parse-props.ts') {
      text = text.replace('    const line: ParaBorderLine = {}', "    const line: ParaBorderLine = {}\n    if (val && val !== 'single') line.val = val");
      text = text.replace('if (line.color !== undefined || line.szPt !== undefined || line.spacePt !== undefined)', 'if (line.val !== undefined || line.color !== undefined || line.szPt !== undefined || line.spacePt !== undefined)');
      changes.push('段落边框保留非单线类型');
    }
    if (path === 'apps/docs/src/renderer/editor/para-border-css.ts') {
      text = text.replace('line?: { color?: string; szPt?: number }', 'line?: { color?: string; szPt?: number; val?: string }');
      text = text.replace(' solid ', " ${line?.val === 'double' ? 'double' : line?.val === 'dotted' ? 'dotted' : /dash/i.test(line?.val ?? '') ? 'dashed' : 'solid'} ");
      changes.push('段落边框按真实双线/虚线/点线渲染');
    }
    // 功能：页眉页脚副本保留所属部件的图片/链接关系，避免新节丢图或改到共享原件；时间：2026-11-11；作者：lq。
    if (path === 'apps/docs/src/renderer/hf-sections.ts') {
      text = text.replace('    if (edited) return edited', '    if (edited && !edited.miniOfficeSourceRid) return edited');
      text = text.replace('  const { pictureOnly: _textKept, ...rest } = next', '  const { pictureOnly: _textKept, keepContent: _kept, ...rest } = next\n  if (prev?.miniOfficeSourceRid && !rest.miniOfficeSourceRid) rest.miniOfficeSourceRid = prev.miniOfficeSourceRid');
      text = text.replace("const images = rId\n    ? parts?.[rId]?.images", "const sourceRid = state.edits[editKey(sec?.lastBlockIndex ?? -1, kind, variant)]?.miniOfficeSourceRid ?? rId\n  const images = sourceRid\n    ? parts?.[sourceRid]?.images");
      text = text.replace("  const copy = resolveHf(state, si, kind, variant).value ?? { text: '' }", `  const resolved = resolveHf(state, si, kind, variant)
  const inherited = resolved.value ?? { text: '' }
  const sourceRid = inherited.miniOfficeSourceRid ?? refOf(state.sections[resolved.owner], kind, variant)
  const copy = { ...inherited, ...(sourceRid ? { miniOfficeSourceRid: sourceRid } : {}), ...(!Object.values(state.edits).includes(inherited) ? { keepContent: true as const } : {}) }`);
      changes.push('解除继承记录原部件来源，克隆图片关系及未知 XML；未编辑的副本按原字节保存');
    }
    // 功能：页缝填充只占未被纵向合并格占据的列段，测量时恢复源 rowspan；时间：2026-11-11；作者：lq。
    if (path === 'apps/docs/src/renderer/editor/pagination-gaps.ts') {
      text = text.replace('export type PageGapSpec = {', `export type PageGapSpec = {
  /** 功能：装饰行可用的连续列段，避免合并格旁边新增幽灵列；时间：2026-11-11；作者：lq。 */
  tableCellSpans?: number[]`);
      text = text.replace('${gap.cols ?? 0},${Math.round(metrics.pageLeft ?? -1)}', '${gap.cols ?? 0},${gap.tableCellSpans?.join(\':\') ?? \'\'},${Math.round(metrics.pageLeft ?? -1)}');
      text = text.replace('      const el = makeGapEl(metrics, as, gap.cols)', `      const el = makeGapEl(metrics, as, gap.cols)
      // 功能：跨页合并格已经占据部分网格列，填充格仅覆盖空闲连续列段；时间：2026-11-11；作者：lq。
      if (as === 'table' && gap.tableCellSpans?.length) {
        const first = el.firstElementChild as HTMLTableCellElement
        first.colSpan = gap.tableCellSpans[0]
        for (const span of gap.tableCellSpans.slice(1)) {
          const cell = document.createElement('td')
          cell.colSpan = span
          cell.contentEditable = 'false'
          el.appendChild(cell)
        }
      }`);
      text = text.replace('export function syncPhantomRowspans(root: HTMLElement): void {', `export function syncPhantomRowspans(root: HTMLElement): void {
  // 功能：自然流隐藏装饰行后必须同步恢复源跨度，否则被隐藏行仍导致后续真实行列错位；时间：2026-11-11；作者：lq。
  if (root.classList.contains('measuring-natural')) {
    for (const cell of root.querySelectorAll<HTMLTableCellElement>('[data-base-rowspan]')) cell.rowSpan = Number(cell.dataset.baseRowspan) || 1
    return
  }`);
      changes.push('页缝仅占空闲列段；自然流度量恢复源 rowspan，显示后桥接装饰行');
    }
    if (path === 'apps/docs/src/renderer/editor/hf-dom.ts') {
      // 功能：探针继承产品隔离后的文档排版 CSS，浮动页码框不再退化为 192px 的普通流块；时间：2026-11-11；作者：lq。
      text = text.replace('    document.body.appendChild(host)', `    const scope = document.createElement('div')
    scope.className = 'mo-word-core'
    scope.appendChild(host)
    document.body.appendChild(scope)`);
      changes.push('页眉页脚探针置于文档 CSS 隔离作用域，正确排除浮动文本框的流高度');
      // 功能：双击浮动页码时定位真实框，不能退回外层空段落；时间：2026-11-11；作者：lq。
      text = text.replace('        boxHost.className = hfTextBoxClass(para.box!)', "        boxHost.className = hfTextBoxClass(para.box!)\n        boxHost.dataset.hfBox = String(para.box!.id)");
      // 功能：页眉页脚格内纸面与直接编辑区共同显示右缩进、自动段距、底纹及段落边框；时间：2026-11-11；作者：lq。
      const cellBoxStart = text.indexOf('export function hfCellParaBox('), cellBoxEnd = text.indexOf('export function hfCellParaStyle', cellBoxStart);
      let cellBox = text.slice(cellBoxStart, cellBoxEnd);
      cellBox = cellBox.replace('        spaceBeforeTwips:', '        spaceBeforeAuto: props.spaceBeforeAuto,\n        spaceAfterAuto: props.spaceAfterAuto,\n        spaceBeforeTwips:')
        .replace('        indentLeftTwips:', '        indentRightTwips: props.indentRight || null,\n        indentLeftTwips:')
        .replace('  return { style, classes: box.classes }', `  // 功能：多行制表位只在首尾绘制横边，左右边和底纹延续；时间：2026-11-11；作者：lq。
  const border = hfParaBorderStyle({ ...props, runs: [] }, { suppressTop: !line.first, suppressBottom: !line.last })
  Object.assign(style, border.style)
  const fill = props.shadingDisplay ?? props.shadingFill
  if (fill) style.backgroundColor = '#' + fill
  return { style, classes: box.classes }`);
      text = text.slice(0, cellBoxStart) + cellBox + text.slice(cellBoxEnd);
      changes.push('页眉页脚格内纸面显示完整段落格式');
      // 功能：纸面区域显示真实列表标记，复用文本框相同的悬挂布局；时间：2026-11-11；作者：lq。
      text = text.replace("    p.className = 'page-hf-para'", `    p.className = 'page-hf-para'
    if (para.listMarker) {
      p.classList.add('doc-textbox-para')
      p.dataset.marker = para.listMarker.text
      p.style.setProperty('--li-hang', String((para.listMarker.hanging ?? 360) / 20) + 'pt')
    }`);
      changes.push('页眉页脚列表标记与直接编辑区同步显示');
    }
    if (path === 'apps/docs/src/renderer/editor/hf-editor.ts') {
      // 功能：页眉页脚原生段落保留样式引用和格式刷边框，JSON 仅用于编辑属性边界；时间：2026-11-11；作者：lq。
      text = "import { paraBorderCss } from './para-border-css';\n" + text;
      text = text.replace("  'bidi',\n] as const", "  'bidi', 'listMarker', 'keepNext', 'keepLines', 'widowControl', 'pageBreakBefore',\n] as const");
      // 功能：普通区域段落与布局表格格内使用同一组格式键，不静默丢弃工具栏字段；时间：2026-11-11；作者：lq。
      text = text.replace(/const HF_CELL_PARA_KEYS = \[[\s\S]*?\] as const/, "const HF_CELL_PARA_KEYS = [...HF_PARA_KEYS, 'wordWrap'] as const");
      text = text.replace('  if (props?.tabStops?.length)', '  if (props?.borderLines) attrs.borderLines = JSON.stringify(props.borderLines)\n  if (props?.tabStops?.length)');
      text = text.replace("    if (typeof p.attrs?.tabStops === 'string')", "    if (typeof p.attrs?.borderLines === 'string') props.borderLines = JSON.parse(p.attrs.borderLines)\n    else delete props.borderLines\n    if (typeof p.attrs?.tabStops === 'string')");
      text = text.replace('      bidi: { default:', `      borderLines: { default: null as string | null },
      keepNext: { default: null as boolean | null }, keepLines: { default: null as boolean | null },
      widowControl: { default: null as boolean | null }, pageBreakBefore: { default: null as boolean | null },
      bidi: { default:`);
      text = text.replace('    if (props.node.attrs.bidi)', `    // 功能：直接编辑区域的段落边框与纸面使用同一线型；时间：2026-11-11；作者：lq。
    if (props.node.attrs.borderLines) {
      const lines = JSON.parse(props.node.attrs.borderLines as string)
      const borders = Object.entries({ t: 'top', r: 'right', b: 'bottom', l: 'left' }).map(([side, name]) => lines[side] ? 'border-' + name + ':' + (lines[side].none ? 'none' : paraBorderCss(lines[side])) : '').filter(Boolean).join(';')
      attrs.style = (attrs.style ? attrs.style + ';' : '') + borders
    }
    if (props.node.attrs.bidi)`);
      text = text.replace('    if (para.tabStops?.length)', '    if (para.borderLines) attrs.borderLines = JSON.stringify(para.borderLines)\n    if (para.tabStops?.length)');
      text = text.replace('  if (typeof attrs.tabStops', '  if (typeof attrs.borderLines === \'string\') para.borderLines = JSON.parse(attrs.borderLines)\n  else delete para.borderLines\n  if (typeof attrs.tabStops');

      text += `
/** 功能：受控区域复用动态页码装饰，更新不产生内容或历史；时间：2026-11-11；作者：lq。 */
export function setHfFieldNumbers(editor: Editor, page: string, total: string, section: string): void {
  // 功能：相同域值不再派发装饰事务，避免每轮重排触发工具栏和区域重绘；时间：2026-11-11；作者：lq。
  const previous = hfFieldNumbersKey.getState(editor.state)
  if (previous?.page === page && previous.total === total && previous.section === section) return
  editor.view.dispatch(editor.state.tr.setMeta(hfFieldNumbersKey, { page, total, section }).setMeta('addToHistory', false))
}
`;
      changes.push('导出动态页码装饰更新函数，供单一主历史的受控编辑区使用');
    }
    if (path === 'packages/docx-engine/src/patch.ts') {
      // 功能：富页脚只在原框内部写回，框内文本不能再作为普通段落重复输出；时间：2026-11-11；作者：lq。
      text = "import { patchWordHfTextboxes } from '../../../../serialization/header-footer-boxes';\n" + text;
      text = text.replace('    return headerFooterPartXml(kind, hf, undefined, base, options.pageCount)', '    return headerFooterPartXml(kind, hf, undefined, base ? patchWordHfTextboxes(base, hf, options.pageCount) : base, options.pageCount)');
      text = text.replace('.filter((para) => !(para.lineOnly && originalXml && para.runs.length === 0))', '.filter((para) => !(originalXml && (para.boxAnchored || para.box || (para.lineOnly && para.runs.length === 0))))');
      changes.push('页眉页脚文本框原位保存及备用分支同步，保留框几何，去掉重复的普通页码段落');
      const anchor = '  const planHeaderFooter = async (';
      if (!text.includes(anchor)) throw new Error('页眉页脚部件入口已变化');
      text = text.replace(anchor, `  /** 功能：副本使用独立部件和独立关系，内部目标转包内绝对路径以兼容子目录；时间：2026-11-11；作者：lq。 */
  const cloneHfSource = async (hf: HeaderFooter, destination: string): Promise<string | null> => {
    if (!hf.miniOfficeSourceRid) return null
    const target = relTargets.get(hf.miniOfficeSourceRid)
    const path = target ? resolveRelationshipTargetPath(docPath, target) : null
    const file = path ? zip.file(path) : null
    if (!path || !file) throw new Error('无法定位页眉页脚来源部件，拒绝丢失内容')
    const slash = path.lastIndexOf('/')
    const relsPath = path.slice(0, slash + 1) + '_rels/' + path.slice(slash + 1) + '.rels'
    const relationships = await zip.file(relsPath)?.async('string')
    if (relationships) {
      const rebased = relationships.replace(/<Relationship\\b[^>]*\\/>/g, tag => {
        if (/TargetMode\\s*=\\s*['"]External['"]/.test(tag)) return tag
        return tag.replace(/(\\bTarget\\s*=\\s*)(['"])(.*?)\\2/, (whole, prefix, quote, value) => {
          const resolved = resolveRelationshipTargetPath(path, value)
          if (!resolved) throw new Error('无效的页眉页脚关系目标')
          return prefix + quote + '/' + resolved + quote
        })
      })
      const destSlash = destination.lastIndexOf('/')
      hfRelsOut.set(destination.slice(0, destSlash + 1) + '_rels/' + destination.slice(destSlash + 1) + '.rels', rebased)
    }
    return hfOriginalXml(hf, await file.async('string'))
  }
` + anchor);
      text = text.replace('const target = rId ? relTargets.get(rId) : undefined\n    if (target)', 'const target = !hf.miniOfficeSourceRid && rId ? relTargets.get(rId) : undefined\n    if (target)');
      text = text.replace('        headerFooterPartXml(kind, hf, undefined, null, options.pageCount),', '        editedHfPartXml(kind, hf, await cloneHfSource(hf, `word/${filename}`)),');
      text = text.replace('const target = rId ? relTargets.get(rId) : undefined\n    if (target)', 'const target = !edit.hf.miniOfficeSourceRid && rId ? relTargets.get(rId) : undefined\n    if (target)');
      text = text.replace('        headerFooterPartXml(edit.kind, edit.hf, undefined, null, options.pageCount),', '        editedHfPartXml(edit.kind, edit.hf, await cloneHfSource(edit.hf, `word/${filename}`)),');
      text = text.replace("    let out = xml\n    for (const u of options.sectionHfUnlink", `    let out = xml
    const copies = (options.sectionHf ?? []).filter(edit => edit.lastBlockIndex === docxIndex && edit.hf.miniOfficeSourceRid)
    for (const edit of copies) out = removeHfReference(out, edit.kind, edit.variant ?? 'default')
    for (const u of options.sectionHfUnlink`);
      // 新建副本注入前移除旧引用，不能把刚注入的新引用一起删除。
      text = text.replace('    if (refTags && refTags.length > 0) {\n      xml = injectIntoSectPr', '    if (fbDocxIndex !== undefined) xml = unlinkSectionHf(xml, fbDocxIndex)\n    if (refTags && refTags.length > 0) {\n      xml = injectIntoSectPr');
      text = text.replace('    if (fbDocxIndex !== undefined) xml = unlinkSectionHf(xml, fbDocxIndex)\n    // The ink', '    // The ink');
      text = text.replace('        if (hfRefTags.length > 0) {', `        xml = unlinkSectionHf(xml, block.docxIndex!)
        for (const [kind, hf, variant] of [
          ['header', options.header, 'default'], ['footer', options.footer, 'default'],
          ['header', options.headerFirst, 'first'], ['footer', options.footerFirst, 'first'],
          ['header', options.headerEven, 'even'], ['footer', options.footerEven, 'even'],
        ] as const) if (hf?.miniOfficeSourceRid) xml = removeHfReference(xml, kind, variant)
        if (hfRefTags.length > 0) {`);
      text = text.replace('        xml = unlinkSectionHf(xml, block.docxIndex)\n', '');
      changes.push('页眉页脚独立副本克隆来源 XML 和关系，替换节引用而不污染继承所有者');
    }
    if (path === 'packages/docx-engine/src/parse.ts') {
      // 功能：记录原形状身份供页眉页脚框局部写回，避免空框/多框导致序号错配；时间：2026-11-11；作者：lq。
      text = text.replace('  const attrs = attrsOf(anchor)\n  if (attrs[\'behindDoc\']', "  const shapeId = attrsOf(findChild(anchor, 'wp:docPr') ?? {})['id']\n  if (shapeId) box.miniOfficeShapeId = 'drawing:' + shapeId\n  const attrs = attrsOf(anchor)\n  if (attrs['behindDoc']");
      text = text.replace("  const style = attrsOf(shape)['style'] ?? ''\n  const px", "  const shapeId = attrsOf(shape)['id']\n  if (shapeId) box.miniOfficeShapeId = 'vml:' + shapeId\n  const style = attrsOf(shape)['style'] ?? ''\n  const px");
      // 功能：区域布局表格采用完整段落格式解析，与正文/区域工具栏保存字段一致；时间：2026-11-11；作者：lq。
      text = text.replace("const props: NonNullable<HfTableCell['paraProps']>[number] = {", "const props: NonNullable<HfTableCell['paraProps']>[number] = {\n      ...hfStyledParaFormat(node, ctx),");
      // 功能：正文文本框的嵌套图片同表格一样预解析关系，重开后继续富文本编辑不丢图；时间：2026-11-11；作者：lq。
      text = text.replace("if (el.name !== 'w:tbl' && el.name !== 'w:sdt') continue", "if (el.name !== 'w:tbl' && el.name !== 'w:sdt' && !documentXml.slice(el.start, el.end).includes('<w:txbxContent')) continue")
        .replace("if (el.name === 'w:sdt') {\n      const from = slice.indexOf", "if (el.name === 'w:sdt' && !slice.includes('<w:txbxContent')) {\n      const from = slice.indexOf");
      changes.push('文本框图片预解析本部件关系以保留重开后的富文本图片');
      text = text.replace('    ...withoutResolvedGeometry(direct),', `    // 功能：区域列表编号身份及几何随原生段落往返；时间：2026-11-11；作者：lq。
    ...(eff.list ? { listMarker: { numId: eff.list.numId, ilvl: eff.list.ilvl, kind: eff.list.level?.numFmt === 'bullet' ? 'bullet' as const : 'ordered' as const, text: eff.list.level?.numFmt === 'bullet' ? eff.list.level.lvlText : '1.', indentLeft: eff.list.level?.indentLeft, hanging: eff.list.level?.hanging } } : {}),
    ...withoutResolvedGeometry(direct),`);
      const hfStart = text.indexOf('function hfParagraphs('), hfEnd = text.indexOf('type HfRawCell', hfStart);
      let hfBody = text.slice(hfStart, hfEnd);
      hfBody = hfBody.replace('  return out\n}', `  const listed = out.filter(para => para.listMarker?.numId)
  const infos = computeListMarkerInfos(listed.map(para => ({ numId: para.listMarker!.numId ?? null, ilvl: para.listMarker!.ilvl ?? 0 })), resolver?.numbering ?? new Map())
  listed.forEach((para, i) => { if (infos[i]) para.listMarker = { ...para.listMarker!, text: infos[i].text } })
  return out
}`);
      text = text.slice(0, hfStart) + hfBody + text.slice(hfEnd);
      text = text.replace('    ...withoutResolvedGeometry(direct),', "    ...(pPr && attrsOf(findChild(pPr, 'w:pStyle') ?? {})['w:val'] ? { styleId: attrsOf(findChild(pPr, 'w:pStyle') ?? {})['w:val'] } : {}),\n    ...withoutResolvedGeometry(direct),");
      changes.push('页眉页脚普通段落保留原生样式引用，套用保存重开不丢失身份');

      text = text.replace('    richParas.push({\n      ...format,', `    richParas.push({
      // 功能：保留格内段落的书签、链接外壳及未知子节点；时间：2026-11-11；作者：lq。
      sourceXml: serializeXNode(p),
      ...format,`);
      changes.push('格内段落携带来源 XML，结构编辑不再必然重建未修改段落');
    }
    if (path === 'apps/docs/src/renderer/editor/convert.ts') {
      // 功能：文本框把新图存为 rich-run，保存快照重新登记媒体并重分配占位符，隔离连续保存与其他会话；时间：2026-11-11；作者：lq。
      text = text.replace('const doc = normalizeZoteroFieldParts(inputDoc)', 'const doc = normalizeZoteroFieldParts(registerTextboxPictures(inputDoc))');
      text += `
/** 功能：为文本框历史中的未落盘图片登记当前保存媒体，保留尺寸和绘图 XML；时间：2026-11-11；作者：lq。 */
function registerTextboxPictures(node: PmNode): PmNode {
  const boxes = node.attrs?.textboxes as TextboxDisplay[] | undefined
  const textboxes = boxes?.map(box => ({ ...box, paras: box.paras.map(para => ({ ...para, runs: para.runs.map(run => {
    const image = run.image
    if (!image?.xml.includes(NEW_PICTURE_TOKEN_PREFIX)) return run
    const bytes = newPictureBytes(image.dataUrl)
    if (!bytes) throw new Error('文本框图片数据缺失，已取消保存')
    const token = NEW_PICTURE_TOKEN_PREFIX + ++newPictureSeq
    newPictures.set(token, bytes)
    return { ...run, image: { ...image, xml: image.xml.replace(new RegExp(NEW_PICTURE_TOKEN_PREFIX + '[0-9]+', 'g'), token) } }
  }) })) }))
  return { ...node, ...(textboxes ? { attrs: { ...node.attrs, textboxes } } : {}), ...(node.content ? { content: node.content.map(registerTextboxPictures) } : {}) }
}
`;
      changes.push('文本框新插图在保存快照重登记媒体，历史和不同会话不共享占位符');
      // 功能：schema 补齐空属性或属性声明顺序变化不算格式编辑，避免无操作表格重建；时间：2026-11-11；作者：lq。
      text = text.replace('.map((mark) => [mark.type, mark.attrs ?? null])', '.map((mark) => [mark.type, mark.attrs ? Object.fromEntries(Object.entries(mark.attrs).filter(([, value]) => value !== null && value !== undefined).sort(([a], [b]) => a.localeCompare(b))) : null])');
      changes.push('表格字符格式签名统一空属性和键顺序，扩展属性不触发无操作重写');

      text = `import { patchWordParagraphStructure } from '${specifier(`upstream/${path}`, 'serialization/paragraph-structure.ts')}';\n` + text;
      text = text.replace('  if (!structureSame) return\n  const formatSame', '  const formatSame');
      text = text.replace('  // Strip pPrChange when', '  if (!structureSame) rawPPr = patchWordParagraphStructure(mergePPrFormat(rawPPr, generated.format, original.format), generated)\n  // Strip pPrChange when');
      changes.push('样式/列表语义切换局部更新 pPr，保留分节及未知属性');

      text = `import { patchWordRunEffects } from '${specifier(`upstream/${path}`, 'serialization/run-effects.ts')}';\n` + text;
      text = text.replace('      if (mark.attrs?.rawRPr) run.rawRPr = String(mark.attrs.rawRPr)', '      if (mark.attrs?.rawRPr) run.rawRPr = String(mark.attrs.rawRPr)\n      if (typeof mark.attrs?.miniOfficeEffectXml === \'string\') run.rawRPr = patchWordRunEffects(run.rawRPr, mark.attrs.miniOfficeEffectXml)');
      text = text.replace("  if (run.underline) marks.push({ type: 'underline' })", "  if (run.underline ?? run.effective?.underline) marks.push({ type: 'underline' })")
        .replace("  if (run.strike) marks.push({ type: 'strike' })", "  if (run.strike ?? run.effective?.strike) marks.push({ type: 'strike' })")
        .replace('    run.italic === false ||', '    run.italic === false ||\n    run.underline === false || run.strike === false ||')
        .replace('        italicOff: run.italic === false || null,', '        italicOff: run.italic === false || null,\n        underlineOff: run.underline === false || null, strikeOff: run.strike === false || null, miniOfficeEffectXml: null,')
        .replace('      if (mark.attrs?.highlight)', '      if (mark.attrs?.underlineOff) run.underline = false\n      if (mark.attrs?.strikeOff) run.strike = false\n      if (mark.attrs?.highlight)');
      // 功能：文本框格式签名与存盘补丁涵盖完整段落设置，保留原始未建模属性；时间：2026-11-11；作者：lq。
      text = text.replace('    para.align ?? null,', `    ...['styleId','align','lineSpacing','lineRule','lineRawTwips','snapToGrid','indentLeft','indentRight','indentFirstLine','spaceBefore','spaceAfter','spaceBeforeAuto','spaceAfterAuto','shadingFill','shadingDisplay','shadingClear','borders','borderLines','keepNext','keepLines','widowControl','pageBreakBefore'].map(key => (para as unknown as Record<string, unknown>)[key] ?? null),
    para.tabStops?.length ? para.tabStops : null,`);
      const patchStart = text.indexOf('function textboxParasPatch('), patchEnd = text.indexOf('function textboxSizesPatch(', patchStart);
      let patchSource = text.slice(patchStart, patchEnd);
      patchSource = patchSource.replaceAll('runs: mergeRuns(p.runs),', "runs: mergeRuns(p.runs), format: { ...p, styleId: p.styleId ?? '', tabStops: p.tabStops ?? [] }, originalFormat: origParas[j],");
      // 新建 shape 的 map 同样提供段落下标，供原始格式对比使用。
      patchSource = patchSource.replace('box.paras.map((p) => ({', 'box.paras.map((p, j) => ({');
      text = text.slice(0, patchStart) + patchSource + text.slice(patchEnd);
      changes.push('下划线/删除线显式关闭作为原生字符属性保存，段落继承显示经有效 Run 读取');
      text = "import { patchParagraphTexts } from '@genoffice/docx-engine/text-patch';\n" + text;
      text = `import { patchWordCellParagraph } from '${specifier(`upstream/${path}`, 'serialization/table-paragraph.ts')}';\n` + text;
      const anchor = '  // fidelity on save comes from the outer table\'s bytes; reverse insertion keeps anchors valid';
      if (!text.includes(anchor)) throw new Error('格内段落来源入口已变化');
      text = text.replace(anchor, `  // 功能：来源只保存不可变 XML 和格式签名，不引入第二份可编辑正文；时间：2026-11-11；作者：lq。
  for (let i = 0; i < paraNodes.length; i++) {
    const xml = cell.richParas?.[i]?.sourceXml
    if (xml) paraNodes[i].attrs = { ...paraNodes[i].attrs, miniOfficeCellSource: { xml, format: nodeFormat(paraNodes[i]), styleId: paraNodes[i].attrs?.styleId ?? null, signature: cellSourceSignature(paraNodes[i]), skeleton: cellSourceSignature(paraNodes[i], true) } }
  }
` + anchor);
      text = text.replace('          richParas: cellParas.map((paragraph) => ({', '          richParas: cellParas.map((paragraph) => ({\n            ...preservedCellParagraph(paragraph),');
      // 功能：续格数据按原行和起格身份寻址，不再因宽度/底纹或前方插行插列就清空；时间：2026-11-11；作者：lq。
      text = text.replace('        heightTwips: model.rowHeightsTwips?.[rowIndex] ?? null,', '        miniOfficeSourceRow: rowIndex,\n        heightTwips: model.rowHeightsTwips?.[rowIndex] ?? null,')
        .replace('                mergeGap: true,', '                miniOfficeSourceColumn: start,\n                mergeGap: true,')
        .replace('              clipHeightTwips: cellClipTwips(model, rowIndex, cell, rowspan),', '              miniOfficeSourceColumn: start,\n              clipHeightTwips: cellClipTwips(model, rowIndex, cell, rowspan),')
        .replace('    remaining: number\n    span: number', '    remaining: number\n    sourceColumn: number | null\n    span: number')
        .replace('  if (!layoutUntouched) continuationCells.clear()', '  // 来源坐标由不可编辑身份携带，结构属性变化不丢弃续格。')
        .replace("      const original = continuationCells.get(rowIndex + ':' + column)", "      const original = continuationCells.get(String(rowNode.attrs?.miniOfficeSourceRow ?? -1) + ':' + (span.sourceColumn ?? -1))")
        .replace("        const record = continuationCells.get(rowIndex + ':' + cursor)", "        const record = continuationCells.get(String(rowNode.attrs?.miniOfficeSourceRow ?? -1) + ':' + (cellNode.attrs?.miniOfficeSourceColumn ?? -1))")
        .replace('          remaining: rowspan - 1,', '          remaining: rowspan - 1,\n          sourceColumn: (cellNode.attrs?.miniOfficeSourceColumn as number | null) ?? null,')
        .replace('    for (const [key, cell] of continuationCells) {', '    if (layoutUntouched) for (const [key, cell] of continuationCells) {');
      text += `
/** 功能：格内段落来源校验包含格式/列表及 run，文字单改保留书签原位置；时间：2026-11-11；作者：lq。 */
function cellSourceSignature(node: PmNode, withoutText = false): string {
  const runs = inlineToRuns(node.content ?? [])
  return JSON.stringify([node.type, nodeFormat(node), node.attrs?.styleId ?? null,
    node.type === 'docListItem' ? [node.attrs?.kind, node.attrs?.numId, node.attrs?.ilvl] : null,
    withoutText ? runs.map(run => ({ ...run, text: '' })) : runs])
}
/** 功能：只复用与当前格式匹配的来源；复杂修改仍由同源 rich-run 生成器负责；时间：2026-11-11；作者：lq。 */
function preservedCellParagraph(node: PmNode): { sourceXml?: string } {
  const source = node.attrs?.miniOfficeCellSource as { xml: string; format?: ParaFormat; styleId?: string | null; signature: string; skeleton: string } | null
  if (!source) return {}
  if (source.signature === cellSourceSignature(node)) return { sourceXml: source.xml }
  if (source.skeleton === cellSourceSignature(node, true)) {
    const xml = patchParagraphTexts(source.xml, inlineToRuns(node.content ?? []).map(run => run.text).join(''))
    if (xml) return { sourceXml: xml }
  }
  const sourceXml = patchWordCellParagraph(source, nodeFormat(node), inlineToRuns(node.content ?? []), (node.attrs?.styleId as string | null) ?? null,
    node.type === 'docListItem' ? { numId: String(node.attrs?.numId), ilvl: Number(node.attrs?.ilvl) || 0 } : undefined)
  return sourceXml ? { sourceXml } : {}
}
`;
      changes.push('格内来源 XML 随原生段落移动；未编辑或仅文字修改的段落保留书签等未知节点');
    }
    if (path === 'apps/docs/src/renderer/editor/marks.ts') {
      text = text.replace("  italicOff: 'boolean',", "  miniOfficeEffectXml: 'string',\n  italicOff: 'boolean',")
        .replace('      italicOff: { default: null as boolean | null },', '      miniOfficeEffectXml: { default: null as string | null, rendered: false },\n      italicOff: { default: null as boolean | null },');
      text = text.replace("  italicOff: 'boolean',", "  italicOff: 'boolean', underlineOff: 'boolean', strikeOff: 'boolean',")
        .replace('      italicOff: { default: null as boolean | null },', '      italicOff: { default: null as boolean | null },\n      underlineOff: { default: null as boolean | null }, strikeOff: { default: null as boolean | null },')
        .replace("    if (mark.attrs.italicOff) styles.push('font-style:normal')", "    if (mark.attrs.italicOff) styles.push('font-style:normal')\n    if (mark.attrs.underlineOff || mark.attrs.strikeOff) styles.push('text-decoration:none')")
        .replace("  { mark: 'italic', off: 'italicOff' },", "  { mark: 'italic', off: 'italicOff' },\n  { mark: 'underline', off: 'underlineOff' }, { mark: 'strike', off: 'strikeOff' },");
      changes.push('增加下划线/删除线关闭属性及剪贴板和标记互斥规则');
    }
    if (path === 'apps/docs/src/renderer/doc-style-css.ts') {
      text = text.replace('    if (d.underline || d.strike) {', "    if (info.type === 'character' && (d.underline || d.strike)) {");
      changes.push('段落下划线通过有效 Run 标记显示，避免祖先 CSS 阻止局部取消');
    }
    // 功能：提取纯模块翻译入口，不搬运 React 上下文或 AI 产品工作流；时间：2026-11-08；作者：lq。
    if (path === 'apps/docs/src/renderer/i18n/locale.tsx') {
      text = text.slice(text.indexOf("import { createI18n"), text.indexOf('const AI_LANG_DIRECTIVES'));
      changes.push('只保留模块级 t/getLang/setModuleLang 和原词典，移除 React Provider 与 AI 指令');
    }
    // 功能：字体度量只承接数据类型；原生字体读取通过本项目平台端口提供；时间：2026-11-08；作者：lq。
    if (path === 'packages/font-metrics/src/metrics.ts') {
      const source = ts.createSourceFile(path, text, ts.ScriptTarget.Latest, true);
      text = source.statements.find(node => ts.isInterfaceDeclaration(node) && node.name.text === 'FaceVerticalMetrics')?.getText(source);
      if (!text) throw new Error('缺少真实字体度量类型');
      changes.push('仅提取 FaceVerticalMetrics；不把 Node 文件系统读取带入浏览器');
    }
    if (path === 'apps/docs/src/renderer/font-resolve.ts') {
      text = text.replace('window.desktop?.fontMetrics', 'undefined');
      changes.push('移除 Electron 全局后门，使用现有 setFontMetricsProvider 注入真实字体能力');
    }
    if (path === 'apps/docs/src/renderer/editor/extensions.ts') {
      // 功能：文本框与页眉页脚采用相同的受控投影边界，保存和主历史即时看到内容；时间：2026-11-11；作者：lq。
      text = `import { paraBorderCss } from './para-border-css';\nimport { wordRegionHistory, nextWordRegionGroup, registerWordRegion } from '${specifier(`upstream/${path}`, 'editing/region-history.ts')}';\nimport { closeHistory as closeRegionHistory } from '@tiptap/pm/history';\n` + text;
      // 功能：点击工具栏/对话框继续操作当前区域，只有点击正文才退出对象文字模式；时间：2026-11-11；作者：lq。
      text = text.replace('      !dom.contains(target) &&', '      !dom.contains(target) && view.dom.contains(target) &&');
      const boxStart = text.indexOf('function mountTextboxEditors('), boxEnd = text.indexOf('/** drag the corner handle', boxStart);
      let boxSource = text.slice(boxStart, boxEnd);
      boxSource = boxSource.replace('  view: EditorView,', '  view: EditorView,\n  owner: Editor,');
      boxSource = boxSource.replace('  const editors: Editor[] = []', `  let syncingRegions = false, disposedRegions = false
  let liveCommit = () => {}
  let regionTransaction: Transaction | undefined
  let pasteGroup: { time: number; group: number } | undefined
  let historyGroup = nextWordRegionGroup(), lastEditTime = 0
  const unregisterRegions: Array<() => void> = []
  const editors: Editor[] = []`);
      boxSource = boxSource.replace('const paras = subEditorParas(sub)', 'const paras = subEditorParas(sub).map((para, j) => ({ ...box.paras[j], ...para }))');
      boxSource = boxSource.replace('extensions: textboxSubExtensions,', "extensions: [...textboxSubExtensions.filter(extension => extension.name !== 'undoRedo'), wordRegionHistory(view)],");
      boxSource = boxSource.replace('    editors[i] = sub', '    editors[i] = sub\n    unregisterRegions.push(registerWordRegion(owner, sub, () => String(getPos()) + ":" + i))');
      boxSource = boxSource.replace('    for (const frame of resizeFrames)', '    for (const unregister of unregisterRegions) unregister()\n    for (const frame of resizeFrames)');
      boxSource = boxSource.replace('  const commit = () => {', '  const commit = () => {\n    if (disposedRegions) return');
      boxSource = boxSource.replace('  const cleanup = () => {', '  const cleanup = () => {\n    disposedRegions = true');
      boxSource = boxSource.replace('content: textboxDocJson(box),', "content: { ...textboxDocJson(box), attrs: { miniOfficeStyles: view.state.doc.attrs.miniOfficeStyles } },");
      boxSource = boxSource.replace('        notifySubEditorState(transaction.docChanged)\n        if (transaction.docChanged) scheduleMeasure(i)', `        if (syncingRegions || transaction.getMeta('miniOfficeRegionProjection')) return
        if (transaction.docChanged) { regionTransaction = transaction; liveCommit(); scheduleMeasure(i) }
        else if (transaction.selectionSet) lastEditTime = 0
        notifySubEditorState(false)`);
      boxSource = boxSource.replace('    if (!changed) return\n    const pos', `    const styles = editors.find(sub => sub && JSON.stringify(sub.state.doc.attrs.miniOfficeStyles) !== JSON.stringify(view.state.doc.attrs.miniOfficeStyles))?.state.doc.attrs.miniOfficeStyles
    if (!changed && styles === undefined) return
    const pos`);
      boxSource = boxSource.replace(`    view.dispatch(
      view.state.tr.setNodeMarkup(pos, undefined, { ...current.attrs, textboxes: next }),
    )`, `    // 功能：一次区域内容/样式事务对应一次主历史提交，连续输入复用组合身份；时间：2026-11-11；作者：lq。
    const now = Date.now()
    const previousGroup = historyGroup
    const boundary = now - lastEditTime > 500 || !regionTransaction || regionTransaction.getMeta('uiEvent') === 'paste' || (!regionTransaction.getMeta('appendedTransaction') && regionTransaction.steps.some(step => step.toJSON().stepType !== 'replace'))
    if (boundary) historyGroup = nextWordRegionGroup()
    if (regionTransaction?.getMeta('uiEvent') === 'paste') {
      if (pasteGroup?.time === regionTransaction.time) historyGroup = pasteGroup.group
      else pasteGroup = { time: regionTransaction.time, group: historyGroup }
    }
    lastEditTime = now
    const tr = previousGroup !== historyGroup ? closeRegionHistory(view.state.tr) : view.state.tr
    tr.setMeta('composition', historyGroup)
    if (changed) tr.setNodeMarkup(pos, undefined, { ...current.attrs, textboxes: next })
    if (styles !== undefined) tr.setDocAttribute('miniOfficeStyles', styles)
    view.dispatch(tr)
    regionTransaction = undefined`);
      boxSource = boxSource.replace('  /** external model change', '  liveCommit = commit\n\n  /** external model change');
      boxSource = boxSource.replace('      if (sub && !sub.isDestroyed) sub.commands.setContent(textboxDocJson(box))', `      if (sub && !sub.isDestroyed) {
        syncingRegions = true
        try {
          const selection = sub.state.selection
          sub.commands.setContent({ ...textboxDocJson(box), attrs: { miniOfficeStyles: view.state.doc.attrs.miniOfficeStyles } }, { emitUpdate: false })
          const from = Math.min(selection.from, sub.state.doc.content.size - 1), to = Math.min(selection.to, sub.state.doc.content.size - 1)
          sub.commands.setTextSelection({ from, to })
          lastEditTime = 0
        } finally { syncingRegions = false }
      }`);
      // 主目录变化而 box 身份相同时也必须投影，不能早退。
      boxSource = boxSource.replace('    if (!boxes || boxes === knownBoxes) return', "    if (!boxes) return\n    if (boxes === knownBoxes && editors.every(sub => !sub || JSON.stringify(sub.state.doc.attrs.miniOfficeStyles) === JSON.stringify(view.state.doc.attrs.miniOfficeStyles))) return");
      text = text.slice(0, boxStart) + boxSource + text.slice(boxEnd);
      text = text.replace('mountTextboxEditors(dom, getNode, pos, editor.view)', 'mountTextboxEditors(dom, getNode, pos, editor.view, editor)');
      // 功能：文本框段落完整承接可编辑格式；JSON 属性只在模型边界转换；时间：2026-11-11；作者：lq。
      const jsonStart = text.indexOf('function textboxDocJson('), jsonEnd = text.indexOf('/**\n * Rich textbox editing:', jsonStart);
      let jsonSource = text.slice(jsonStart, jsonEnd);
      jsonSource = jsonSource.replace('      styleId: para.styleId ?? null,', `      tabStops: para.tabStops ? JSON.stringify(para.tabStops) : null,
      borderLines: para.borderLines ? JSON.stringify(para.borderLines) : null,
      keepNext: para.keepNext ?? null, keepLines: para.keepLines ?? null,
      widowControl: para.widowControl ?? null, pageBreakBefore: para.pageBreakBefore ?? null,
      styleId: para.styleId ?? null,`);
      jsonSource = jsonSource.replace("      'styleId',", "      'keepNext', 'keepLines', 'widowControl', 'pageBreakBefore',\n      'styleId',");
      jsonSource = jsonSource.replace('      if (value !== null && value !== undefined) Object.assign(para, { [key]: value })', '      Object.assign(para, { [key]: value ?? undefined })');
      jsonSource = jsonSource.replace('    return para', `    para.tabStops = attrs.tabStops ? JSON.parse(String(attrs.tabStops)) : []
    para.borderLines = attrs.borderLines ? JSON.parse(String(attrs.borderLines)) : undefined
    return para`);
      text = text.slice(0, jsonStart) + jsonSource + text.slice(jsonEnd);
          // other attribute change (resize, table commit, ...): recreate DOM`);
      changes.push('文本框编辑/样式实时提交到主文档并共享历史，撤销投影保留选区');

      // 功能：区域列表支持空项退出和 Tab 层级，保持单段落原生标记；时间：2026-11-11；作者：lq。
      const paragraphStart = text.indexOf('export const TextboxParagraph ='), paragraphEnd = text.indexOf('export function textboxOrderedMarkerTexts', paragraphStart);
      let paragraphSource = text.slice(paragraphStart, paragraphEnd);
      paragraphSource = paragraphSource.replace('      styleId: { default:', `      tabStops: { default: null as string | null }, borderLines: { default: null as string | null },
      keepNext: { default: null as boolean | null }, keepLines: { default: null as boolean | null }, widowControl: { default: null as boolean | null },
      pageBreakBefore: { default: null as boolean | null },
      styleId: { default:`);
      paragraphSource = paragraphSource.replace('      ...fontStyles,', `      ...fontStyles,
      ...(['t','l','b','r'] as const).filter(side => String(node.attrs.borders ?? '').includes(side)).map(side => 'border-' + ({ t: 'top', l: 'left', b: 'bottom', r: 'right' }[side]) + ':' + paraBorderCss(node.attrs.borderLines ? JSON.parse(node.attrs.borderLines)[side] : undefined)),`);
      text = text.slice(0, paragraphStart) + paragraphSource + text.slice(paragraphEnd);
      const regionKeys = `
/** 功能：精简区域 schema 的列表键盘语义，不向正文派发命令；时间：2026-11-11；作者：lq。 */
const MiniOfficeRegionListKeys = Extension.create({
  name: 'miniOfficeRegionListKeys',
  priority: 1100,
  addKeyboardShortcuts() {
    const level = (delta: number) => {
      const { state, view } = this.editor, marker = state.selection.$from.parent.attrs.listMarker
      if (!marker || view.composing) return false
      const ilvl = Math.max(0, Math.min(8, (marker.ilvl ?? 0) + delta))
      return this.editor.commands.updateAttributes('docParagraph', { listMarker: { ...marker, ilvl, indentLeft: 720 * (ilvl + 1), hanging: 360 }, indentLeft: null, indentFirstLine: null })
    }
    return {
      Tab: () => level(1), 'Shift-Tab': () => level(-1),
      Enter: () => {
        const { state, view } = this.editor
        if (!state.selection.empty || !state.selection.$from.parent.attrs.listMarker || view.composing) return false
        if (state.selection.$from.parent.content.size === 0) return this.editor.commands.updateAttributes('docParagraph', { listMarker: null })
        return this.editor.commands.splitBlock()
      },
    }
  },
})
`;
      text = text.replace('export const textboxSubExtensions = [', regionKeys + '\nexport const textboxSubExtensions = [\n  MiniOfficeRegionListKeys,');

      const port = specifier(`upstream/${path}`, 'platform/ports.ts');
      text = `import { wordPlatform } from '${port}';\n` + text.replace('window.desktop?.copyImageToClipboard?.(dataUrl, imageMetaJson(node.attrs))', 'wordPlatform().copyImageToClipboard?.(dataUrl, imageMetaJson(node.attrs))');
      changes.push('剪贴板图片写入改用本项目注入端口；保留失败时不删除图片的原逻辑');
      // 功能：延迟列宽约束同样读取目标表格所在节，不能误用第一节正文净宽；时间：2026-11-11；作者：lq。
      text = text.replace("                const raw = getComputedStyle(view.dom).getPropertyValue('--section-content-w')", `                const cell = view.nodeDOM(handle)
                const table = cell instanceof Element ? cell.closest('table.doc-table') : null
                if (!table || !view.dom.contains(table)) return
                const raw = getComputedStyle(table).getPropertyValue('--section-content-w')`);
      changes.push('延迟列宽约束采用目标表格所在分节净宽');
    }
    // 功能：Worker 由实际宿主明确创建，核心不猜测构建后 URL，也不伪装原生能力；时间：2026-11-08；作者：lq。
    if (path === 'apps/docs/src/renderer/parse-off-thread.ts') {
      const port = specifier(`upstream/${path}`, 'platform/ports.ts');
      text = `import { wordPlatform } from '${port}';\n` + text.replace("worker = new Worker(new URL('./parse-worker.ts', import.meta.url), { type: 'module' })", "worker = wordPlatform().createParseWorker?.() ?? null\n    if (!worker) return null");
      text += `\n/** 功能：关闭解析运行时并明确失败所有未完成请求；时间：2026-11-08；作者：lq。 */
export function disposeParseWorker(): void {
  worker?.terminate(); worker = null;
  const waiting = [...pending.entries()]; pending.clear();
  for (const [id, resolve] of waiting) resolve({ id, ok: false, error: 'Word parse worker disposed', bytes: new ArrayBuffer(0) });
}\n`;
      changes.push('解析 Worker 创建改为显式平台端口，保留 transferable、失败处理和实际运行计数；增加释放函数');
    }
    // 功能：同一注释渲染器支持脚注与尾注各自的编号格式，不把尾注误画成脚注；时间：2026-11-09；作者：lq。
    if (path === 'apps/docs/src/renderer/editor/note-dom.ts') {
      const signature = '  onEdit: (id: string) => void,\n): HTMLElement {';
      if (!text.includes(signature)) throw new Error('上游注释 DOM 接口已变化');
      text = text.replace(signature, "  onEdit: (id: string) => void,\n  // 功能：尾注使用自己的编号格式；时间：2026-11-09；作者：lq。\n  kind: NoteKind = 'footnote',\n): HTMLElement {")
        .replace("appendNoteRuns(row, it, 'footnote')", 'appendNoteRuns(row, it, kind)');
      changes.push('注释区增加可选 NoteKind，默认保持脚注，尾注按自身编号格式显示');
    }
    // 功能：同格式的表格纯文字修改保留原段落外壳和书签，避免原 rich-cell 重建丢失未知节点；时间：2026-11-08；作者：lq。
    if (path === 'packages/docx-engine/src/generate.ts') {
      // 功能：格内完整段落格式仅重建改动组，缩进/段距仍按属性合并以保留零值及未知属性；时间：2026-11-11；作者：lq。
      text = text.replace("['spaceBefore', 'spaceAfter', 'lineSpacing', 'lineRule', 'lineRawTwips']", "['spaceBefore', 'spaceAfter', 'spaceBeforeAuto', 'spaceAfterAuto', 'lineSpacing', 'lineRule', 'lineRawTwips']")
        .replace("['w:ind', ['indentLeft', 'indentFirstLine']]", "['w:ind', ['indentLeft', 'indentRight', 'indentFirstLine']]")
        .replace("  'w:ind': [", "  'w:ind': [\n    { attrs: ['w:right', 'w:end', 'w:rightChars', 'w:endChars'], fields: ['indentRight'] },")
        .replace("fields: ['spaceBefore']", "fields: ['spaceBefore', 'spaceBeforeAuto']")
        .replace("fields: ['spaceAfter']", "fields: ['spaceAfter', 'spaceAfterAuto']");
      const cellFormatStart = text.indexOf('export function mergeCellParaFormat('), cellFormatEnd = text.indexOf('export function mergePPrFormat(', cellFormatStart);
      let cellFormat = text.slice(cellFormatStart, cellFormatEnd);
      cellFormat = cellFormat.replace('  if (tags.size === 0) return rawPPr\n', '')
        .replace("  return /^<w:pPr(?: [^>]*)?><\\/w:pPr>$/.test(out) ? '' : out", `  // 功能：其余段落组共享同源差异判断，避免再次重建已精确合并的缩进和段距；时间：2026-11-11；作者：lq。
  const extraTags = new Set(PPR_CHILD_ORDER.filter(tag => !CELL_PPR_GROUPS.some(([handled]) => handled === tag)))
  return mergePPrFormat(out, format, original, extraTags)`);
      text = text.slice(0, cellFormatStart) + cellFormat + text.slice(cellFormatEnd);
      changes.push('布局表格单元格承接完整段落格式，未修改的段距/字符单位缩进/未知属性保留');
      // 功能：新图片自带绘图与关系命名空间，纯文字/VML 来源不必预先声明这些前缀；时间：2026-11-11；作者：lq。
      text = text.replace("'<w:drawing><wp:inline distT=", "'<w:drawing xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><wp:inline distT=");
      changes.push('新插入图片声明所需命名空间，兼容没有绘图前缀的来源');
      // 功能：文本框格式回写复用同源 pPr 增量合并，未修改组与扩展保留原字节；时间：2026-11-11；作者：lq。
      text = text.replace('export interface TextboxParaPatch {', 'export interface TextboxParaPatch {\n  format?: ParaFormat\n  originalFormat?: ParaFormat');
      text = text.replace('    if (patch.numPr !== undefined) pPr = pPrWithNumPr', '    if (patch.format) pPr = mergePPrFormat(pPr, patch.format, patch.originalFormat)\n    if (patch.numPr !== undefined) pPr = pPrWithNumPr');

      // 功能：区域与布局表格通过 ParaFormat 管理样式引用，未提供身份时保留来源引用；时间：2026-11-11；作者：lq。
      text = text.replace('  if (!format) return []\n  const out: PPrChild[] = []', `  if (!format) return []
  const out: PPrChild[] = []
  if (format.styleId) out.push({ name: 'w:pStyle', xml: '<w:pStyle w:val="' + escapeXmlAttr(format.styleId) + '"/>' })`);
      text = text.replace('  if (format.styleId) out.push', `  if (format.listMarker?.numId) out.push({ name: 'w:numPr', xml: '<w:numPr><w:ilvl w:val="' + (format.listMarker.ilvl ?? 0) + '"/><w:numId w:val="' + escapeXmlAttr(format.listMarker.numId) + '"/></w:numPr>' })
  if (format.styleId) out.push`);
      text = text.replace('  const managedTags = new Set(FORMAT_MANAGED_TAGS)', "  const managedTags = new Set(FORMAT_MANAGED_TAGS)\n  if (format?.styleId !== undefined) managedTags.add('w:pStyle')\n  if (format?.listMarker !== undefined || original?.listMarker !== undefined) managedTags.add('w:numPr')");

      text = text.replace("return rawBool(rawOf('w:strike')) === !!run.strike", "return run.strike === false ? !!rawOf('w:strike') && !rawBool(rawOf('w:strike')) : rawBool(rawOf('w:strike')) === !!run.strike")
        .replace("return (val !== undefined && val !== 'none') === !!run.underline", "return run.underline === false ? val === 'none' : (val !== undefined && val !== 'none') === !!run.underline");
      text = text.replace("  if (run.strike) out.push({ name: 'w:strike', xml: '<w:strike/>' })", "  if (run.strike !== undefined) out.push({ name: 'w:strike', xml: onOffXml('w:strike', run.strike) })")
        .replace("  if (run.underline) out.push({ name: 'w:u', xml: '<w:u w:val=\"single\"/>' })", "  if (run.underline !== undefined) out.push({ name: 'w:u', xml: `<w:u w:val=\"${run.underline ? 'single' : 'none'}\"/>` })");
      text = text.replace('return `<w:${side} w:val="single" w:sz="${sz}"', 'return `<w:${side} w:val="${escapeXmlAttr(declared?.val ?? \'single\')}" w:sz="${sz}"');
      text = text.replace("      const line: NonNullable<ParaFormat['borderLines']>[typeof ch] = {}", "      const line: NonNullable<ParaFormat['borderLines']>[typeof ch] = {}\n      if (val && val !== 'single') line.val = val");
      text = text.replace('if (line.color !== undefined || line.szPt !== undefined || line.spacePt !== undefined)', 'if (line.val !== undefined || line.color !== undefined || line.szPt !== undefined || line.spacePt !== undefined)');
      text = text.replace('        lines?.[ch]?.color ?? null,', '        lines?.[ch]?.val ?? null,\n        lines?.[ch]?.color ?? null,');
      changes.push('段落边框类型参与生成及来源一致性比较，未改双线保留原 XML');
      text = "import { patchParagraphTexts } from './text-patch';\n" + text;
      const anchor = '      const emittedMedia = new Set(mediaRIds(emitted))';
      const fastPath = String.raw`      // 功能：先比较 run 结构和格式，仅纯文字改动走原 XML 补丁；时间：2026-11-08；作者：lq。
      if (aligned && entry.xml === undefined && pPr === pPrOf(template)) {
        const signature = (xml: string) => xmlSegments(xml, 'w:r', 0, xml.length)
          .map(seg => xml.slice(seg.start, seg.end).replace(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/g, '<w:t/>')).join('')
        const newText = entry.runs.map(run => run.text ?? '').join('')
        if (!/[\t\n\r\f\v\u001e\u2011\u00ad]/.test(newText) && signature(template) === signature(emitted)) {
          const patched = patchParagraphTexts(template, newText)
          if (patched !== null) return patched
        }
      }
`;
      if (!text.includes(anchor)) throw new Error('上游单元格回写位置已变化，不能自动应用本地补丁');
      text = text.replace(anchor, fastPath + anchor);
      text = text.replace('  const paraXmls = paragraphs.map((paragraph) => {', `  const paraXmls = paragraphs.map((paragraph) => {
    // 功能：来源校验已由转换器完成，表格属性改变不重建未目标段落；时间：2026-11-11；作者：lq。
    if ('sourceXml' in paragraph && paragraph.sourceXml) return paragraph.sourceXml`);
      changes.push('增加同格式纯文字的原段落 XML 补丁，保留书签和未知段落属性；复杂结构仍走原有生成逻辑');
    }
    const info = ts.preProcessFile(text, true, true);
    const edits = [];
    for (const imported of info.importedFiles) {
      const value = imported.fileName;
      let resolved;
      if (value.startsWith('.')) {
        if (value.includes('platform/ports') || value.includes('serialization/table-paragraph') || value.includes('serialization/run-effects') || value.includes('serialization/paragraph-structure') || value.includes('serialization/header-footer-boxes') || value.includes('editing/region-history')) continue;
        resolved = resolveModule(posix.normalize(posix.join(posix.dirname(path), value)));
        if (resolved === 'apps/docs/src/shared/ipc.ts') resolved = 'packages/agent-core/src/types.ts';
      } else {
        resolved = aliases[value];
        if (!resolved && value.startsWith('@genoffice/docx-engine/')) resolved = resolveModule('packages/docx-engine/src/' + value.slice('@genoffice/docx-engine/'.length));
        if (!resolved) {
          if (value.startsWith('@genoffice/') || value.startsWith('node:')) throw new Error(`未分离的跨应用/原生依赖：${path}: ${value}`);
          externals.add(value);
          continue;
        }
      }
      if (resolved !== shared) queue.push(resolved);
      const replacement = specifier(path, resolved);
      const start = imported.pos;
      // TypeScript 的模块区间包含引号，用原始字符串定位避免破坏其他标识。
      const quoteStart = text.lastIndexOf(value, imported.end);
      if (quoteStart < start - 1) throw new Error(`无法安全替换导入：${path}: ${value}`);
      edits.push({ start: quoteStart, end: quoteStart + value.length, replacement });
      if (value.startsWith('@genoffice/') || value.includes('shared/ipc')) changes.push(`导入 ${value} → ${replacement}`);
    }
    for (const edit of edits.sort((a, b) => b.start - a.start)) text = text.slice(0, edit.start) + edit.replacement + text.slice(edit.end);
    // 功能：M2 新增来源文件标注实际集成日期，既有 M1 来源保持原始哈希；时间：2026-11-09；作者：lq。
    const m2Source = ['apps/docs/src/renderer/editor/note-dom.ts', 'apps/docs/src/renderer/hf-heights.ts', 'apps/docs/src/renderer/hf-phantom.ts'].includes(path);
    const m3Source = /editor\/(?:table-ops|table-properties|table-borders|insert-picture|new-table)\.ts$/.test(path);
    bytes = Buffer.from((m3Source ? banner.replace('2026-11-08', '2026-11-11') : m2Source ? banner.replace('2026-11-08', '2026-11-09') : banner) + text);
    changes.unshift('增加本地集成功能、日期与作者注释；保留原作者及源码注释');
  }
  save(path, bytes, sourcePath, changes);
}

// 功能：提取文档绘制规则并限制在核心容器内，不加载原产品 Ribbon/App 样式；时间：2026-11-08；作者：lq。
const cssPath = 'apps/docs/src/renderer/styles.css';
const css = postcss.parse(original(cssPath).toString('utf8'));
// 功能：图片八向尺寸手柄属于正文对象，基础定位规则必须随显示规则一起保留；时间：2026-11-11；作者：lq。
// 功能：保留混合纸张的逐页纸面，避免不同宽度页面丢失底色；时间：2026-11-11；作者：lq。
const selectors = /\.(?:doc-|img-resize-handle|ProseMirror|page-gap|page-float|page-hf|page-sheets?|editor-scroll|hf-|footnote|endnote|note-|float-|col-|textbox)/;
/** 功能：递归裁剪非文档规则，保留相关媒体和 keyframes，避免外壳 CSS 污染；时间：2026-11-08；作者：lq。 */
function prune(container) {
  for (const node of [...(container.nodes ?? [])]) {
    if (node.type === 'comment') { node.remove(); continue; }
    if (node.type === 'rule') {
      if (!selectors.test(node.selector) && !/^:root/.test(node.selector) && !node.parent?.name?.includes('keyframes')) { node.remove(); continue; }
      if (node.parent?.name?.includes('keyframes')) continue;
      node.selector = postcss.list.comma(node.selector).map(selector => {
        if (/^:root/.test(selector)) return selector.replace(':root', '.mo-word-core');
        if (/^html(?:\[[^\]]+\])?\s/.test(selector)) return selector.replace(/^(html(?:\[[^\]]+\])?)\s+/, '$1 .mo-word-core ');
        return `.mo-word-core ${selector}`;
      }).join(', ');
    } else if (node.type === 'atrule') {
      if (node.nodes) { prune(node); if (!node.nodes.length) node.remove(); }
      else node.remove();
    }
  }
}
prune(css);
save('../styles/document.css', banner + css.toString(), cssPath, ['仅保留文档排版规则并加 .mo-word-core 范围；排除产品界面样式']);

// 功能：保留文档字体及许可证，重写共享字体路径，不引入上游 UI 组件；时间：2026-11-08；作者：lq。
for (const directory of ['apps/docs/src/renderer/fonts', 'packages/ui/src/fonts']) {
  for (const name of readdirSync(join(origin, directory))) {
    const path = `${directory}/${name}`;
    const input = original(path);
    const content = name.endsWith('.css') ? banner + input.toString('utf8').replaceAll('@genoffice/ui/fonts/', '../../../../../packages/ui/src/fonts/') : input;
    save(path, content, path, name.endsWith('.css') ? ['字体 CSS 路径改为包内资源'] : []);
  }
}
mkdirSync(join(root, 'packages/word-core/src/styles'), { recursive: true });
writeFileSync(join(root, 'packages/word-core/src/styles/fonts.css'), banner + "@import '../upstream/apps/docs/src/renderer/fonts/fonts.css';\n");
for (const name of ['LICENSE', 'NOTICE']) {
  mkdirSync(join(root, 'packages/word-core/licenses'), { recursive: true });
  writeFileSync(join(root, 'packages/word-core/licenses', `GenOffice-${name}.txt`), original(name));
}
writeFileSync(join(root, 'packages/word-core/SOURCE-MANIFEST.json'), JSON.stringify({ _comment: '功能：记录定制 Word 核心逐文件来源与集成修改；时间：2026-11-08；作者：lq。', schemaVersion: 1, upstreamCommit: baseline.commit, extractedFiles: mappings.sort((a, b) => a.localPath.localeCompare(b.localPath, 'en')), externalModules: [...externals].sort() }, null, 2) + '\n');
console.log(JSON.stringify({ files: mappings.length, externalModules: [...externals].sort() }));
