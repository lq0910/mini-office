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
      const port = specifier(`upstream/${path}`, 'platform/ports.ts');
      text = `import { wordPlatform } from '${port}';\n` + text.replace('window.desktop?.copyImageToClipboard?.(dataUrl, imageMetaJson(node.attrs))', 'wordPlatform().copyImageToClipboard?.(dataUrl, imageMetaJson(node.attrs))');
      changes.push('剪贴板图片写入改用本项目注入端口；保留失败时不删除图片的原逻辑');
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
      changes.push('增加同格式纯文字的原段落 XML 补丁，保留书签和未知段落属性；复杂结构仍走原有生成逻辑');
    }
    const info = ts.preProcessFile(text, true, true);
    const edits = [];
    for (const imported of info.importedFiles) {
      const value = imported.fileName;
      let resolved;
      if (value.startsWith('.')) {
        if (value.includes('platform/ports')) continue;
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
    bytes = Buffer.from((m2Source ? banner.replace('2026-11-08', '2026-11-09') : banner) + text);
    changes.unshift('增加本地集成功能、日期与作者注释；保留原作者及源码注释');
  }
  save(path, bytes, sourcePath, changes);
}

// 功能：提取文档绘制规则并限制在核心容器内，不加载原产品 Ribbon/App 样式；时间：2026-11-08；作者：lq。
const cssPath = 'apps/docs/src/renderer/styles.css';
const css = postcss.parse(original(cssPath).toString('utf8'));
const selectors = /\.(?:doc-|ProseMirror|page-gap|page-float|editor-scroll|hf-|footnote|endnote|note-|float-|col-|textbox)/;
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
