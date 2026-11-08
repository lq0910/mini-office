/** 功能：独立构建 Word 核心并核对浏览器依赖、字体资源和许可；时间：2026-11-08；作者：lq。 */
import { createRequire } from 'node:module';
import { resolve, dirname, join, relative } from 'node:path';
import { readFile, writeFile, mkdir, readdir, copyFile, realpath } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { existsSync } from 'node:fs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const workspaceRequire = createRequire(join(root, 'package.json'));
const { build } = createRequire(workspaceRequire.resolve('vite/package.json'))('esbuild');
const packageRoot = join(root, 'packages/word-core');
const output = join(packageRoot, 'dist');
const result = await build({
  absWorkingDir: packageRoot,
  entryPoints: { index: 'src/index.ts', editing: 'src/editing/index.ts', layout: 'src/layout/index.ts', document: 'src/upstream/packages/docx-engine/src/index.ts', 'parse-worker': 'src/upstream/apps/docs/src/renderer/parse-worker.ts', styles: 'src/styles/document.css', fonts: 'src/styles/fonts.css' },
  outdir: output, bundle: true, splitting: true, format: 'esm', platform: 'browser', target: 'es2022',
  minify: true, metafile: true, legalComments: 'inline',
  assetNames: 'assets/[name]-[hash]', chunkNames: 'chunks/[name]-[hash]',
  loader: { '.ttf': 'file', '.otf': 'file', '.woff': 'file', '.woff2': 'file' },
  plugins: [{
    name: 'word-browser-boundary',
    /** 功能：原生及旧应用依赖未分离时立即中止，不让构建器生成空替身；时间：2026-11-08；作者：lq。 */
    setup(builder) {
      builder.onResolve({ filter: /^(?:node:|electron(?:\/|$)|@genoffice\/)/ }, args => ({ errors: [{ text: `Word 核心仍依赖原应用或 Node 接口：${args.path} (${args.importer})` }] }));
    },
  }],
});

// 功能：发布资源附带实际引入的字体及上游许可，避免只保留根 LICENSE；时间：2026-11-08；作者：lq。
await mkdir(join(output, 'licenses'), { recursive: true });
for (const file of await readdir(join(packageRoot, 'licenses'))) await copyFile(join(packageRoot, 'licenses', file), join(output, 'licenses', file));
for (const directory of ['apps/docs/src/renderer/fonts', 'packages/ui/src/fonts']) {
  const base = join(packageRoot, 'src/upstream', directory);
  for (const file of await readdir(base)) {
    if (/^(?:LICENSE|NOTICE|README)/i.test(file)) await copyFile(join(base, file), join(output, 'licenses', directory.startsWith('apps') ? `docs-fonts-${file}` : `shared-fonts-${file}`));
  }
}
await copyFile(join(packageRoot, 'src/upstream/packages/docx-engine/src/vendor/emf-converter/LICENSE'), join(output, 'licenses/emf-converter-LICENSE.txt'));

const coreRequire = createRequire(join(packageRoot, 'package.json'));
const seen = new Set();
let dependencyNotices = 'Mini Office Word core dependencies\nIntegration: lq, 2026-11-08\n';
/** 功能：沿已安装依赖树收集许可，不从网络补造缺失的归属；时间：2026-11-08；作者：lq。 */
async function collect(name, from) {
  let entry;
  try { entry = from.resolve(`${name}/package.json`); } catch {
    // 功能：部分 ProseMirror 包只导出子路径，从 Node 的真实搜索目录读取包元数据；时间：2026-11-08；作者：lq。
    entry = from.resolve.paths(name)?.map(base => join(base, name, 'package.json')).find(path => existsSync(path));
    if (!entry) entry = from.resolve(name);
  }
  let directory = dirname(await realpath(entry));
  while (directory !== dirname(directory)) {
    try { const content = JSON.parse(await readFile(join(directory, 'package.json'), 'utf8')); if (content.name === name) break; } catch {}
    directory = dirname(directory);
  }
  const content = JSON.parse(await readFile(join(directory, 'package.json'), 'utf8'));
  const id = `${content.name}@${content.version}`;
  if (seen.has(id)) return;
  seen.add(id);
  const names = await readdir(directory);
  const licenses = names.filter(file => /^licen[sc]e(?:[._-].*)?$/i.test(file));
  dependencyNotices += `\n\n=== ${id} (${content.license ?? '见包内许可'}) ===\n`;
  if (licenses.length) for (const license of licenses) dependencyNotices += await readFile(join(directory, license), 'utf8');
  else if (id === '@nodable/entities@3.1.0') dependencyNotices += await readFile(join(packageRoot, 'licenses/nodable-entities-3.1.0-LICENSE.txt'), 'utf8');
  else {
    // 功能：部分旧依赖把完整许可放在 README 中，保留实际授权文本而非只记录 SPDX 名称；时间：2026-11-08；作者：lq。
    const readme = names.find(file => /^readme(?:\.[a-z]+)?$/i.test(file));
    const text = readme ? await readFile(join(directory, readme), 'utf8') : '';
    if (!text.includes('Permission is hereby granted')) throw new Error(`依赖缺少可分发许可文件：${id}`);
    dependencyNotices += text;
  }
  const next = createRequire(join(directory, 'package.json'));
  for (const child of Object.keys(content.dependencies ?? {})) await collect(child, next);
}
for (const name of Object.keys(JSON.parse(await readFile(join(packageRoot, 'package.json'), 'utf8')).dependencies)) await collect(name, coreRequire);
await writeFile(join(output, 'licenses/dependencies.txt'), dependencyNotices);

const inputs = Object.keys(result.metafile.inputs);
if (inputs.some(path => /(?:node_modules\/(?:react|react-dom|electron)\/|apps\/shell\/|ai-provider|ai-search|node:)/.test(path))) throw new Error('构建图包含不允许的应用/框架依赖');
const outputs = Object.entries(result.metafile.outputs).map(([path, data]) => ({ path: relative(packageRoot, resolve(packageRoot, path)), bytes: data.bytes }));
const report = { _comment: '功能：记录 M1 实际浏览器构建，不代表产品已切换内核；时间：2026-11-08；作者：lq。', status: 'passed', modules: inputs.length, reactIncluded: false, electronIncluded: false, nodeBuiltinsIncluded: false, outputs, totalBytes: outputs.reduce((sum, file) => sum + file.bytes, 0), dependenciesWithLicenses: seen.size };
await writeFile(join(output, 'metafile.json'), JSON.stringify(result.metafile, null, 2));
await writeFile(join(output, 'build-report.json'), JSON.stringify(report, null, 2));
console.log(JSON.stringify({ status: report.status, modules: report.modules, outputs: outputs.length, totalBytes: report.totalBytes, dependenciesWithLicenses: seen.size }));
