/** 功能：从固定的 GenOffice 源码构建浏览器兼容的 DOC 转换与原始 XML 工具；时间：2026-11-03；作者：lq。 */
import { createRequire } from 'node:module';
import { resolve, dirname, join, isAbsolute } from 'node:path';
import { readFile, writeFile, mkdir, readdir } from 'node:fs/promises';
const require = createRequire(import.meta.url);
const { build } = createRequire(require.resolve('vite/package.json'))('esbuild');
const banner = '/** @license GenOffice Copyright 2026 Mainfunc, Inc. Apache-2.0. 功能：DOC/OOXML 兼容模块；时间：2026-11-03；集成作者：lq；上游版本 e1f50e90，许可及源码见 third_party/genoffice。 */';
await build({
  entryPoints: {
    'doc-converter': 'third_party/genoffice/doc2docx/index.ts',
    'xml-scan': 'third_party/genoffice/docx-engine/scan.ts',
    'xml-text-patch': 'third_party/genoffice/docx-engine/text-patch.ts',
  },
  outdir: 'packages/platform-adapter/src/vendor/genoffice',
  bundle: true,
  platform: 'browser',
  format: 'esm',
  target: 'es2022',
  minify: true,
  banner: { js: banner },
  alias: {
    '@genoffice/docx-engine': resolve('third_party/genoffice/engine-entry.ts'),
    'node:zlib': resolve('third_party/genoffice/browser-inflate.ts'),
  },
  external: ['jszip', 'cfb', 'fflate', 'fast-xml-parser'],
  legalComments: 'inline',
});

// 功能：将复用源码的许可和依赖许可随 Web/Tauri 静态资源分发；时间：2026-11-03；作者：lq。
const notices = resolve('apps/desktop/public/licenses');
await mkdir(notices, { recursive: true });
for (const name of ['LICENSE', 'NOTICE']) {
  await writeFile(resolve(notices, `GenOffice-${name}.txt`), await readFile(resolve('third_party/genoffice', name)));
}

const seen = new Set();
let dependencyText = 'Mini Office DOC compatibility dependencies\nIntegration: lq, 2026-11-03\n';
/** 功能：递归收集新增转换依赖的已安装许可证，跳过 Node 内置模块；时间：2026-11-03；作者：lq。 */
async function collectLicense(name, from) {
  let entry;
  try { entry = from.resolve(name); } catch { return; }
  if (!isAbsolute(entry)) return;
  let root = dirname(entry);
  let pkg;
  while (!pkg) {
    try { pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8')); }
    catch { if (root === dirname(root)) return; root = dirname(root); }
  }
  const id = `${pkg.name}@${pkg.version}`;
  if (seen.has(id)) return;
  seen.add(id);
  dependencyText += `\n\n=== ${id} (${pkg.license}) ===\n`;
  const license = (await readdir(root)).find((file) => /^licen[sc]e(?:\.(?:txt|md))?$/i.test(file));
  if (license) dependencyText += await readFile(join(root, license), 'utf8');
  const next = createRequire(join(root, 'package.json'));
  for (const child of Object.keys(pkg.dependencies ?? {})) await collectLicense(child, next);
}
const platformRequire = createRequire(resolve('packages/platform-adapter/package.json'));
for (const name of ['cfb', 'fflate', 'fast-xml-parser', 'jszip']) await collectLicense(name, platformRequire);
await writeFile(resolve(notices, 'DOC-dependencies.txt'), dependencyText);
