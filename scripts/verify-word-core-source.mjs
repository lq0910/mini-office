/** 功能：核验 M1 提取来源、真实依赖版本和独立浏览器构建边界；时间：2026-11-08；作者：lq。 */
import { createHash } from 'node:crypto';
import { readFileSync, existsSync } from 'node:fs';
import { resolve, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const core = join(root, 'packages/word-core');
const manifest = JSON.parse(readFileSync(join(core, 'SOURCE-MANIFEST.json'), 'utf8'));
const baseline = JSON.parse(readFileSync(join(root, 'docs/word-core-m0/upstream.json'), 'utf8'));
const known = new Map(baseline.files.map(file => [file.path, file.sha256]));
const errors = [];
/** 功能：以原字节检查本地来源映射，拒绝未经记录的修改；时间：2026-11-08；作者：lq。 */
const sha = path => createHash('sha256').update(readFileSync(path)).digest('hex');
if (manifest.upstreamCommit !== baseline.commit) errors.push('上游提交与 M0 起点不一致');
for (const file of manifest.extractedFiles) {
  if (!file.localPath.startsWith('packages/word-core/src/')) { errors.push(`来源映射越界：${file.localPath}`); continue; }
  const path = join(root, file.localPath);
  if (!existsSync(path) || sha(path) !== file.localSha256) errors.push(`本地源码变化：${file.localPath}`);
  if (file.upstreamPath && known.has(file.upstreamPath) && known.get(file.upstreamPath) !== file.upstreamSha256) errors.push(`上游指纹不一致：${file.upstreamPath}`);
}
const pkg = JSON.parse(readFileSync(join(core, 'package.json'), 'utf8'));
const versions = {};
for (const [name, version] of Object.entries(pkg.dependencies)) {
  const installed = JSON.parse(readFileSync(join(core, 'node_modules', name, 'package.json'), 'utf8'));
  versions[name] = installed.version;
  if (version !== installed.version) errors.push(`实际依赖版本漂移：${name}`);
}
const supplements = JSON.parse(readFileSync(join(core, 'licenses/SOURCES.json'), 'utf8'));
for (const item of supplements.supplements) if (sha(join(core, 'licenses', item.file)) !== item.sha256) errors.push(`补充许可变化：${item.file}`);
const metadata = JSON.parse(readFileSync(join(core, 'dist/metafile.json'), 'utf8'));
const inputs = Object.keys(metadata.inputs);
const forbidden = inputs.filter(path => /(?:output\/|private\/tmp|node_modules\/(?:react|react-dom|electron)\/|editor-word\/|document-model\/|ai-provider|ai-search)/.test(path));
errors.push(...forbidden.map(path => `构建边界违规：${path}`));
if (!Object.keys(metadata.outputs).some(path => path.endsWith('parse-worker.js'))) errors.push('没有实际构建解析 Worker');
const result = { status: errors.length ? 'failed' : 'passed', upstreamCommit: manifest.upstreamCommit, sourceFiles: manifest.extractedFiles.length, modules: inputs.length, dependencyVersions: versions, productEditorSwitched: false, reactUpgradeRequired: false, errors };
console.log(JSON.stringify(result, null, 2));
if (errors.length) process.exitCode = 1;
