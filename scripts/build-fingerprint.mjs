/** 功能：对工作区实际源码生成指纹，包含未提交文件；打包前核对 dist 与当前源码一致；时间：2026-11-14；作者：lq。 */
import { createHash } from 'node:crypto';
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

/** 功能：按稳定顺序读取源码文件，跳过构建目录与依赖；时间：2026-11-14；作者：lq。 */
function sources(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap(entry => {
    const path = resolve(directory, entry.name);
    if (entry.isDirectory()) return ['node_modules', 'dist', 'target', '.git'].includes(entry.name) ? [] : sources(path);
    return /\.(?:tsx?|js|css|json|rs|toml|mjs|html)$/.test(entry.name) ? [path] : [];
  });
}

/** 功能：源码路径和内容共同参与 SHA256，不能仅用 Git 提交号替代脏工作区；时间：2026-11-14；作者：lq。 */
export function sourceFingerprint() {
  const files = [
    ...sources(resolve(root, 'packages')),
    ...sources(resolve(root, 'apps/desktop/src')),
    ...sources(resolve(root, 'apps/desktop/src-tauri/src')),
    // 功能：旧 DOC 转换源码和实际生成的 JS 都纳入构建溯源；时间：2026-11-14；作者：lq。
    ...sources(resolve(root, 'third_party/genoffice')),
    ...['package.json', 'pnpm-lock.yaml', 'Cargo.toml', 'Cargo.lock', 'apps/desktop/vite.config.ts', 'apps/desktop/index.html', 'apps/desktop/src-tauri/build.rs', 'apps/desktop/src-tauri/Cargo.toml', 'apps/desktop/src-tauri/tauri.conf.json', 'scripts/build-fingerprint.mjs', 'scripts/package-macos-local.sh'].map(path => resolve(root, path)),
  ].sort();
  const hash = createHash('sha256');
  for (const path of files) hash.update(relative(root, path)).update('\0').update(readFileSync(path)).update('\0');
  return hash.digest('hex');
}

// 功能：复用打包时校验，发现源码已变更就拒绝复用旧 dist；时间：2026-11-14；作者：lq。
if (process.argv[2] === '--check') {
  const info = JSON.parse(readFileSync(resolve(root, 'apps/desktop/dist/build-info.json'), 'utf8'));
  if (info.sourceSha256 !== sourceFingerprint()) throw new Error('当前源码与前端构建不一致，请重新完整构建');
  process.stdout.write(info.buildId);
}
