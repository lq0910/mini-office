import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'node:path';
// 功能：每次构建记录实际源码指纹，并在应用内显示独立构建号；时间：2026-11-14；作者：lq。
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { sourceFingerprint } from '../../scripts/build-fingerprint.mjs';

const sourceSha256 = sourceFingerprint();
const builtAt = new Date().toISOString();
const localTime = new Date(Date.now() + 8 * 60 * 60 * 1000).toISOString().replace(/[-:]/g, '');
const buildId = `${localTime.slice(0, 8)}-${localTime.slice(9, 15)}-${sourceSha256.slice(0, 8)}`;

const pkg = (name: string) => resolve(__dirname, `../../packages/${name}/src/index.ts`);

export default defineConfig({
  define: { 'import.meta.env.VITE_BUILD_ID': JSON.stringify(buildId) },
  plugins: [react(), {
    name: 'mini-office-build-provenance',
    // 功能：随前端输出构建信息与资源哈希，供原生打包核对；时间：2026-11-14；作者：lq。
    generateBundle(_options, bundle) {
      const assets = Object.fromEntries(Object.entries(bundle).map(([path, item]) => [path, createHash('sha256').update(item.type === 'chunk' ? item.code : item.source).digest('hex')]));
      this.emitFile({ type: 'asset', fileName: 'build-info.json', source: JSON.stringify({ _comment: '功能：构建溯源；时间：2026-11-14；作者：lq。', buildId, builtAt, sourceSha256, assets }, null, 2) });
    },
    // 功能：Vite 后续插件会改写入口的预加载路径，必须按最终落盘字节更新清单，避免提前哈希误判旧代码；时间：2026-11-14；作者：lq。
    writeBundle(options, bundle) {
      const output = resolve(__dirname, options.dir ?? 'dist');
      const assets = Object.fromEntries(Object.keys(bundle).filter(path => path !== 'build-info.json').map(path => [path, createHash('sha256').update(readFileSync(resolve(output, path))).digest('hex')]));
      writeFileSync(resolve(output, 'build-info.json'), JSON.stringify({ _comment: '功能：最终落盘资源溯源；时间：2026-11-14；作者：lq。', buildId, builtAt, sourceSha256, assets }, null, 2));
    },
  }],
  resolve: {
    alias: {
      '@mini-office/design-system': pkg('design-system'),
      '@mini-office/shared': pkg('shared'),
      '@mini-office/ui': pkg('ui'),
      '@mini-office/platform-adapter': pkg('platform-adapter'),
      '@mini-office/document-model': pkg('document-model'),
      '@mini-office/ooxml': pkg('ooxml'),
      '@mini-office/editor-shell': pkg('editor-shell'),
      '@mini-office/editor-word': pkg('editor-word'),
      '@mini-office/editor-sheet': pkg('editor-sheet'),
      '@mini-office/editor-slides': pkg('editor-slides'),
      '@mini-office/export-service': pkg('export-service'),
      '@app': resolve(__dirname, 'src'),
    },
  },
  clearScreen: false,
  server: {
    port: 1420,
    strictPort: true,
    host: '127.0.0.1',
    watch: { ignored: ['**/src-tauri/**'] },
  },
  envPrefix: ['VITE_', 'TAURI_ENV_'],
  build: {
    target: 'es2022',
    sourcemap: false,
    chunkSizeWarningLimit: 1600,
  },
});
