/** 功能：按需加载 Word 时安装真实桌面平台能力；时间：2026-11-11；作者：lq。 */
import { isMacOS, isTauriRuntime, coreWordFontMetrics, copyCoreWordImage } from '@mini-office/platform-adapter';
import { configureWordPlatform } from '../../word-core/src/platform/ports';

// 功能：Web 保留浏览器复制，macOS 使用原生位图桥；解析采用独立模块 worker；时间：2026-11-11；作者：lq。
configureWordPlatform({
  ...(isTauriRuntime() ? { fontMetrics: coreWordFontMetrics, ...(isMacOS() ? { copyImageToClipboard: copyCoreWordImage } : {}) } : {}),
  createParseWorker: typeof Worker === 'undefined' ? undefined : () => new Worker(new URL('../../word-core/src/upstream/apps/docs/src/renderer/parse-worker.ts', import.meta.url), { type: 'module' }),
});
