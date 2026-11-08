/** 功能：定义 Word 所需的具体平台能力，移除 Electron 全局依赖；时间：2026-11-08；作者：lq。 */
import type { FaceVerticalMetrics } from '../upstream/packages/font-metrics/src/metrics';
import { setFontMetricsProvider } from '../upstream/apps/docs/src/renderer/font-resolve';

/** 功能：只描述实际需要的平台能力，不提供伪成功的空实现；时间：2026-11-08；作者：lq。 */
export interface WordPlatform {
  fontMetrics?: (family: string) => Promise<FaceVerticalMetrics | null>;
  copyImageToClipboard?: (dataUrl: string, metadata: string) => Promise<boolean>;
  createParseWorker?: () => Worker;
}
let platform: Readonly<WordPlatform> = Object.freeze({});

/** 功能：安装应用提供的真实能力；缺少原生字体索引时使用上游内嵌/度量数据及浏览器测量；时间：2026-11-08；作者：lq。 */
export function configureWordPlatform(value: WordPlatform): void {
  platform = Object.freeze({ ...value });
  setFontMetricsProvider(value.fontMetrics ?? null);
}
/** 功能：读取明确可用的平台能力，调用方继续保留原来的失败回退语义；时间：2026-11-08；作者：lq。 */
export function wordPlatform(): Readonly<WordPlatform> { return platform; }
/** 功能：给产品报告实际能力状态，不把未接通的原生操作标记为可用；时间：2026-11-08；作者：lq。 */
export function wordPlatformCapabilities() {
  return { fontMetrics: !!platform.fontMetrics, copyImageToClipboard: !!platform.copyImageToClipboard, parseWorker: !!platform.createParseWorker };
}
export type { FaceVerticalMetrics };
