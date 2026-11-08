/** 功能：Mini Office 定制 Word 核心的公开入口，不包含产品 UI；时间：2026-11-08；作者：lq。 */
export * from './upstream/packages/docx-engine/src/index';
export { editorExtensions, blocksToPmDoc, pmDocOptions, pmDocToSavePlan } from './editing/index';
export { configureWordPlatform, wordPlatformCapabilities, type WordPlatform, type FaceVerticalMetrics } from './platform/ports';
export { WordCoreSession, type WordSessionState, type WordSessionOptions, type PreparedWordSave } from './session/word-session';
export { serializeWordDocument } from './serialization/document';
export { parseDocxOffThread, parseStats, disposeParseWorker } from './upstream/apps/docs/src/renderer/parse-off-thread';
