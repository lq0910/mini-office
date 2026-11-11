/** 功能：M2 开发入口直接打开完整 GenOffice 衍生 Word 会话，杜绝旧模型保存旁路；时间：2026-11-09；作者：lq。 */
import type { CoreWordDocument, FormatHandler, FormatOpenResult } from '@mini-office/document-model';
import { WordCoreSession } from '../../word-core/src/session/word-session';
// 功能：新核心按真实包检测能力及只读限制，不用统一 L3 隐藏未开放功能；时间：2026-11-11；作者：lq。
import { detectKind, loadOoxmlPackage } from '@mini-office/ooxml';
import { ErrorCodes, OfficeError, type CapabilityReport } from '@mini-office/shared';
import { analyzeCoreWordCapabilities } from './core-capabilities';

/** 功能：能力来自新核心实际会话，保留原始未知 XML 不沿用旧解析器的功能降级判断；时间：2026-11-09；作者：lq。 */
function opened(wordCore: WordCoreSession, capability: CapabilityReport): FormatOpenResult<CoreWordDocument> {
  const readonly = !wordCore.editor.isEditable;
  if (readonly) capability.level = 1;
  return { data: { engine: 'word-core' }, wordCore, readonly,
    capability };
}

/** 功能：新建、打开均创建同源会话，实际序列化只能从工作区持有的会话发起；时间：2026-11-09；作者：lq。 */
export const coreWordFormatHandler: FormatHandler<CoreWordDocument> = {
  kind: 'docx', extensions: ['.docx', '.docm', '.dotx'],
  createBlank: () => ({ engine: 'word-core' }),
  createNew: async () => opened(await WordCoreSession.create(), { level: 3, warnings: [], missingFonts: [], scannedAt: Date.now(), features: [{ code: 'word-core', label: 'Word 文字与段落', supported: true, severity: 'info' }] }),
  async open(input) {
    const pkg = await loadOoxmlPackage(input.bytes);
    try {
      if ((await detectKind(pkg, input.name)).kind !== 'docx') throw new OfficeError(ErrorCodes.UNSUPPORTED_FORMAT, '该文件不是 Word 文档（DOCX）。');
      const capability = await analyzeCoreWordCapabilities(pkg, input.knownFonts);
      return { ...opened(await WordCoreSession.open(input.bytes, { readonly: capability.level <= 1 }), capability), displayName: input.name };
    } finally { pkg.dispose(); }
  },
  async serialize() { throw new Error('Word 新核心必须通过当前会话生成保存快照'); },
};
