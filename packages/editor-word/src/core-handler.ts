/** 功能：M2 开发入口直接打开完整 GenOffice 衍生 Word 会话，杜绝旧模型保存旁路；时间：2026-11-09；作者：lq。 */
import type { CoreWordDocument, FormatHandler, FormatOpenResult } from '@mini-office/document-model';
import { WordCoreSession } from '../../word-core/src/session/word-session';

/** 功能：能力来自新核心实际会话，保留原始未知 XML 不沿用旧解析器的功能降级判断；时间：2026-11-09；作者：lq。 */
function opened(wordCore: WordCoreSession): FormatOpenResult<CoreWordDocument> {
  const readonly = !wordCore.editor.isEditable;
  return { data: { engine: 'word-core' }, wordCore, readonly,
    capability: { level: readonly ? 1 : 3, warnings: [], missingFonts: [], scannedAt: Date.now(),
      features: [{ code: 'word-core', label: 'Word 文档', supported: true, severity: 'info' }] } };
}

/** 功能：新建、打开均创建同源会话，实际序列化只能从工作区持有的会话发起；时间：2026-11-09；作者：lq。 */
export const coreWordFormatHandler: FormatHandler<CoreWordDocument> = {
  kind: 'docx', extensions: ['.docx', '.docm', '.dotx'],
  createBlank: () => ({ engine: 'word-core' }),
  createNew: async () => opened(await WordCoreSession.create()),
  open: async input => ({ ...opened(await WordCoreSession.open(input.bytes)), displayName: input.name }),
  async serialize() { throw new Error('Word 新核心必须通过当前会话生成保存快照'); },
};
