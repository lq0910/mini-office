/** 功能：启动时只注册轻量格式代理，首次打开或保存时才加载对应解析器；时间：2026-10-24；作者：lq。 */
import {
  createBlankDoc,
  createBlankWorkbook,
  createBlankDeck,
  registerFormatHandler,
  type AnyFormatHandler,
  type PreviewDocumentData,
} from '@mini-office/document-model';

/** 功能：合并并发模块加载，加载失败后允许下一次操作重试；时间：2026-10-24；作者：lq。 */
function registerDeferredHandler(
  kind: AnyFormatHandler['kind'],
  extensions: string[],
  createBlank: AnyFormatHandler['createBlank'],
  loader: () => Promise<AnyFormatHandler>,
  asyncNew = false,
): void {
  let pending: Promise<AnyFormatHandler> | undefined;
  const load = () => pending ??= loader().catch((error) => {
    pending = undefined;
    throw error;
  });
  registerFormatHandler({
    kind,
    extensions,
    createBlank,
    // 功能：保持异步核心初始化按需加载；旧处理器仍沿用同步空白模型；时间：2026-11-09；作者：lq。
    createNew: asyncNew ? async () => {
      const handler = await load();
      return handler.createNew ? handler.createNew() : { data: handler.createBlank(), capability: { level: 3, warnings: [], features: [], missingFonts: [], scannedAt: Date.now() } };
    } : undefined,
    open: async (input) => (await load()).open(input),
    serialize: async (data, context) => (await load()).serialize(data, context),
  });
}

/** 功能：保留新建文档的即时响应，同时将三种格式解析代码移出首页启动路径；时间：2026-10-24；作者：lq。 */
export function registerDeferredFormats(): void {
  // 功能：正式默认启用新核心；仅开发构建允许显式回退，整份会话固定引擎且保存失败不回退；时间：2026-11-11；作者：lq。
  const coreWord = !(import.meta.env.DEV && import.meta.env.VITE_WORD_LEGACY === 'true');
  registerDeferredHandler('docx', ['.docx', '.docm', '.dotx'], coreWord ? () => ({ engine: 'word-core' as const }) : createBlankDoc,
    coreWord
      ? () => import('../../../packages/editor-word/src/core-handler').then(module => module.coreWordFormatHandler)
      : () => import('../../../packages/editor-word/src/handler').then((module) => module.wordFormatHandler), coreWord);
  registerDeferredHandler('xlsx', ['.xlsx', '.xlsm', '.xltx', '.csv'], createBlankWorkbook,
    () => import('../../../packages/editor-sheet/src/handler').then((module) => module.sheetFormatHandler));
  registerDeferredHandler('pptx', ['.pptx', '.pptm', '.potx'], createBlankDeck,
    () => import('../../../packages/editor-slides/src/handler').then((module) => module.slidesFormatHandler));
  // 功能：注册零依赖的 PDF、Markdown、HTML 只读预览处理器，解析代码保持在主包内且不引入大型库；时间：2026-10-31；作者：lq。
  const createPreview = (format: PreviewDocumentData['format']): PreviewDocumentData => ({ format });
  (['pdf', 'md', 'html'] as const).forEach((kind) => {
    registerFormatHandler({
      kind,
      extensions: kind === 'pdf' ? ['.pdf'] : kind === 'md' ? ['.md', '.markdown'] : ['.html', '.htm'],
      createBlank: () => createPreview(kind),
      async open(input) {
        const text = kind === 'pdf' ? undefined : new TextDecoder().decode(input.bytes);
        return {
          data: { format: kind, bytes: kind === 'pdf' ? input.bytes : undefined, text },
          readonly: true,
          displayName: input.name,
          capability: { level: 1, warnings: [], features: [{ code: 'preview', label: '只读预览', supported: true, severity: 'info' }], missingFonts: [], scannedAt: Date.now() },
        };
      },
      async serialize(data) {
        const preview = data as PreviewDocumentData;
        return preview.bytes ?? new TextEncoder().encode(preview.text ?? '');
      },
    });
  });
}
