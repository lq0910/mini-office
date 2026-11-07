/** 功能：验证邮件合并面板挂载不污染普通文档及导入模板；时间：2026-11-07；作者：lq。 */
import { act, createElement, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { expect, it } from 'vitest';
import { createBlankDoc, useWorkspaceStore, type WordSession } from '@mini-office/document-model';
import { useMailMergeController, type MailMergeController } from '../packages/editor-word/src/mail-merge-controller';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import { parseDocx } from '../packages/editor-word/src/parse';

it('已有邮件合并状态的无操作保存字节一致，实际修改配置可以继续保存', async () => {
  const doc = createBlankDoc();
  doc.meta.mailMergeState = JSON.stringify({ search: '原始筛选' });
  const bytes = await serializeDocx(doc, { name: 'template.docx', mailMergeState: doc.meta.mailMergeState });
  const imported = await parseDocx(bytes);
  expect(await serializeDocx(imported, { name: 'template.docx', sourceBytes: bytes, mailMergeState: imported.meta.mailMergeState })).toEqual(bytes);
  imported.meta.mailMergeState = JSON.stringify({ search: '实际修改' });
  const saved = await serializeDocx(imported, { name: 'template.docx', sourceBytes: bytes, mailMergeState: imported.meta.mailMergeState });
  expect((await parseDocx(saved)).meta.mailMergeState).toBe(imported.meta.mailMergeState);
});

it.each([false, true])('面板初始化保留文档元数据，实际配置修改才持久化：已有模板=%s', async imported => {
  (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  const doc = createBlankDoc();
  if (imported) {
    doc.meta.mailMergeState = JSON.stringify({ search: '原始筛选', allowEmpty: false });
    doc.meta.mailMergeTemplate = { dataSourcePath: '/original/source.xlsx', sheetName: 'Sheet1', fields: ['姓名'], conditions: [] };
  }
  const initial = structuredClone(doc.meta);
  const session = { id: `noop-merge-${imported}`, kind: 'docx', data: doc, displayName: '测试.docx', readonly: false } as WordSession;
  useWorkspaceStore.setState({ sessions: { [session.id]: session }, activeId: session.id });
  let controller!: MailMergeController;
  /** 功能：只挂载实际控制器，模拟用户修改筛选而不依赖面板布局；时间：2026-11-07；作者：lq。 */
  function Harness() {
    controller = useMailMergeController({ session, doc, renderBlocks: () => null });
    return null;
  }
  const host = document.createElement('div');
  const root = createRoot(host);
  try {
    await act(async () => root.render(createElement(StrictMode, null, createElement(Harness))));
    expect((useWorkspaceStore.getState().sessions[session.id].data as typeof doc).meta).toEqual(initial);
    expect(controller.allowEmpty).toBe(!imported);
    await act(async () => controller.setSearch('用户筛选'));
    const changed = (useWorkspaceStore.getState().sessions[session.id].data as typeof doc).meta;
    expect(JSON.parse(changed.mailMergeState!).search).toBe('用户筛选');
    expect(changed.mailMergeTemplate).toEqual(initial.mailMergeTemplate);
    await act(async () => controller.setSearch(imported ? '原始筛选' : ''));
    const restored = (useWorkspaceStore.getState().sessions[session.id].data as typeof doc).meta;
    expect(JSON.parse(restored.mailMergeState!).search).toBe(imported ? '原始筛选' : '');
  } finally {
    await act(async () => root.unmount());
    useWorkspaceStore.setState({ sessions: {}, activeId: null });
  }
});
