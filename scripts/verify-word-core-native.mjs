/** 功能：在实际 Tauri WKWebView 内验收默认新核心及原生平台，不模拟字体、文件和恢复 IPC；时间：2026-11-11；作者：lq。 */
const root = decodeURIComponent(new URL('..', import.meta.url).pathname).replace(/^\/@fs/, '').replace(/\/$/, '');
const output = `${root}/output/playwright/word-core-m4.local`;
const report = { status: 'running', engine: navigator.userAgent, checks: [], limitations: ['中文输入法采用组合事件回归；系统候选窗/打印机出纸需人工验收。'] };
// 功能：捕获 React 异步清理/布局异常，全部操作后也必须无未捕获错误；时间：2026-11-11；作者：lq。
const errors = []; window.addEventListener('error', event => errors.push(event.error?.stack ?? event.message));
window.addEventListener('unhandledrejection', event => errors.push(event.reason?.stack ?? String(event.reason)));
const { invoke } = await import('/@fs' + root + '/packages/platform-adapter/node_modules/@tauri-apps/api/core.js');
await invoke('word_smoke_report', { report: JSON.stringify(report) });
/** 功能：等待实际 UI/原生事件，不用固定延时假装任务完成；时间：2026-11-11；作者：lq。 */
async function until(check, label) {
  const start = performance.now();
  while (!check()) { if (performance.now() - start > 30000) throw new Error(`等待超时：${label}；${document.body.innerText.slice(0, 1800)}`); await new Promise(resolve => setTimeout(resolve, 50)); }
}
/** 功能：每项检查失败即阻止将原生验收标记通过；时间：2026-11-11；作者：lq。 */
function assert(value, label) { if (!value) throw new Error(label); report.checks.push(label); void invoke('word_smoke_report', { report: JSON.stringify(report) }); }
try {
  const { useWorkspaceStore: store } = await import('/@fs' + root + '/packages/document-model/src/index.ts');
  const { getPlatformAdapter, coreWordFontMetrics, copyCoreWordImage } = await import('/@fs' + root + '/packages/platform-adapter/src/index.ts');
  await until(() => store.getState().hydrated && store.getState().activeId && store.getState().sessions[store.getState().activeId]?.wordCore, '冷启动中文路径文件队列');
  const first = store.getState().activeId, core = store.getState().sessions[first].wordCore;
  assert(store.getState().sessions[first].path === `${output}/原生 验收.docx`, 'cold-start-file-queue-default-core');
  await until(() => document.querySelector('[data-word-engine="word-core"] .ProseMirror')?.getBoundingClientRect().width > 0, '实际正文');
  const metrics = await coreWordFontMetrics('Songti SC');
  assert(metrics?.unitsPerEm > 0 && metrics.hheaAscender > 0, 'native-system-font-table-metrics');
  // 功能：在真实 WKWebView 检查页间画布色和无灰框纸面，避免浏览器通过而桌面接缝仍透明；时间：2026-11-11；作者：lq。
  await until(() => document.querySelector('.page-gap:not(.page-gap-carry)')?.getBoundingClientRect().height > 28, '原生分页接缝');
  const gap = document.querySelector('.page-gap:not(.page-gap-carry)'), paper = core.editor.view.dom;
  const band = getComputedStyle(gap, '::before'), paperCanvas = getComputedStyle(paper).getPropertyValue('--canvas').trim();
  const probe = document.createElement('div'); probe.style.color = paperCanvas; document.body.append(probe);
  const canvasColor = getComputedStyle(probe).color; probe.remove();
  assert(band.backgroundColor === canvasColor && band.boxShadow === 'none' && getComputedStyle(paper).boxShadow === 'none', 'native-separated-pages-with-borderless-paper');
  const before = core.editor.getText(); core.editor.commands.focus('start'); core.editor.commands.insertContent('原生 WebKit 中文编辑');
  assert(await store.getState().saveSession(first), 'native-atomic-save');
  // 功能：两次独立输入跨过原生历史分组间隔，验证保存不丢历史；时间：2026-11-11；作者：lq。
  await new Promise(resolve => setTimeout(resolve, 600));
  core.editor.commands.insertContent('第二次保存'); assert(await store.getState().saveSession(first), 'native-second-save-with-backup');
  const bytes = await getPlatformAdapter().openPath(`${output}/原生 验收.docx`);
  const second = await store.getState().openBytes({ name: '原生重开.docx', bytes: bytes.bytes });
  assert(store.getState().sessions[second].wordCore.editor.getText().includes('第二次保存'), 'native-saved-docx-reopen');
  store.getState().activateTab(first);
  await until(() => document.querySelector('[data-word-engine="word-core"] .ProseMirror') === core.editor.view.dom, '标签恢复');
  core.editor.commands.undo();
  assert(core.editor.getText().includes('原生 WebKit 中文编辑') && !core.editor.getText().includes('第二次保存'), 'native-tab-switch-retains-history');
  const id = await store.getState().createDocument('docx');
  assert(await store.getState().saveSession(first), 'native-inactive-tab-save');
  await store.getState().closeTab(id, true);
  store.getState().activateTab(first);
  await until(() => document.querySelector('[data-word-engine="word-core"] .ProseMirror') === core.editor.view.dom, '正文重挂载');
  const dom = core.editor.view.dom;
  dom.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: 'zhong' }));
  assert(!await store.getState().saveSession(first), 'native-composition-save-guard');
  dom.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '' }));
  await until(() => !core.editor.view.composing, '结束组合输入');
  store.getState().updateSettings({ autoSave: true, autoSaveDelayMs: 500 });
  core.editor.commands.insertContent('原生快照恢复内容');
  await until(() => store.getState().snapshots.some(snapshot => snapshot.id === first && snapshot.dirty), '实际恢复快照');
  const recovered = await store.getState().openSnapshot(first);
  assert(store.getState().sessions[recovered]?.wordCore.editor.getText().includes('原生快照恢复内容'), 'native-docx-recovery-snapshot');
  const { coreWordOutputBytes, mountCoreWordReport } = await import('/@fs' + root + '/packages/editor-word/src/core-output.ts');
  const { exportElementsToPdf } = await import('/@fs' + root + '/packages/export-service/src/index.ts');
  const snapshot = await coreWordOutputBytes(core), preview = await mountCoreWordReport(snapshot);
  try {
    assert(preview.pages.length > 0 && preview.frames.every(frame => frame.widthPx > 500), 'native-webkit-real-pagination-print-pages');
    const pdf = await exportElementsToPdf({ pages: preview.pages, fileName: '原生输出.pdf', scale: 1 });
    await getPlatformAdapter().saveBytes(pdf, { suggestedName: '原生输出.pdf', path: `${output}/原生输出.pdf` }, 'application/pdf');
    assert(new TextDecoder().decode(pdf.slice(0, 5)) === '%PDF-', 'native-webkit-pdf-real-file');
  } finally { preview.dispose(); }
  const canvas = document.createElement('canvas'); canvas.width = 20; canvas.height = 10; canvas.getContext('2d').fillRect(0, 0, 20, 10);
  assert(await copyCoreWordImage(canvas.toDataURL(), JSON.stringify({ imageWidthPx: 80, imageHeightPx: 40, imageWrap: 'inline' })), 'native-png-and-html-pasteboard-write');
  assert(window.__docsParseStats.worker > 0, 'native-module-worker-docx-parse');
  report.parse = window.__docsParseStats; report.originalCharacters = before.length;
  await store.getState().closeAll(true);
  await new Promise(resolve => setTimeout(resolve, 200));
  assert(core.editor.isDestroyed, 'native-close-releases-editors');
  report.errors = errors;
  assert(!errors.length && !document.body.innerText.includes('界面遇到问题'), 'native-no-react-cleanup-errors');
  report.status = 'passed';
} catch (error) { report.status = 'failed'; report.error = String(error); report.stack = error.stack; report.errors = errors; report.body = document.body.innerText.slice(0, 1800); }
await invoke('word_smoke_report', { report: JSON.stringify(report, null, 2) });
