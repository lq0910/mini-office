/** 功能：在真实 macOS WKWebView 中采样合成 300 页文档，使用独立验收数据目录；时间：2026-11-13；作者：lq。 */
const root = decodeURIComponent(new URL('..', import.meta.url).pathname).replace(/^\/@fs/, '').replace(/\/$/, '');
const { invoke } = await import('/@fs' + root + '/packages/platform-adapter/node_modules/@tauri-apps/api/core.js');
const report = { status: 'running', engine: navigator.userAgent, scope: 'actual Tauri WKWebView + React shell; synthetic 300 pages; rAF sampling, not compositor presentation FPS', passes: [], errors: [] };
window.addEventListener('error', event => report.errors.push(event.error?.stack ?? event.message));
window.addEventListener('unhandledrejection', event => report.errors.push(event.reason?.stack ?? String(event.reason)));
/** 功能：每个关键阶段写入诊断，失败或超时不会被误报为验收通过；时间：2026-11-13；作者：lq。 */
const publish = () => invoke('word_smoke_report', { report: JSON.stringify(report) });
async function until(check, label) {
  const start = performance.now();
  while (!check()) { if (performance.now() - start > 90000) throw new Error(`等待超时：${label}`); await new Promise(resolve => setTimeout(resolve, 25)); }
}
await publish();
try {
  report.stage = 'import-store'; await publish();
  const { useWorkspaceStore: store } = await import('/@fs' + root + '/packages/document-model/src/index.ts');
  report.stage = 'import-surface'; await publish();
  const { WordCoreSurface } = await import('/@fs' + root + '/packages/editor-word/src/core-surface.ts');
  report.stage = 'import-controller'; await publish();
  const { WordLayoutController } = await import('/@fs' + root + '/packages/word-core/src/layout/controller.ts');
  const { WordSurfaceLayout } = await import('/@fs' + root + '/packages/editor-word/src/core-layout.ts');
  report.stage = 'hydrate'; await publish();
  await until(() => store.getState().hydrated, '工作区初始化');
  // 功能：可选诊断开关仅在隔离验收中移除 :has 规则，用于证明 WebKit 样式失效开销；默认关闭；时间：2026-11-13；作者：lq。
  if (window.__wordProfileStripHas) {
    report.removedHas = [];
    const strip = sheet => {
      const rules = sheet.cssRules;
      if (!rules) return;
      for (let index = rules.length - 1; index >= 0; index--) {
        const rule = rules[index];
        const mode = window.__wordProfileStripHas;
        const eligible = mode === 'universal' ? /\.doc-page\s+\*?:has\(/.test(rule.selectorText ?? '') : rule.selectorText?.includes(':has(') && (mode !== 'static' || rule.selectorText.includes('.mo-word-core')) && (mode !== 'dynamic' || !rule.selectorText.includes('.mo-word-core'));
        if (eligible) { report.removedHas.push(rule.selectorText); sheet.deleteRule(index); }
        else if (rule.cssRules) strip(rule);
      }
    };
    const stripAll = () => { for (const sheet of document.styleSheets) { try { strip(sheet); } catch {} } };
    stripAll();
    const observer = new MutationObserver(stripAll); observer.observe(document.documentElement, { childList: true, subtree: true });
  }
  for (const [type, method] of [[WordLayoutController, 'flush'], [WordSurfaceLayout, 'apply']]) {
    const original = type.prototype[method];
    type.prototype[method] = function (...args) {
      const start = performance.now();
      try { return original.apply(this, args); }
      finally { report.passes.push({ method, start, ms: performance.now() - start }); if (method === 'apply' && !report.firstLayoutAt) report.firstLayoutAt = performance.now(); }
    };
  }
  const bytes = new Uint8Array(await (await fetch('/@fs' + root + '/output/playwright/word-performance.local/synthetic-300.docx')).arrayBuffer());
  const start = performance.now();
  report.stage = 'open'; await publish();
  if (window.__wordProfileDomStack) {
    // 功能：仅诊断时低频记录 WebKit DOM 插入的 JavaScript 来源，关联原生 CPU 栈；时间：2026-11-14；作者：lq。
    const insert = Node.prototype.insertBefore;
    let recorded = 0;
    Node.prototype.insertBefore = function (...args) {
      if (performance.now() - recorded > 1500) {
        recorded = performance.now(); report.domStack = { parent: this.nodeName, connected: this.isConnected, stack: new Error().stack }; void publish();
      }
      return insert.apply(this, args);
    };
  }
  // 功能：原生接口或开发模块若未返回，保留具体打开状态并明确超时；时间：2026-11-13；作者：lq。
  let timeout;
  const heartbeat = setInterval(() => {
    report.openDiagnostic = { busy: store.getState().busy, fonts: store.getState().fonts.length, text: document.body.innerText.slice(0, 180), parse: window.__docsParseStats };
    void publish();
  }, 5000);
  let id;
  try {
    id = await Promise.race([store.getState().openBytes({ name: '原生性能验收300页.docx', bytes }), new Promise((_, reject) => { timeout = setTimeout(() => reject(new Error('原生打开 90 秒内未完成')), 90000); })]);
  } finally { clearTimeout(timeout); clearInterval(heartbeat); }
  const core = store.getState().sessions[id].wordCore;
  await until(() => WordCoreSurface.current(core), '正文挂载');
  const surface = WordCoreSurface.current(core);
  report.mountMs = performance.now() - start;
  await core.whenContentReady(); report.contentReadyMs = performance.now() - start; await publish();
  await until(() => surface.layout.isSettled() && report.passes.length && performance.now() - report.passes.at(-1).start - report.passes.at(-1).ms > 1200, '完整分页收敛');
  report.firstLayoutMs = report.firstLayoutAt - start;
  report.pageCount = surface.layout.current().pageCount;
  report.parseStats = window.__docsParseStats;
  report.domNodes = surface.root.querySelectorAll('*').length;
  report.focus = { hidden: document.hidden, hasFocus: document.hasFocus() };
  // 功能：原生后台限帧与前台渲染开销必须分开报告，禁止将此诊断冒充真实滚动帧率；时间：2026-11-14；作者：lq。
  report.limitations = report.focus.hasFocus ? [] : ['窗口未获得前台焦点；rAF 时间含 WebKit 后台限帧，不能用作前台滚动 FPS 验收。'];
  report.originalHasRulesRetained = !window.__wordProfileStripHas;
  if (report.pageCount !== 300) throw new Error(`完整页数不符：${report.pageCount}`);
  const viewport = surface.root.closest('.mo-scroll');
  const idle = []; let idleLast = performance.now();
  for (let index = 0; index < 20; index++) { await new Promise(resolve => requestAnimationFrame(resolve)); const now = performance.now(); idle.push(now - idleLast); idleLast = now; }
  report.idleFrameMs = idle.reduce((sum, value) => sum + value, 0) / idle.length;
  const frames = [], before = report.passes.length;
  let previous = performance.now();
  // 功能：稳定后测量真实产品滚动，检查滚动不会启动分页；时间：2026-11-13；作者：lq。
  for (let index = 0; index < 180; index++) {
    await new Promise(resolve => requestAnimationFrame(resolve));
    const now = performance.now(); frames.push(now - previous); previous = now;
    viewport.scrollTop = Math.min(viewport.scrollHeight - viewport.clientHeight, index * 300);
  }
  const sorted = [...frames].sort((a, b) => a - b);
  report.scroll = { p95FrameMs: sorted[Math.floor(sorted.length * 0.95)], maxFrameMs: Math.max(...frames), framesOver34Ms: frames.filter(ms => ms > 34).length, paginationPasses: report.passes.length - before };
  if (window.__wordProfileStripHas === true) {
    // 功能：隔离诊断整篇阴影与 React 页码订阅的滚动成本，不进入正常产品；时间：2026-11-13；作者：lq。
    report.scrollDiagnostics = [];
    for (const scenario of ['no-paper-shadow', 'no-scroll-notification']) {
      const style = document.createElement('style');
      style.textContent = '.mo-word-core .doc-page {box-shadow:none!important;isolation:auto!important}'; document.head.append(style);
      const block = event => { if (event.target === viewport) event.stopImmediatePropagation(); };
      if (scenario === 'no-scroll-notification') document.addEventListener('scroll', block, true);
      const timing = []; let last = performance.now();
      for (let index = 0; index < 60; index++) {
        await new Promise(resolve => requestAnimationFrame(resolve));
        const now = performance.now(); timing.push(now - last); last = now;
        viewport.scrollTop = index * 300;
      }
      const ordered = [...timing].sort((a, b) => a - b);
      report.scrollDiagnostics.push({ scenario, p95: ordered[Math.floor(ordered.length * 0.95)], max: Math.max(...timing) });
      style.remove(); document.removeEventListener('scroll', block, true); await publish();
    }
  }
  const editStart = performance.now(), passBefore = { ...surface.layout.performanceStats };
  core.editor.commands.insertContentAt(2, '原生性能验收');
  report.transactionMs = performance.now() - editStart;
  await until(() => surface.layout.isSettled(), '编辑后分页检查');
  report.stats = { ...surface.layout.performanceStats };
  if (report.stats.fastPasses <= passBefore.fastPasses) throw new Error('原生编辑没有命中同源快速分页');
  if (report.errors.length) throw new Error('原生页面出现异步异常');
  report.status = window.__wordProfileStripHas ? 'diagnostic-complete' : 'passed';
} catch (error) { report.status = 'failed'; report.error = error.stack ?? String(error); }
await publish();
