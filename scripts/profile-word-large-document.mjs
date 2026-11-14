/** 功能：用合成大文档采样真实 Word 宿主的加载、分页、滚动与编辑开销；时间：2026-11-13；作者：lq。 */
import { chromium } from 'playwright-core';
import JSZip from 'jszip';
import { createRequire } from 'node:module';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { homedir } from 'node:os';

const root = resolve(import.meta.dirname, '..');
const output = join(root, 'output/playwright/word-performance.local');
const require = createRequire(join(root, 'package.json'));
const { build } = createRequire(require.resolve('vite/package.json'))('esbuild');
await mkdir(output, { recursive: true });
// 功能：构建真实宿主并导出采样入口，性能探针只进入此隔离产物；时间：2026-11-13；作者：lq。
await build({ absWorkingDir: root, stdin: { contents: `export { WordCoreSurface } from './packages/editor-word/src/core-surface'; export { WordSurfaceLayout } from './packages/editor-word/src/core-layout'; export { WordLayoutController } from './packages/word-core/src/layout/controller'; export { configureWordPlatform } from './packages/word-core/src/platform/ports';`, resolveDir: root }, outfile: join(output, 'surface.js'), bundle: true, format: 'esm', platform: 'browser', target: 'es2022', loader: { '.ttf': 'file', '.otf': 'file', '.woff': 'file', '.woff2': 'file' } });
await build({ absWorkingDir: root, entryPoints: { styles: 'packages/word-core/src/styles/document.css', fonts: 'packages/word-core/src/styles/fonts.css', 'parse-worker': 'packages/word-core/src/upstream/apps/docs/src/renderer/parse-worker.ts' }, outdir: output, bundle: true, format: 'esm', platform: 'browser', target: 'es2022', loader: { '.ttf': 'file', '.otf': 'file', '.woff': 'file', '.woff2': 'file' } });
const origin = process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1431';
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
// 功能：同一采样脚本支持完整产品外壳，区分内核与 React 页码订阅的开销；时间：2026-11-13；作者：lq。
const shell = process.env.WORD_PROFILE_SHELL === '1';
const report = { date: '2026-11-13', browser: browser.version(), scope: `${shell ? 'actual app with React shell' : 'isolated production WordCoreSurface'}; synthetic text; Chromium; not native WKWebView`, samples: [] };
const reportFile = join(output, shell ? 'report-shell.json' : 'report.json');
try {
  const sizes = (process.env.WORD_PROFILE_PAGES ?? '20,100,300').split(',').map(Number);
  for (const pages of sizes) {
    const zip = await JSZip.loadAsync(await readFile(join(root, 'tests/fixtures/docx/contract.docx')));
    // 功能：可选混合表格样本覆盖分批挂载时的真实 NodeView、页缝及保存来源；时间：2026-11-14；作者：lq。
    const table = process.env.WORD_PROFILE_TABLES === '1' ? `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${Array.from({ length: 3 }, (_, row) => `<w:tr>${Array.from({ length: 2 }, (_, column) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>表格来源${row}-${column}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>` : '';
    // 功能：显式分页构造确定规模的文本样本，不读取或改写用户文档；时间：2026-11-13；作者：lq。
    const body = Array.from({ length: pages * 24 }, (_, index) => `<w:p><w:pPr>${index && index % 24 === 0 ? '<w:pageBreakBefore/>' : ''}<w:spacing w:after="80"/></w:pPr><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>第${index + 1}段：智慧物业信息系统部署实施，长文档加载与滚动性能采样。</w:t></w:r></w:p>${index % 24 === 10 ? table : ''}`).join('');
    zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
    const bytes = Array.from(await zip.generateAsync({ type: 'uint8array' }));
    await writeFile(join(output, `synthetic-${pages}.docx`), new Uint8Array(bytes));
    const page = await browser.newPage({ viewport: { width: 1100, height: 900 } });
    page.setDefaultTimeout(180000);
    if (!shell) await page.route('**/__word_profile__', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html><html><head><meta charset="utf-8"></head><body style="margin:0"><div class="mo-scroll" style="height:900px;overflow:auto"><div class="doc-zoom" style="zoom:1"><div id="core" class="mo-word-core"></div></div></div></body></html>' }));
    await page.goto(shell ? origin : `${origin}/__word_profile__`);
    const base = `${origin}/@fs${output}`;
    if (!shell) for (const file of ['styles.css', 'fonts.css', 'surface.css']) await page.addStyleTag({ url: `${base}/${file}` });
    const opened = await page.evaluate(async ({ base, bytes, shell, root }) => {
      const mod = shell ? Object.assign({}, await import(`/@fs${root}/packages/editor-word/src/core-surface.ts`), await import(`/@fs${root}/packages/editor-word/src/core-layout.ts`), await import(`/@fs${root}/packages/word-core/src/layout/controller.ts`)) : await import(`${base}/surface.js`);
      if (!shell) mod.configureWordPlatform({ createParseWorker: () => new Worker(`${base}/parse-worker.js`, { type: 'module' }) });
      window.__profile = { passes: [], longTasks: [], errors: [], layouts: 0 };
      // 功能：记录整轮分页和应用装饰时间，不在高频 DOM API 上插桩干扰结果；时间：2026-11-13；作者：lq。
      for (const [type, method] of [[mod.WordLayoutController, 'flush'], [mod.WordSurfaceLayout, 'apply']]) {
        const original = type.prototype[method];
        type.prototype[method] = function (...args) {
          const start = performance.now();
          try { return original.apply(this, args); }
          finally {
            window.__profile.passes.push({ method, start, ms: performance.now() - start });
            if (method === 'apply' && !window.__profile.firstLayoutAt) window.__profile.firstLayoutAt = performance.now();
          }
        };
      }
      new PerformanceObserver(list => window.__profile.longTasks.push(...list.getEntries().map(entry => ({ start: entry.startTime, ms: entry.duration })))).observe({ type: 'longtask', buffered: true });
      if (shell) {
        const { useWorkspaceStore } = await import(`/@fs${root}/packages/document-model/src/index.ts`);
        window.__workspace = useWorkspaceStore;
        const start = performance.now();
        window.__profile.openStarted = start;
        const id = await useWorkspaceStore.getState().openBytes({ name: '性能采样.docx', bytes: new Uint8Array(bytes) });
        const session = useWorkspaceStore.getState().sessions[id].wordCore;
        while (!mod.WordCoreSurface.current(session)) {
          if (performance.now() - start > 30000) throw new Error(`产品正文挂载超时：${document.body.innerText.slice(0, 1500)}`);
          await new Promise(resolve => setTimeout(resolve, 20));
        }
        window.__surface = mod.WordCoreSurface.current(session);
        return { mountMs: performance.now() - start };
      }
      const start = performance.now();
      window.__profile.openStarted = start;
      window.__surface = await mod.WordCoreSurface.mount(document.getElementById('core'), new Uint8Array(bytes), { onLayout: () => window.__profile.layouts++, onLayoutError: error => window.__profile.errors.push(String(error)) });
      return { mountMs: performance.now() - start };
    }, { base, bytes, shell, root });
    // 功能：等待分页静止后采样，避免把首次排版与稳定滚动混成一个指标；时间：2026-11-13；作者：lq。
    try {
      await page.waitForFunction(() => window.__surface.layout.isSettled() && window.__profile.passes.length && performance.now() - window.__profile.passes.at(-1).start - window.__profile.passes.at(-1).ms > 1500, undefined, { timeout: 60000 });
    } catch (error) {
      const diagnostic = await page.evaluate(() => ({ profile: window.__profile, state: window.__surface.session.state(), pageCount: window.__surface.layout.current()?.pageCount, stats: window.__surface.layout.performanceStats, pending: window.__surface.layout.controller.hasPendingPass(), documentMatches: window.__surface.layout.lastDocument === window.__surface.session.editor.state.doc }));
      await writeFile(join(output, 'failure.json'), JSON.stringify(diagnostic, null, 2));
      throw error;
    }
    const layout = await page.evaluate(() => ({ firstLayoutMs: window.__profile.firstLayoutAt - window.__profile.openStarted, pageCount: window.__surface.layout.current().pageCount, domNodes: window.__surface.root.querySelectorAll('*').length, parseStats: window.__docsParseStats, stats: { ...window.__surface.layout.performanceStats }, ...window.__profile }));
    const scrolls = [];
    // 功能：交换开关顺序并重复滚动，避免把冷缓存与层提升开销混淆；时间：2026-11-13；作者：lq。
    for (const gpu of [false, true, true, false]) {
      const scroll = await page.evaluate(async gpu => {
        const override = document.createElement('style');
        override.textContent = gpu ? '.mo-word-core .page-gap {transform:translateZ(0)!important;will-change:transform!important}.mo-word-core .doc-col-block {will-change:transform!important}' : '.mo-word-core .page-gap {transform:none!important;will-change:auto!important}.mo-word-core .doc-col-block {will-change:auto!important}';
        document.head.append(override);
        const viewport = window.__surface.root.closest('.mo-scroll');
        viewport.scrollTop = 0;
        await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        const before = window.__profile.passes.length;
        const frames = [];
        let previous = performance.now();
        const start = previous;
        // 功能：每帧移动三百像素并记录帧间隔；这是 JS 滚动采样，不冒充合成器呈现帧率；时间：2026-11-13；作者：lq。
        for (let index = 0; index < 180; index++) {
          await new Promise(resolve => requestAnimationFrame(resolve));
          const now = performance.now(); frames.push(now - previous); previous = now;
          viewport.scrollTop = Math.min(viewport.scrollHeight - viewport.clientHeight, index * 300);
        }
        override.remove();
        const sorted = [...frames].sort((a, b) => a - b);
        return { gpu, elapsedMs: performance.now() - start, p95FrameMs: sorted[Math.floor(sorted.length * 0.95)], maxFrameMs: Math.max(...frames), framesOver34Ms: frames.filter(ms => ms > 34).length, paginationPasses: window.__profile.passes.length - before };
      }, gpu);
      scrolls.push(scroll);
    }
    const edit = await page.evaluate(async () => {
      const editor = window.__surface.session.editor;
      const before = window.__profile.passes.length;
      const started = performance.now();
      editor.commands.insertContentAt(2, '测');
      const transactionMs = performance.now() - started;
      await new Promise(resolve => setTimeout(resolve, 3500));
      return { transactionMs, passes: window.__profile.passes.slice(before), stats: { ...window.__surface.layout.performanceStats }, dirty: window.__surface.session.state().dirty };
    });
    let verification;
    if (process.env.WORD_PROFILE_VERIFY === '1') {
      // 功能：使尾部段落真实增高，比较续排与随后强制全量排版的切点，验证性能缓存不会牺牲分页正确性；时间：2026-11-13；作者：lq。
      const changed = await page.evaluate(() => {
        const editor = window.__surface.session.editor;
        const before = { ...window.__surface.layout.performanceStats };
        let pos = 0;
        const target = Math.max(0, editor.state.doc.childCount - 30);
        for (let index = 0; index < target; index++) pos += editor.state.doc.child(index).nodeSize;
        editor.commands.insertContentAt(pos + 2, '尾部重排验收内容。'.repeat(180));
        return before;
      });
      await page.waitForFunction(() => window.__surface.layout.isSettled());
      verification = await page.evaluate(before => {
        const surface = window.__surface;
        const snapshot = () => JSON.stringify(surface.layout.current().cuts.map(cut => [cut.start, cut.end, cut.section, cut.pageNo]));
        const incremental = snapshot(), after = { ...surface.layout.performanceStats };
        surface.layout.flush();
        const a = JSON.parse(incremental), b = JSON.parse(snapshot());
        // 功能：浏览器缩放/滚动产生亚像素浮点误差；页数、分节和编号严格一致，切点容差小于 0.01px；时间：2026-11-13；作者：lq。
        const equivalent = a.length === b.length && a.every((cut, index) => Math.abs(cut[0] - b[index][0]) < 0.01 && Math.abs(cut[1] - b[index][1]) < 0.01 && cut[2] === b[index][2] && cut[3] === b[index][3]);
        if (!equivalent) {
          const differences = a.flatMap((cut, index) => JSON.stringify(cut) === JSON.stringify(b[index]) ? [] : [{ page: index, incremental: cut, full: b[index] }]);
          throw new Error(`续排与完整分页的切点不同：${JSON.stringify({ counts: [a.length, b.length], differences: differences.slice(0, 6), before, after })}`);
        }
        if (after.resumedPasses <= before.resumedPasses) throw new Error('尾部增高没有命中安全续排');
        return { incrementalMatchesFull: true, before, after, pageCount: surface.layout.current().pageCount };
      }, changed);
      // 功能：保存实际编辑并在同一宿主重开，核对尾部完整性、无修改重存及撤销隔离；时间：2026-11-13；作者：lq。
      const saved = await page.evaluate(async () => Array.from((await window.__surface.session.prepareSave()).bytes));
      const savedZip = await JSZip.loadAsync(new Uint8Array(saved));
      const xml = await savedZip.file('word/document.xml').async('text');
      if (!xml.includes(`第${pages * 24}段`) || !xml.includes('尾部重排验收内容。') || !xml.includes('测')) throw new Error('长文档保存缺少尾部或编辑内容');
      await writeFile(join(output, 'verified-saved.docx'), new Uint8Array(saved));
      const reopened = await page.evaluate(async ({ base, saved }) => {
        const mod = await import(`${base}/surface.js`);
        const container = window.__surface.root;
        window.__surface.dispose();
        window.__surface = await mod.WordCoreSurface.mount(container, new Uint8Array(saved));
        await window.__surface.session.whenContentReady();
        const roundtrip = await window.__surface.session.prepareSave();
        return { bytes: Array.from(roundtrip.bytes), undo: window.__surface.session.state().canUndo };
      }, { base, saved });
      if (reopened.undo || Buffer.compare(Buffer.from(saved), Buffer.from(reopened.bytes))) throw new Error('重开后无操作保存或撤销隔离失败');
      verification.saveReopenNoop = true;
    }
    const sample = { requestedPages: pages, ...opened, ...layout, scrolls, edit, verification };
    report.samples.push(sample);
    console.log(JSON.stringify(sample));
    await page.close();
    await writeFile(reportFile, JSON.stringify(report, null, 2));
  }
} finally { await browser.close(); }
