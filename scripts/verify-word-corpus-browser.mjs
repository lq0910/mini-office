/**
 * 功能：打开固定公开语料，逐页截图并与显式提供的 Office PNG 基准比较，缺失基准不能算保真通过。
 * 时间：2026-11-03；作者：lq。运行：node scripts/verify-word-corpus-browser.mjs [--require-office]
 */
import { chromium } from 'playwright-core';
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { createRequire } from 'node:module';
import { readFile, writeFile, mkdir, access, realpath } from 'node:fs/promises';
import { dirname, resolve, join, relative, isAbsolute } from 'node:path';
import { homedir } from 'node:os';
import { baselinePath, validateOfficeBaseline } from './word-fidelity-baseline.mjs';

const root = resolve(import.meta.dirname, '..');
const fixtureRoot = join(root, 'third_party/genoffice/fixtures');
const output = join(root, 'output/playwright/word-corpus');
const corpus = JSON.parse(await readFile(join(fixtureRoot, 'manifest.json'), 'utf8'));
// 功能：默认发现已导出的本机基准，显式环境变量仍优先，避免重复提示用户重新导出；时间：2026-11-04；作者：lq。
const localBaselineRoot = join(root, 'output/office-baselines.local');
const baselineRoot = process.env.MINI_OFFICE_OFFICE_BASELINES
  ?? await access(localBaselineRoot).then(() => localBaselineRoot, (error) => {
    if (error.code === 'ENOENT') return undefined;
    throw error;
  });
const strict = process.argv.includes('--require-office');
const selected = process.env.MINI_OFFICE_CORPUS_IDS?.split(',').filter(Boolean);
const fixtures = selected ? corpus.fixtures.filter((item) => selected.includes(item.id)) : corpus.fixtures;
if (selected?.some((id) => !fixtures.some((item) => item.id === id))) throw new Error('存在未知语料 ID');
if (!fixtures.length) throw new Error('语料清单为空');
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
const delay = (ms) => new Promise((done) => setTimeout(done, ms));
const url = process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:4196/';
const report = { generatedAt: new Date().toISOString(), commit: corpus.commit, dpi: 96, deviceScaleFactor: 1,
  viewport: { width: 1440, height: 1800 }, baselineRoot: baselineRoot ?? null, browser: '', samples: [] };
await mkdir(output, { recursive: true });
let server;
let browser;

/** 功能：按需启动本地 Vite 并等待就绪，不重用未知端口的旧服务；时间：2026-11-03；作者：lq。 */
async function startServer() {
  if (!process.env.MINI_OFFICE_TEST_URL) {
    const require = createRequire(import.meta.url);
    const vite = join(dirname(require.resolve('vite/package.json')), 'bin/vite.js');
    server = spawn(process.execPath, [vite, '--host', '127.0.0.1', '--port', '4196', '--strictPort'],
      { cwd: join(root, 'apps/desktop'), stdio: ['ignore', 'pipe', 'pipe'] });
    let diagnostics = '';
    let startError;
    server.on('error', (error) => { startError = error; });
    server.stdout.on('data', (chunk) => { diagnostics = (diagnostics + chunk).slice(-6000); });
    server.stderr.on('data', (chunk) => { diagnostics = (diagnostics + chunk).slice(-6000); });
    for (let index = 0; index < 80; index++) {
      if (startError) throw startError;
      if (server.exitCode !== null) throw new Error(`Vite 启动失败：${diagnostics}`);
      // 功能：必须观察到本次 Vite 就绪消息后才连接，避免端口冲突时误测其他服务；时间：2026-11-03；作者：lq。
      if (diagnostics.includes('Local:')) break;
      await delay(200);
    }
    if (!diagnostics.includes('Local:')) throw new Error(`Vite 未就绪：${diagnostics}`);
  }
  for (let index = 0; index < 60; index++) {
    try { if ((await fetch(url, { signal: AbortSignal.timeout(1000) })).ok) return; } catch {}
    await delay(200);
  }
  throw new Error(`无法连接测试地址 ${url}`);
}

/** 功能：只读读取外部基准，并阻断符号链接越出指定目录；时间：2026-11-03；作者：lq。 */
async function baselineFile(directory, name) {
  const [base, file] = await Promise.all([realpath(directory), realpath(baselinePath(directory, name))]);
  const inside = relative(base, file);
  if (inside === '..' || inside.startsWith('../') || isAbsolute(inside)) throw new Error('基准符号链接越界');
  return readFile(file);
}

/** 功能：比较原尺寸 PNG 的 RGB 像素，仅报告差异，不用重采样掩盖页面尺寸不符；时间：2026-11-03；作者：lq。 */
async function comparePng(page, actual, expected) {
  return page.evaluate(async ({ actual, expected }) => {
    const decode = async (encoded) => {
      const image = new Image();
      image.src = `data:image/png;base64,${encoded}`;
      await image.decode();
      const canvas = document.createElement('canvas');
      canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
      const context = canvas.getContext('2d');
      context.fillStyle = '#fff'; context.fillRect(0, 0, canvas.width, canvas.height); context.drawImage(image, 0, 0);
      return { width: canvas.width, height: canvas.height, data: context.getImageData(0, 0, canvas.width, canvas.height).data };
    };
    const [a, b] = await Promise.all([decode(actual), decode(expected)]);
    if (a.width !== b.width || a.height !== b.height) return { status: 'dimension-mismatch',
      actualSize: [a.width, a.height], expectedSize: [b.width, b.height] };
    let changed = 0; let absoluteError = 0;
    for (let index = 0; index < a.data.length; index += 4) {
      let delta = 0;
      for (let channel = 0; channel < 3; channel++) delta += Math.abs(a.data[index + channel] - b.data[index + channel]);
      if (delta) changed++;
      absoluteError += delta;
    }
    return { status: changed ? 'difference-detected' : 'pixel-identical', changedPixelRatio: changed / (a.width * a.height),
      meanAbsoluteChannelError: absoluteError / (a.width * a.height * 3) };
  }, { actual: actual.toString('base64'), expected: expected.toString('base64') });
}

/** 功能：安全生成离线对照页，样本内容与外部环境信息作为纯文本处理；时间：2026-11-03；作者：lq。 */
async function writeReport() {
  const escape = (value) => String(value).replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
  // 功能：把逐页原尺寸像素差异写入可浏览报告，尺寸不符时直接列出两侧尺寸；时间：2026-11-04；作者：lq。
  const comparisonText = (comparison) => !comparison ? '' : comparison.status === 'dimension-mismatch'
    ? `页面尺寸不同：Mini Office ${comparison.actualSize.join(' × ')}；Office ${comparison.expectedSize.join(' × ')}`
    : `差异像素 ${(comparison.changedPixelRatio * 100).toFixed(3)}% · 通道平均绝对误差 ${comparison.meanAbsoluteChannelError.toFixed(3)} / 255`;
  const cards = report.samples.map((sample) => `<section><h2>${escape(sample.id)}</h2><p>${escape(sample.status)} · Mini Office ${sample.pageCount ?? '?'} 页</p>
    ${sample.error ? `<pre>${escape(sample.error)}</pre>` : ''}
    ${sample.baseline ? `<p>${escape(sample.baseline.application)} ${escape(sample.baseline.version)} · ${escape(sample.baseline.operatingSystem)} · ${escape(sample.baseline.fonts.join(', '))}</p>` : '<p>尚无已绑定的 Office 基准，当前结果仅为浏览器回归。</p>'}
    ${(sample.pages ?? []).map((page) => `<h3>第 ${page.number} 页 · ${escape(page.comparison?.status ?? 'pending-office-baseline')}</h3><p>${escape(comparisonText(page.comparison))}</p><div class=pair><figure><figcaption>Mini Office / Chromium</figcaption><img src="${escape(page.file)}"></figure>
    ${page.baselineFile ? `<figure><figcaption>外部 Office 基准</figcaption><img src="${escape(page.baselineFile)}"></figure>` : ''}</div>`).join('')}</section>`).join('');
  report.summary = { planned: fixtures.length, captured: report.samples.filter((sample) => sample.pages?.length).length,
    pendingOffice: report.samples.filter((sample) => sample.status === 'pending-office-baseline').length,
    failed: report.samples.filter((sample) => sample.status === 'failed').length,
    pixelIdentical: report.samples.filter((sample) => sample.status === 'pixel-identical').length };
  await writeFile(join(output, 'report.json'), JSON.stringify(report, null, 2) + '\n');
  await writeFile(join(output, 'index.html'), `<!doctype html><html lang=zh-CN><head><meta charset=utf-8><title>Mini Office Word 语料对照</title><style>body{font-family:Arial,sans-serif;margin:32px;background:#f6f7f9;color:#17202a}section{border-top:1px solid #ccd3da;padding:20px 0}.pair{display:flex;gap:20px;align-items:flex-start}figure{margin:0;flex:1;min-width:0}figcaption{margin-bottom:10px}img{max-width:100%;height:auto;border:1px solid #ccd3da}pre{white-space:pre-wrap}</style></head><body><h1>Word 公开语料逐页对照</h1><p>功能：展示浏览器回归与外部基准的独立状态；时间：2026-11-03；作者：lq。</p><p>差异值包含字体、抗锯齿和布局因素，不等于兼容百分比。未提供 Office 基准时，不作 Office 保真结论。</p>${report.fatalError ? `<pre>本次运行未完成：${escape(report.fatalError)}</pre>` : ''}${cards}</body></html>`);
}

try {
  await startServer();
  const bundled = chromium.executablePath();
  const local = join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing');
  let executablePath = process.env.MINI_OFFICE_CHROME;
  if (!executablePath) { try { await access(bundled); executablePath = bundled; } catch { executablePath = local; } }
  browser = await chromium.launch({ executablePath, headless: true });
  report.browser = browser.version();
  for (const fixture of fixtures) {
    const sample = { id: fixture.id, sourceSha256: fixture.sha256, status: 'pending-office-baseline', pages: [] };
    report.samples.push(sample);
    const page = await browser.newPage({ viewport: report.viewport, deviceScaleFactor: 1, colorScheme: 'light' });
    const errors = [];
    page.on('pageerror', (error) => errors.push(String(error)));
    try {
      const file = join(fixtureRoot, fixture.file);
      if (digest(await readFile(file)) !== fixture.sha256) throw new Error('语料 SHA-256 不匹配');
      await page.addInitScript(() => { window.showOpenFilePicker = undefined; });
      await page.goto(url, { waitUntil: 'domcontentloaded' });
      const chooser = page.waitForEvent('filechooser');
      await page.getByRole('button', { name: '打开文件', exact: true }).first().click();
      await (await chooser).setFiles(file);
      await page.waitForSelector('[data-word-editor]', { timeout: 30000 });
      sample.model = await page.evaluate(async ({ root }) => {
        // 功能：从应用共用的包入口读取当前文档状态，避免绕过入口创建独立 store 实例；时间：2026-11-03；作者：lq。
        const { useWorkspaceStore } = await import(`/@fs${root}/packages/document-model/src/index.ts`);
        const store = useWorkspaceStore.getState();
        const session = store.sessions[store.activeId];
        if (!session || session.kind !== 'docx') throw new Error('文档没有进入 Word 编辑器');
        store.updateView(session.id, { zoom: 1 });
        await document.fonts.ready;
        return { warnings: session.capability.warnings, dirty: session.dirty, page: session.data.page,
          requestedFonts: [...new Set([...JSON.stringify(session.data).matchAll(/"fontFamily":"([^"]+)"/g)].map((match) => match[1]))] };
      }, { root });
      // 功能：输出白纸且不改变正文布局，保持源码文字颜色并清除光标和页边距辅助线；时间：2026-11-03；作者：lq。
      await page.evaluate(() => { document.documentElement.dataset.theme = 'light'; });
      // 功能：格式标记只属于编辑辅助，零宽隐藏不改行盒，避免蓝色 Tab 箭头和换行符进入 Office 对照图；时间：2026-11-03；作者：lq。
      await page.addStyleTag({ content: '.mo-word-page{box-shadow:none!important;background:white!important;--mo-paper:#fff;--mo-paper-ink:#000} *{caret-color:transparent!important} [data-margin-guides],.mo-word-page [data-tab] [aria-hidden],.mo-word-page [data-line-break]{visibility:hidden!important}' });
      await page.evaluate(async () => {
        await document.fonts.ready;
        await Promise.all([...document.querySelectorAll('.mo-word-page img')].map((image) => image.decode()));
      });
      let previous = ''; let stable = 0;
      for (let attempt = 0; attempt < 60 && stable < 4; attempt++) {
        const current = await page.locator('.mo-word-page').evaluateAll((pages) => JSON.stringify(pages.map((item) => [item.offsetWidth, item.offsetHeight, item.innerHTML])));
        stable = current === previous ? stable + 1 : 0; previous = current; await delay(100);
      }
      if (stable < 4) throw new Error('分页布局未稳定');
      const pages = page.locator('.mo-word-page');
      sample.pageCount = await pages.count();
      if (!sample.pageCount) throw new Error('没有生成页面');
      // 功能：真实图片语料核对 18 幅图的显示尺寸、基线、效果占位及首幅图实际着色，防止图片框正确但内部图案被缩成方形；时间：2026-11-04；作者：lq。
      if (fixture.id === 'para-picture-only-line') {
        sample.imageGeometry = await pages.locator('[data-image-frame]').evaluateAll((frames) => frames.map((frame) => {
          const image = frame.querySelector('img');
          const paper = frame.closest('.mo-word-page');
          const scale = paper.getBoundingClientRect().width / paper.offsetWidth;
          const rect = image.getBoundingClientRect();
          const parent = frame.closest('[data-paragraph-id]');
          return { widthPx: rect.width / scale, heightPx: rect.height / scale,
            bottomPx: (rect.bottom - parent.getBoundingClientRect().top) / scale,
            paragraphHeightPx: parent.getBoundingClientRect().height / scale,
            effectBottomPx: Number.parseFloat(getComputedStyle(frame).marginBottom) };
        }));
        // 功能：在真实 DOM 中把两幅倍数行距图片限宽成两行，验证额外间距逐行生效而非只加在段尾；时间：2026-11-05；作者：lq。
        sample.imageAutoLines = await page.evaluate(() => {
          const source = [...document.querySelectorAll('.mo-word-page [data-paragraph-id]')].find((paragraph) =>
            paragraph.querySelectorAll('[data-image-frame]').length === 2
            && Number.parseFloat(getComputedStyle(paragraph).getPropertyValue('--mo-inline-image-line-extra')) > 0);
          if (!source) throw new Error('未找到倍数行距的图片段落');
          const clone = source.cloneNode(true);
          Object.assign(clone.style, { width: '200px', position: 'fixed', left: '0', top: '0', visibility: 'hidden' });
          document.body.append(clone);
          try {
            const frames = [...clone.querySelectorAll('[data-image-frame]')];
            const extra = Number.parseFloat(getComputedStyle(clone).getPropertyValue('--mo-inline-image-line-extra'));
            const images = frames.map((frame) => frame.querySelector('img').getBoundingClientRect());
            return { extraPx: extra, heightPx: clone.getBoundingClientRect().height,
              expectedHeightPx: images.reduce((height, image) => height + image.height + extra, 0),
              secondLineOffsetPx: images[1].top - images[0].top,
              expectedLineOffsetPx: images[0].height + extra };
          } finally { clone.remove(); }
        });
        if (Math.abs(sample.imageAutoLines.heightPx - sample.imageAutoLines.expectedHeightPx) > 0.1
          || Math.abs(sample.imageAutoLines.secondLineOffsetPx - sample.imageAutoLines.expectedLineOffsetPx) > 0.1) {
          throw new Error('图片自动换行后的逐行间距断言失败：' + JSON.stringify(sample.imageAutoLines));
        }
        const ink = await pages.locator('img').first().screenshot({ animations: 'disabled' });
        sample.imagePaint = await page.evaluate(async (encoded) => {
          const image = new Image(); image.src = `data:image/png;base64,${encoded}`; await image.decode();
          const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
          const context = canvas.getContext('2d'); context.drawImage(image, 0, 0);
          return [1, Math.floor(image.width / 2), image.width - 2].map((x) =>
            [...context.getImageData(x, Math.floor(image.height / 2), 1, 1).data]);
        }, ink.toString('base64'));
        const geometry = sample.imageGeometry;
        if (geometry.length !== 18 || geometry.some((item) => Math.abs(item.widthPx - 160) > 0.1)
          || geometry.filter((item) => Math.abs(item.heightPx - 80) < 0.1).length !== 17
          || Math.abs(geometry[7].heightPx - 160 / 3) > 0.1
          || Math.abs(geometry[6].bottomPx - geometry[7].bottomPx) > 0.1
          || Math.abs(geometry[0].paragraphHeightPx - 80) > 0.1
          || geometry.filter((item) => Math.abs(item.effectBottomPx - 32 / 3) < 0.1).length !== 2
          || sample.imagePaint.some((pixel) => pixel[0] !== 128 || pixel[1] !== 255 || pixel[2] !== 128 || pixel[3] !== 255)) {
          throw new Error('图片语料尺寸、基线、效果占位或实际绘制断言失败');
        }
      }
      // 功能：真实浏览器核对探针中的五个点线制表位，页码必须同线并对齐正文右边缘，避免仅凭成功截图判定布局正常；时间：2026-11-03；作者：lq。
      if (fixture.id === 'tab-stops-column-edge') {
        sample.tabGeometry = await pages.first().locator('[data-paragraph-id]').evaluateAll((paragraphs, setup) => paragraphs.slice(0, 5).map((paragraph) => {
          const paper = paragraph.closest('.mo-word-page').getBoundingClientRect();
          const scale = paper.width / setup.widthPx;
          const expectedRight = paper.right - setup.margins.right * scale;
          const tab = paragraph.querySelector('[data-tab]');
          const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
          const texts = [];
          let node;
          while ((node = walker.nextNode())) if (node.textContent.trim() && !node.parentElement.closest('[data-tab]')) texts.push(node);
          const last = document.createRange(); last.selectNodeContents(texts.at(-1));
          const first = document.createRange(); first.selectNodeContents(texts[0]);
          const end = last.getBoundingClientRect();
          return { leader: tab ? getComputedStyle(tab).borderBottomStyle : null,
            tabWidthPx: tab ? tab.getBoundingClientRect().width / scale : 0,
            rightGapPx: (expectedRight - end.right) / scale,
            lineOffsetPx: (end.top - first.getBoundingClientRect().top) / scale };
        }), sample.model.page);
      }
      let baseline;
      if (baselineRoot) {
        let data;
        try {
          data = await baselineFile(baselineRoot, `${fixture.id}/manifest.json`);
        } catch (error) { if (error.code !== 'ENOENT') throw error; }
        if (data) {
          baseline = validateOfficeBaseline(JSON.parse(data.toString('utf8')), fixture);
          const pdf = await baselineFile(baselineRoot, `${fixture.id}/${baseline.pdfFile}`);
          if (digest(pdf) !== baseline.pdfSha256) throw new Error('Office PDF 哈希不匹配');
          sample.baseline = baseline;
          sample.status = baseline.pages.length === sample.pageCount ? 'pixel-identical' : 'page-count-mismatch';
        }
      }
      for (let index = 0; index < sample.pageCount; index++) {
        const name = `${fixture.id}-page-${String(index + 1).padStart(3, '0')}.png`;
        const png = await pages.nth(index).screenshot({ path: join(output, name), animations: 'disabled', caret: 'hide' });
        const capture = { number: index + 1, file: name, sha256: digest(png) };
        sample.pages.push(capture);
        if (baseline?.pages[index]) {
          const expected = await baselineFile(baselineRoot, `${fixture.id}/${baseline.pages[index].file}`);
          if (digest(expected) !== baseline.pages[index].sha256) throw new Error('Office 基准 PNG 哈希不匹配');
          capture.baselineFile = `${fixture.id}-office-${String(index + 1).padStart(3, '0')}.png`;
          await writeFile(join(output, capture.baselineFile), expected);
          capture.comparison = await comparePng(page, png, expected);
          if (capture.comparison.status !== 'pixel-identical' && sample.status !== 'page-count-mismatch') sample.status = capture.comparison.status;
        }
      }
      if (errors.length) throw new Error(errors.join('\n'));
      if (sample.tabGeometry && (sample.tabGeometry.length !== 5 || sample.tabGeometry.some((item) =>
        item.leader !== 'dotted' || item.tabWidthPx <= 0 || Math.abs(item.rightGapPx) > 1.5 || Math.abs(item.lineOffsetPx) > 1.5))) {
        throw new Error('点线制表位的浏览器几何断言失败：' + JSON.stringify(sample.tabGeometry));
      }
    } catch (error) { sample.status = 'failed'; sample.error = String(error); }
    finally { await page.close(); }
    console.log(JSON.stringify({ id: sample.id, pages: sample.pageCount, status: sample.status, error: sample.error }));
  }
} catch (error) {
  report.fatalError = String(error);
  process.exitCode = 1;
  console.error(report.fatalError);
} finally {
  await browser?.close();
  if (server && server.exitCode === null) server.kill('SIGTERM');
  await writeReport();
}
if (report.samples.some((sample) => sample.status === 'failed' || (strict && sample.status !== 'pixel-identical'))) process.exitCode = 1;
console.log(`逐页对照报告：${join(output, 'index.html')}`);
