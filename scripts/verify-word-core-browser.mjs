/** 功能：在独立浏览器页验证 M1 构建产物、真实输入法和 Word 保存，不切换产品入口；时间：2026-11-08；作者：lq。 */
import { chromium } from 'playwright-core';
import JSZip from 'jszip';
import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { homedir } from 'node:os';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const output = join(root, 'output/playwright/word-core-m1.local');
await mkdir(output, { recursive: true });
const zip = await JSZip.loadAsync(await readFile(join(root, 'tests/fixtures/docx/contract.docx')));
const continuation = '<w:tc><w:tcPr><w:vMerge/><w:tcBorders><w:bottom w:val="double" w:sz="8" w:color="FF0000"/></w:tcBorders><w:shd w:fill="DDDDDD"/></w:tcPr><w:p/></w:tc>';
zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p w:rsidR="KEEP"><w:bookmarkStart w:id="1" w:name="keep"/><w:r><w:rPr><w:lang w:val="zh-CN"/></w:rPr><w:t>建设内容</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p></w:tc></w:tr><w:tr>${continuation}</w:tr></w:tbl><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
const input = await zip.generateAsync({ type: 'uint8array' });
const origin = process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1420';
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const errors = [];
const report = { checks: [], status: 'running' };
try {
  const page = await browser.newPage({ viewport: { width: 1100, height: 800 } });
  page.setDefaultTimeout(15000);
  page.on('pageerror', error => errors.push(String(error)));
  await page.route('**/__word_core_m1__', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"></head><body><div id="core" class="mo-word-core"><div id="editor"></div></div></body></html>' }));
  await page.goto(`${origin}/__word_core_m1__`);
  const base = `${origin}/@fs${join(root, 'packages/word-core/dist')}`;
  await page.addStyleTag({ url: `${base}/styles.css` });
  await page.addStyleTag({ url: `${base}/fonts.css` });
  const opened = await page.evaluate(async ({ url, bytes }) => {
    const core = await import(url);
    // 功能：使用构建出的真实 module Worker，验证资源路径及可转移来源字节；时间：2026-11-08；作者：lq。
    core.configureWordPlatform({ createParseWorker: () => new Worker(url.replace(/index\.js$/, 'parse-worker.js'), { type: 'module' }) });
    const session = await core.WordCoreSession.open(new Uint8Array(bytes), { element: document.getElementById('editor') });
    window.__wordSession = session;
    const css = document.createElement('style');
    css.textContent = await session.activate();
    document.head.append(css);
    const scope = document.getElementById('core');
    scope.style.cssText = 'width:794px;margin:24px auto;background:white;color:black';
    session.editor.view.dom.style.cssText = 'width:600px;min-height:300px;padding:30px;';
    const noOp = await session.prepareSave();
    window.__wordCore = core;
    return { bytes: Array.from(noOp.bytes), platform: core.wordPlatformCapabilities(), desktopBridgePresent: 'desktop' in window, parseStats: { ...core.parseStats } };
  }, { url: `${base}/index.js`, bytes: Array.from(input) });
  assert.deepEqual(new Uint8Array(opened.bytes), input);
  assert.equal(opened.desktopBridgePresent, false);
  assert.deepEqual(opened.platform, { fontMetrics: false, copyImageToClipboard: false, parseWorker: true });
  assert.equal(opened.parseStats.worker, 1);
  assert.equal(opened.parseStats.inline, 0);
  report.checks.push('built-module-without-react-electron-or-desktop-global', 'real-parse-worker-and-transferable-source', 'exact-noop-roundtrip');
  await page.evaluate(() => {
    const editor = window.__wordSession.editor;
    let from = -1;
    editor.state.doc.descendants((node, pos) => { if (from < 0 && node.isText) from = pos; });
    editor.commands.setTextSelection({ from, to: from + '建设内容'.length });
    editor.view.focus();
  });
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Input.imeSetComposition', { text: 'zhi hui wu ye', selectionStart: 12, selectionEnd: 12 });
  const composing = await page.evaluate(async () => {
    try { await window.__wordSession.prepareSave(); return ''; } catch (error) { return String(error); }
  });
  assert.match(composing, /输入尚未确认/);
  const chinese = '智慧物业信息系统部署实施';
  await cdp.send('Input.insertText', { text: chinese });
  await page.waitForFunction(chinese => !window.__wordSession.editor.view.composing && window.__wordSession.editor.getText().includes(chinese), chinese);
  const saved = await page.evaluate(async () => {
    const session = window.__wordSession;
    const prepared = await session.prepareSave();
    const dirtyBeforeAck = session.state().dirty;
    session.acknowledgeSave(prepared);
    return { bytes: Array.from(prepared.bytes), dirtyBeforeAck, dirtyAfterAck: session.state().dirty };
  });
  assert.equal(saved.dirtyBeforeAck, true);
  assert.equal(saved.dirtyAfterAck, false);
  const bytes = new Uint8Array(saved.bytes);
  const savedZip = await JSZip.loadAsync(bytes);
  const xml = await savedZip.file('word/document.xml').async('text');
  for (const value of [chinese, continuation, 'w:rsidR="KEEP"', '<w:bookmarkStart w:id="1" w:name="keep"/>']) assert.ok(xml.includes(value), value);
  await writeFile(join(output, 'saved.docx'), bytes);
  await page.screenshot({ path: join(output, 'core.png') });
  report.checks.push('native-composition-save-guard', 'committed-chinese-save', 'continuation-and-bookmark-preservation', 'explicit-save-ack');
  await page.evaluate(() => { window.__wordSession.dispose(); window.__wordCore.disposeParseWorker(); });
  assert.deepEqual(errors, []);
  report.browser = browser.version();
  report.status = 'passed';
} catch (error) {
  report.status = 'failed'; report.error = String(error); process.exitCode = 1;
} finally {
  await browser.close();
  await writeFile(join(output, 'report.json'), JSON.stringify({ ...report, errors }, null, 2));
  console.log(JSON.stringify({ ...report, errors }));
}
