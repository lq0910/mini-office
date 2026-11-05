/** 功能：浏览器验证裁剪标记的版心定位、缩放、显隐与打印隔离；时间：2026-11-05；作者：lq。 */
import { chromium } from 'playwright-core';
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { homedir } from 'node:os';

const root = resolve(import.meta.dirname, '..');
const output = join(root, 'output/playwright/word-crop-marks');
await mkdir(output, { recursive: true });
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME
  ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const report = { generatedAt: new Date().toISOString(), checks: [] };
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1600 }, deviceScaleFactor: 1 });
  const errors = [];
  page.on('pageerror', (error) => errors.push(String(error)));
  await page.goto(process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1420/');
  // 功能：使用不对称边距的两节空白样本，检测标记是否错误复用末节或首页位置；时间：2026-11-05；作者：lq。
  let baseline = await page.evaluate(async (root) => {
    const { useWorkspaceStore } = await import(`/@fs${root}/packages/document-model/src/index.ts`);
    await useWorkspaceStore.getState().createDocument('docx');
    const state = useWorkspaceStore.getState();
    const session = state.sessions[state.activeId];
    const firstPage = { ...session.data.page, margins: { ...session.data.page.margins, left: 80, right: 100, top: 90, bottom: 70 } };
    const lastPage = { ...session.data.page, margins: { ...session.data.page.margins, left: 100, right: 80, top: 70, bottom: 90 } };
    const paragraph = session.data.blocks[0];
    state.updateData(session.id, { ...session.data, page: lastPage, blocks: [
      { ...paragraph, runs: [{ text: '裁剪标记 · 第一节', style: {} }], sectionEnd: { page: firstPage } },
      { ...paragraph, id: 'crop-second', runs: [{ text: '裁剪标记 · 第二节', style: {} }] },
    ] }, { dirty: false });
    state.updateSettings({ wordCropMarks: true });
    const current = useWorkspaceStore.getState().sessions[session.id];
    return { data: JSON.stringify(current.data), dirty: current.dirty, pages: [firstPage, lastPage] };
  }, root);
  await page.waitForFunction(() => document.querySelectorAll('[data-word-page] [data-crop-corner]').length === 8);
  // 功能：编辑器初始化会规范化模型，显隐前的稳定状态才是比较基准；时间：2026-11-05；作者：lq。
  await page.evaluate(() => document.fonts.ready);
  const stable = await page.evaluate(async (root) => {
    const { useWorkspaceStore } = await import(`/@fs${root}/packages/document-model/src/index.ts`);
    const state = useWorkspaceStore.getState(); const current = state.sessions[state.activeId];
    return { data: JSON.stringify(current.data), dirty: current.dirty };
  }, root);
  baseline = { ...baseline, ...stable };
  for (const zoom of [0.5, 1, 1.5]) {
    await page.evaluate(async ({ root, zoom }) => {
      const { useWorkspaceStore } = await import(`/@fs${root}/packages/document-model/src/index.ts`);
      const state = useWorkspaceStore.getState(); state.updateView(state.activeId, { zoom });
    }, { root, zoom });
    await page.waitForFunction((zoom) => document.querySelector('[data-word-page]').style.transform === `scale(${zoom})`, zoom);
    const checks = await page.locator('[data-word-page]').evaluateAll((papers, { zoom, setups }) => papers.map((paper, index) => {
      const rect = paper.getBoundingClientRect(); const setup = setups[index];
      return [...paper.querySelectorAll('[data-crop-corner]')].map((element) => {
        const r = element.getBoundingClientRect(); const key = element.dataset.cropCorner; const style = getComputedStyle(element);
        const x = (key.endsWith('l') ? r.right : r.left) - rect.left;
        const y = (key.startsWith('t') ? r.bottom : r.top) - rect.top;
        return { section: index + 1, key,
          errorX: Math.abs(x - (key.endsWith('l') ? setup.margins.left : setup.widthPx - setup.margins.right) * zoom),
          errorY: Math.abs(y - (key.startsWith('t') ? setup.margins.top : setup.heightPx - setup.margins.bottom) * zoom),
          horizontal: key.startsWith('t') ? style.borderBottomWidth : style.borderTopWidth,
          vertical: key.endsWith('l') ? style.borderRightWidth : style.borderLeftWidth,
          pointerEvents: getComputedStyle(element.parentElement).pointerEvents };
      });
    }), { zoom, setups: baseline.pages });
    for (const check of checks.flat()) {
      assert(check.errorX < 0.1 && check.errorY < 0.1, JSON.stringify(check));
      assert.equal(check.horizontal, '1px'); assert.equal(check.vertical, '1px'); assert.equal(check.pointerEvents, 'none');
    }
    report.checks.push({ zoom, corners: checks });
  }
  // 功能：切换只改变应用显示设置，不改变文档内容或脏标记；时间：2026-11-05；作者：lq。
  for (const visible of [false, true]) {
    await page.evaluate(async ({ root, visible }) => {
      const { useWorkspaceStore } = await import(`/@fs${root}/packages/document-model/src/index.ts`);
      const state = useWorkspaceStore.getState(); state.updateSettings({ wordCropMarks: visible }); state.updateView(state.activeId, { zoom: 1 });
    }, { root, visible });
    await page.waitForFunction((visible) => Boolean(document.querySelector('[data-word-crop-marks]')) === visible, visible);
  }
  const unchanged = await page.evaluate(async ({ root, baseline }) => {
    const { useWorkspaceStore } = await import(`/@fs${root}/packages/document-model/src/index.ts`);
    const state = useWorkspaceStore.getState(); const session = state.sessions[state.activeId];
    return JSON.stringify(session.data) === baseline.data && session.dirty === baseline.dirty;
  }, { root, baseline });
  assert(unchanged); report.toggleDoesNotEditDocument = true;
  await page.locator('[data-word-page]').first().screenshot({ path: join(output, 'editor.png') });
  await page.emulateMedia({ media: 'print' });
  assert.equal(await page.locator('[data-word-crop-marks]').first().evaluate((element) => getComputedStyle(element).display), 'none');
  report.hiddenInPrintMedia = true;
  await page.emulateMedia({ media: 'screen' });
  await page.evaluate(() => window.dispatchEvent(new CustomEvent('mo:print')));
  const preview = page.getByLabel('文档打印预览', { exact: true });
  await preview.locator('[data-word-page]').first().waitFor();
  assert.equal(await preview.locator('[data-word-crop-marks]').count(), 0);
  report.hiddenInPrintPreview = true;
  assert.deepEqual(errors, []); report.status = 'passed';
  await writeFile(join(output, 'report.json'), JSON.stringify(report, null, 2));
  console.log(JSON.stringify({ status: report.status, zooms: report.checks.map((check) => check.zoom), sections: 2,
    toggleDoesNotEditDocument: true, hiddenInPrintMedia: true, hiddenInPrintPreview: true }));
} finally { await browser.close(); }
