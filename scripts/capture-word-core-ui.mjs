/** 功能：以公共夹具记录 Word 核心迁移前的本项目 UI，不执行正文编辑或保存；时间：2026-11-08；作者：lq。 */
import { chromium } from 'playwright-core';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { homedir } from 'node:os';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const output = join(root, 'output/playwright/word-core-m0');
const sample = join(root, 'tests/fixtures/docx/contract.docx');
const manifestPath = join(output, 'manifest.json');
if (existsSync(manifestPath)) throw new Error('UI 基线已存在，拒绝覆盖');
mkdirSync(output, { recursive: true });
const viewport = { width: 1440, height: 1000 };
const errors = [];
const screenshots = [];
/** 功能：生成样本和截图哈希，供不可变基线校验；时间：2026-11-08；作者：lq。 */
const sha = path => createHash('sha256').update(readFileSync(path)).digest('hex');
const sourceHash = sha(sample);
const baseline = JSON.parse(readFileSync(join(root, 'docs/word-core-m0/workspace.json'), 'utf8'));
const currentCommit = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim();
if (currentCommit !== baseline.commit) throw new Error('UI 截图与工作区基线提交不同');
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
try {
  const page = await browser.newPage({ viewport, deviceScaleFactor: 1 });
  page.setDefaultTimeout(15000);
  page.on('pageerror', error => errors.push(String(error)));
  await page.addInitScript(() => { window.showOpenFilePicker = undefined; window.showSaveFilePicker = undefined; });
  await page.goto(process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1420/', { waitUntil: 'domcontentloaded' });
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: '打开文件', exact: true }).first().click();
  await (await chooser).setFiles(sample);
  await page.waitForSelector('[data-word-editor]', { timeout: 60000 });
  await page.evaluate(() => document.fonts.ready);
  const toolbar = page.locator('[aria-label="Word 功能区"]');
  await toolbar.waitFor({ state: 'visible' });
  /** 功能：截图前等待界面稳定，同时记录可见控件，便于后续按行为和视觉对照；时间：2026-11-08；作者：lq。 */
  const capture = async (name, description) => {
    await page.waitForTimeout(200);
    const path = join(output, `${name}.png`);
    await page.screenshot({ path });
    const controls = await toolbar.locator('button,select,input,[role="combobox"]').evaluateAll(elements => elements.filter(element => element.getClientRects().length).map(element => ({ role: element.getAttribute('role') ?? element.tagName.toLowerCase(), name: element.getAttribute('aria-label') ?? element.getAttribute('title') ?? element.textContent?.trim() ?? '', disabled: element.disabled === true })));
    screenshots.push({ name, description, path: relative(root, path), sha256: sha(path), viewport: page.viewportSize(), controls });
  };
  // 功能：主题只作用于本次隔离浏览器，不改用户设置或项目源码；时间：2026-11-08；作者：lq。
  await page.evaluate(() => { document.documentElement.dataset.theme = 'light'; });
  for (const [index, tab] of ['开始', '插入', '布局', '引用', '邮件', '视图'].entries()) {
    await toolbar.getByRole('tab', { name: tab, exact: true }).click();
    await capture(`ribbon-${index + 1}`, `浅色主题 Word 功能区：${tab}`);
  }
  await toolbar.getByRole('tab', { name: '开始', exact: true }).click();
  await toolbar.getByRole('button', { name: '更多操作', exact: true }).click();
  await page.getByRole('menuitem', { name: /布局与属性/ }).click();
  await capture('properties', '现有布局与属性面板');
  await toolbar.getByRole('button', { name: '更多操作', exact: true }).click();
  await page.getByRole('menuitem', { name: /隐藏属性面板/ }).click();
  const cell = page.locator('[data-word-editor] [data-cell-id]').first();
  if (!await cell.count()) throw new Error('公共合同夹具未显示表格，缺少上下文基线');
  await cell.click({ button: 'right' });
  await page.getByRole('menu').waitFor({ state: 'visible' });
  await capture('table-context', '现有表格右键菜单');
  await page.keyboard.press('Escape');
  await page.evaluate(() => { document.documentElement.dataset.theme = 'dark'; });
  await capture('dark', '深色主题 Word 界面');
  await page.setViewportSize({ width: 900, height: 700 });
  await capture('narrow', '窄窗口功能区布局');
  if (errors.length) throw new Error(errors.join('\n'));
  if (sha(sample) !== sourceHash) throw new Error('公共样本字节变化');
  const manifest = { _comment: '功能：记录公共夹具下的 Word UI 迁移基线；时间：2026-11-08；作者：lq。', schemaVersion: 1, workspaceCommit: baseline.commit, browser: browser.version(), environment: 'Chromium 浏览器，非 Tauri WebView 验收', capturedAt: new Date().toISOString(), sample: { path: relative(root, sample), sha256: sourceHash }, sourceUnchanged: true, errors, screenshots };
  writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + '\n');
  console.log(JSON.stringify({ status: 'captured', screenshots: screenshots.length, sampleUnchanged: true, manifest: relative(root, manifestPath) }));
} finally {
  await browser.close();
}
