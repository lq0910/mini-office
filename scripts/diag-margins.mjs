/** 功能：验证「拖动页边距」与「打印标题」两处新交互的实际表现；时间：2026-11-01；作者：lq。 */
import { chromium } from 'playwright-core';
import fs from 'node:fs';
import path from 'node:path';

const CHROME = `${process.env.HOME}/Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing`;
const URL = 'http://127.0.0.1:1420/';
const OUT = path.join(process.cwd(), 'output/playwright');
fs.mkdirSync(OUT, { recursive: true });

const browser = await chromium.launch({ executablePath: CHROME, headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => { if (m.type() === 'error') errors.push(`console: ${m.text()}`); });
await page.goto(URL, { waitUntil: 'domcontentloaded' });
await page.waitForTimeout(2500);

// ---------- Excel 打印预览：页边距线 + 打印标题 ----------
await page.getByRole('button', { name: /Excel|表格/ }).first().click();
await page.waitForSelector('.mo-sheet', { timeout: 20000 });
await page.waitForTimeout(600);
const cellAt = (row, col) => page.locator(`[data-cell="${row}:${col}"]`).first();
for (let r = 0; r < 40; r += 1) {
  await cellAt(r, 0).click();
  await page.keyboard.type(r === 0 ? '表头' : `行${r}`);
  await page.keyboard.press('Enter');
}
await page.waitForTimeout(300);

await page.getByRole('button', { name: '打印', exact: true }).first().click();
await page.waitForSelector('[aria-label="工作表打印预览"]', { timeout: 10000 });
await page.waitForTimeout(900);
const marginLayer = await page.locator('[data-preview-margins]').count();
const bandBefore = await page.evaluate(() => {
  const band = document.querySelector('[data-preview-margins] > div');
  return band ? { left: band.style.left, width: band.style.width } : null;
});
console.log('预览页边距层数量:', marginLayer, '参考线:', JSON.stringify(bandBefore));
await page.screenshot({ path: path.join(OUT, 'diag-preview-margins.png') });

// 拖动左侧边距线
const leftStrip = page.locator('[data-preview-margins] [aria-label="拖动调整左页边距"]').first();
const box = await leftStrip.boundingBox();
if (box) {
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + 60, box.y + box.height / 2, { steps: 8 });
  await page.mouse.up();
  await page.waitForTimeout(900);
}
const bandAfter = await page.evaluate(() => {
  const band = document.querySelector('[data-preview-margins] > div');
  return band ? { left: band.style.left, width: band.style.width } : null;
});
console.log('拖动后参考线:', JSON.stringify(bandAfter));
await page.screenshot({ path: path.join(OUT, 'diag-preview-margins-dragged.png') });
await page.keyboard.press('Escape');
await page.waitForTimeout(400);

// 页面设置里填打印标题
await page.getByRole('button', { name: '页面设置' }).first().click();
await page.waitForTimeout(500);
const titleInput = page.locator('input[aria-label="打印标题行"]');
await titleInput.fill('1:1');
await page.getByRole('button', { name: '确定' }).click();
await page.waitForTimeout(600);
await page.getByRole('button', { name: '打印', exact: true }).first().click();
await page.waitForSelector('[aria-label="工作表打印预览"]', { timeout: 10000 });
await page.waitForTimeout(900);
const titleNote = await page.evaluate(() => document.body.innerText.match(/每页重复标题行[^\n]*/)?.[0] ?? null);
console.log('打印标题提示:', titleNote);
await page.screenshot({ path: path.join(OUT, 'diag-preview-title.png') });
await page.keyboard.press('Escape');
await page.waitForTimeout(300);

// ---------- Word：拖动页面边缘改页边距 ----------
// 功能：回到欢迎页再新建文档，避免工作区里找不到新建入口；时间：2026-11-01；作者：lq。
await page.reload({ waitUntil: 'domcontentloaded' });
await page.waitForTimeout(2200);
await page.getByRole('button', { name: /Word|文档/ }).first().click();
await page.waitForTimeout(1800);
const pageEl = page.locator('[data-word-page="1"]').first();
await page.waitForSelector('[data-word-page="1"]', { timeout: 10000 });
const para = page.locator('[data-word-page="1"] [data-paragraph-id]').first();
console.log('Word 段落数量:', await page.locator('[data-word-page="1"] [data-paragraph-id]').count());
await para.click();
await page.keyboard.type('Margin drag test');
await page.waitForTimeout(400);
const strip = page.locator('[aria-label="拖动调整左页边距"]').first();
const sbox = await strip.boundingBox();
console.log('Word 左页边距热区:', JSON.stringify(sbox));
if (sbox) {
  await page.mouse.move(sbox.x + sbox.width / 2, sbox.y + sbox.height / 2);
  await page.mouse.down();
  await page.mouse.move(sbox.x + 80, sbox.y + sbox.height / 2, { steps: 10 });
  await page.waitForTimeout(200);
  await page.screenshot({ path: path.join(OUT, 'diag-word-margin-drag.png') });
  await page.mouse.up();
  await page.waitForTimeout(700);
}
const guideInfo = await page.evaluate(() => {
  const guide = document.querySelector('[data-margin-guides]');
  return guide ? { left: guide.style.left, text: guide.textContent?.trim().slice(-60) } : null;
});
console.log('Word 页边距提示:', JSON.stringify(guideInfo));
const contentLeft = await page.evaluate(() => {
  const first = document.querySelector('[data-word-page="1"] [data-paragraph-id]');
  const pageEl = document.querySelector('[data-word-page="1"]');
  if (!first || !pageEl) return null;
  return Math.round(first.getBoundingClientRect().left - pageEl.getBoundingClientRect().left);
});
console.log('正文左边界偏移:', contentLeft);
await page.screenshot({ path: path.join(OUT, 'diag-word-margin-after.png') });
console.log('页面尺寸:', JSON.stringify(await pageEl.boundingBox()));
console.log('错误:', JSON.stringify(errors.slice(0, 8)));
await browser.close();
