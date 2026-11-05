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

// 新建表格
const sheetBtn = page.getByRole('button', { name: /Excel|表格/ }).first();
await sheetBtn.click();
await page.waitForSelector('.mo-sheet', { timeout: 20000 });
await page.waitForTimeout(800);

const cellAt = (row, col) => page.locator(`[data-cell="${row}:${col}"]`).first();

// 输入数字
for (let r = 0; r < 5; r += 1) {
  await cellAt(r, 0).click();
  await page.keyboard.type(String((r + 1) * 10));
  await page.keyboard.press('Enter');
}
await page.waitForTimeout(400);
const values = await page.evaluate(() =>
  Array.from(document.querySelectorAll('[data-cell$=":0"]')).slice(0, 6).map((el) => el.textContent),
);
console.log('A 列内容:', JSON.stringify(values));

// 拖选 A1:A5
const a1 = await cellAt(0, 0).boundingBox();
const a5 = await cellAt(4, 0).boundingBox();
await page.mouse.move(a1.x + a1.width / 2, a1.y + a1.height / 2);
await page.mouse.down();
await page.mouse.move(a1.x + a1.width / 2, a5.y + a5.height / 2, { steps: 12 });
await page.mouse.up();
await page.waitForTimeout(300);
const rangeLabel = await page.locator('[aria-label="名称框"]').inputValue();
console.log('拖选后名称框:', rangeLabel);
const status = await page.evaluate(() => {
  const bar = document.querySelector('.mo-sheet')?.parentElement?.querySelector('div');
  return document.body.innerText.match(/求和[^\n]*/g)?.slice(0, 2);
});
console.log('状态栏:', JSON.stringify(status));
await page.screenshot({ path: path.join(OUT, 'diag-sheet-drag.png') });

// 点击求和
await page.getByRole('button', { name: '自动求和' }).click();
await page.waitForTimeout(600);
const afterSum = await page.evaluate(() => ({
  editing: !!document.querySelector('input[aria-label="单元格编辑"]'),
  value: document.querySelector('input[aria-label="单元格编辑"]')?.value ?? null,
  formula: document.querySelector('input[aria-label="公式栏"]')?.value ?? null,
  reference: document.querySelector('[data-formula-reference]')?.getAttribute('data-formula-reference') ?? null,
}));
console.log('求和后:', JSON.stringify(afterSum));
await page.screenshot({ path: path.join(OUT, 'diag-sheet-sum.png') });
await page.keyboard.press('Escape');

// 填充柄拖动
await cellAt(0, 0).click();
await page.waitForTimeout(200);
const handle = await page.locator('[title="拖动填充"]').boundingBox();
console.log('填充柄:', JSON.stringify(handle));
if (handle) {
  const target = await cellAt(4, 0).boundingBox();
  await page.mouse.move(handle.x + 3, handle.y + 3);
  await page.mouse.down();
  await page.mouse.move(target.x + target.width / 2, target.y + target.height / 2, { steps: 10 });
  await page.mouse.up();
  await page.waitForTimeout(500);
  const filled = await page.evaluate(() =>
    Array.from(document.querySelectorAll('[data-cell$=":0"]')).slice(0, 6).map((el) => el.textContent),
  );
  console.log('填充后 A 列:', JSON.stringify(filled));
}
await page.screenshot({ path: path.join(OUT, 'diag-sheet-fill.png') });

console.log('错误:', JSON.stringify(errors.slice(0, 8)));
await browser.close();
