/** 功能：核验冻结窗格滚动后选中框仍跟随冻结区；时间：2026-11-01；作者：lq。 */
import { chromium } from 'playwright-core';

const CHROME = `${process.env.HOME}/Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing`;
const browser = await chromium.launch({ executablePath: CHROME, headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
await page.goto('http://127.0.0.1:1420/', { waitUntil: 'domcontentloaded' });
await page.waitForTimeout(2500);
await page.getByRole('button', { name: /Excel|表格/ }).first().click();
await page.waitForSelector('.mo-sheet', { timeout: 20000 });
await page.waitForTimeout(700);

const cell = (r, c) => page.locator(`[data-cell="${r}:${c}"]`).first();

// 冻结第 1-2 行（活动单元格放在 A3）
await cell(2, 0).click();
await page.getByRole('button', { name: '冻结窗格' }).first().click();
await page.waitForTimeout(400);

// 选中 A1:B6，跨越冻结边界
const a = await cell(0, 0).boundingBox();
const b = await cell(5, 1).boundingBox();
await page.mouse.move(a.x + a.width / 2, a.y + a.height / 2);
await page.mouse.down();
await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2, { steps: 10 });
await page.mouse.up();
await page.waitForTimeout(300);

const before = await page.evaluate(() => {
  const box = document.querySelector('[data-range-box="selection"]');
  const rect = box?.getBoundingClientRect();
  return { count: document.querySelectorAll('[data-range-box="selection"]').length, top: rect ? Math.round(rect.top) : null, height: rect ? Math.round(rect.height) : null };
});

// 向下滚动 400px，冻结区应保持不动、选中框仍贴住冻结区
await page.mouse.move(700, 500);
await page.mouse.wheel(0, 400);
await page.waitForTimeout(500);
const after = await page.evaluate(() => {
  const box = document.querySelector('[data-range-box="selection"]');
  const rect = box?.getBoundingClientRect();
  const firstRow = document.querySelector('[data-cell="0:0"]')?.getBoundingClientRect();
  return {
    count: document.querySelectorAll('[data-range-box="selection"]').length,
    top: rect ? Math.round(rect.top) : null,
    height: rect ? Math.round(rect.height) : null,
    firstRowTop: firstRow ? Math.round(firstRow.top) : null,
  };
});
await page.screenshot({ path: 'output/playwright/audit-freeze-scrolled.png' });
console.log(JSON.stringify({ 滚动前: before, 滚动后: after }, null, 2));
await browser.close();
