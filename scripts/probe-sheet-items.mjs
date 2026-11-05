/** 功能：核验填充柄数列、合并居中、边框内外粗细、冻结区选中框四项；时间：2026-11-01；作者：lq。 */
import { chromium } from 'playwright-core';

const CHROME = `${process.env.HOME}/Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing`;
const browser = await chromium.launch({ executablePath: CHROME, headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => { if (m.type() === 'error') errors.push(`console: ${m.text()}`); });
await page.goto('http://127.0.0.1:1420/', { waitUntil: 'domcontentloaded' });
await page.waitForTimeout(2500);
await page.getByRole('button', { name: /Excel|表格/ }).first().click();
await page.waitForSelector('.mo-sheet', { timeout: 20000 });
await page.waitForTimeout(700);

const cell = (r, c) => page.locator(`[data-cell="${r}:${c}"]`).first();
const box = async (r, c) => cell(r, c).boundingBox();
const dragSelect = async (r1, c1, r2, c2) => {
  const a = await box(r1, c1);
  const b = await box(r2, c2);
  await page.mouse.move(a.x + a.width / 2, a.y + a.height / 2);
  await page.mouse.down();
  await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2, { steps: 10 });
  await page.mouse.up();
  await page.waitForTimeout(200);
};
const typeInto = async (r, c, text) => {
  await cell(r, c).click();
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
  await page.waitForTimeout(80);
};

const report = {};

// 1) 填充柄：两格等差数列延展
await typeInto(0, 0, '10');
await typeInto(1, 0, '20');
await dragSelect(0, 0, 1, 0);
const handle = await page.locator('[title="拖动填充"]').boundingBox();
if (handle) {
  const target = await box(4, 0);
  await page.mouse.move(handle.x + 3, handle.y + 3);
  await page.mouse.down();
  await page.mouse.move(target.x + target.width / 2, target.y + target.height / 2, { steps: 10 });
  await page.mouse.up();
  await page.waitForTimeout(500);
}
report.填充柄两格 = await page.evaluate(() =>
  Array.from(document.querySelectorAll('[data-cell$=":0"]')).slice(0, 6).map((el) => el.textContent),
);

// 2) 合并居中（带文字，检查 flex 居中）
await typeInto(2, 1, '123');
await dragSelect(2, 1, 3, 2);
await page.getByRole('button', { name: '合并后居中' }).first().click();
await page.waitForTimeout(400);
report.合并居中 = await page.evaluate(() => {
  const merged = document.querySelector('[data-cell="2:1"]');
  if (!merged) return null;
  const s = getComputedStyle(merged);
  return { justifyContent: s.justifyContent, alignItems: s.alignItems, text: merged.textContent };
});

// 3) 边框内外粗细：选一个未合并区域应用粗外框
await dragSelect(6, 3, 8, 5);
await page.locator('button[aria-label="边框"]').first().click();
await page.waitForTimeout(250);
await page.getByRole('menuitem', { name: /粗外框线/ }).first().click();
await page.waitForTimeout(400);
report.边框 = await page.evaluate(() => {
  const read = (key) => {
    const el = document.querySelector(`[data-cell="${key}"]`);
    if (!el) return null;
    const s = getComputedStyle(el);
    return { top: s.borderTopWidth, left: s.borderLeftWidth, right: s.borderRightWidth, bottom: s.borderBottomWidth };
  };
  return { 外框左上角: read('6:3'), 内部中心: read('7:4'), 外框右下角: read('8:5') };
});

// 4) 冻结区选中框
await cell(2, 0).click();
await page.getByRole('button', { name: '冻结窗格' }).first().click();
await page.waitForTimeout(500);
await dragSelect(0, 0, 6, 1);
report.冻结选区 = await page.evaluate(() => ({
  冻结选区框: document.querySelectorAll('[data-range-box="selection"]').length,
  总选区框: document.querySelectorAll('[data-range-box]').length,
  名称框: document.querySelector('[aria-label="名称框"]')?.value,
}));
await page.screenshot({ path: 'output/playwright/audit-freeze.png' });

console.log(JSON.stringify(report, null, 2));
console.log('错误:', JSON.stringify(errors.slice(0, 5)));
await browser.close();
