import { chromium } from 'playwright-core';
const CHROME = `${process.env.HOME}/Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing`;
const browser = await chromium.launch({ executablePath: CHROME, headless: true });
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
await page.goto('http://127.0.0.1:1420/', { waitUntil: 'domcontentloaded' });
await page.waitForTimeout(2500);
await page.getByRole('button', { name: /Excel|表格/ }).first().click();
await page.waitForSelector('.mo-sheet', { timeout: 20000 });
await page.waitForTimeout(800);
const c1 = page.locator('[data-cell="0:0"]').first();
const b1 = await c1.boundingBox();
const c2 = page.locator('[data-cell="4:0"]').first();
const b2 = await c2.boundingBox();
await page.mouse.move(b1.x + 10, b1.y + 8);
await page.mouse.down();
await page.mouse.move(b2.x + 10, b2.y + 8, { steps: 10 });
await page.mouse.up();
await page.waitForTimeout(400);
const info = await page.evaluate(() => {
  const all = Array.from(document.querySelectorAll('[data-range-box]'));
  const accent = Array.from(document.querySelectorAll('.mo-sheet div')).filter((el) => {
    const s = getComputedStyle(el);
    return s.position === 'absolute' && s.borderTopWidth === '2px' && s.pointerEvents === 'none';
  });
  return {
    rangeBoxes: all.map((el) => el.getAttribute('data-range-box')),
    accentCount: accent.length,
    accentHtml: accent.slice(0, 3).map((el) => el.outerHTML.slice(0, 220)),
    name: document.querySelector('[aria-label="名称框"]')?.value,
  };
});
console.log(JSON.stringify(info));
await browser.close();
