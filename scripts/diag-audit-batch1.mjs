/**
 * 功能：回归核验第一批用户反馈（顶栏/打印预览/表格交互/颜色/对齐）的实际表现；
 * 时间：2026-11-01；作者：lq。
 */
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

const report = {};

// ---------------- Excel ----------------
await page.getByRole('button', { name: /Excel|表格/ }).first().click();
await page.waitForSelector('.mo-sheet', { timeout: 20000 });
await page.waitForTimeout(700);
const cellAt = (row, col) => page.locator(`[data-cell="${row}:${col}"]`).first();

for (let r = 0; r < 5; r += 1) {
  await cellAt(r, 0).click();
  await page.keyboard.type(String((r + 1) * 10));
  await page.keyboard.press('Enter');
}
await page.waitForTimeout(300);

// 拖选
const a1 = await cellAt(0, 0).boundingBox();
const a5 = await cellAt(4, 0).boundingBox();
await page.mouse.move(a1.x + a1.width / 2, a1.y + a1.height / 2);
await page.mouse.down();
await page.mouse.move(a1.x + a1.width / 2, a5.y + a5.height / 2, { steps: 12 });
await page.mouse.up();
await page.waitForTimeout(250);
report.拖选 = {
  名称框: await page.locator('[aria-label="名称框"]').inputValue(),
  选区框数量: await page.locator('[data-range-box="selection"]').count(),
  状态栏求和: await page.evaluate(() => document.body.innerText.match(/求和[^\n]{0,12}/)?.[0] ?? null),
};

// 自动求和 + 引用高亮
await page.getByRole('button', { name: '自动求和' }).click();
await page.waitForTimeout(500);
report.自动求和 = {
  公式栏: await page.locator('[aria-label="公式栏"]').inputValue(),
  引用高亮: await page.locator('[data-editing-reference]').count(),
};
await page.keyboard.press('Escape');
await page.waitForTimeout(200);

// 填充柄
await cellAt(0, 0).click();
await page.waitForTimeout(200);
const handle = await page.locator('[title="拖动填充"]').boundingBox();
if (handle) {
  const target = await cellAt(4, 0).boundingBox();
  await page.mouse.move(handle.x + 3, handle.y + 3);
  await page.mouse.down();
  await page.mouse.move(target.x + target.width / 2, target.y + target.height / 2, { steps: 10 });
  await page.mouse.up();
  await page.waitForTimeout(500);
}
report.填充柄 = await page.evaluate(() =>
  Array.from(document.querySelectorAll('[data-cell$=":0"]')).slice(0, 5).map((el) => el.textContent),
);

// 合并居中
await cellAt(1, 1).click();
await page.mouse.move((await cellAt(1, 1).boundingBox()).x + 10, (await cellAt(1, 1).boundingBox()).y + 8);
const b2 = await cellAt(1, 1).boundingBox();
const c3 = await cellAt(2, 2).boundingBox();
await page.mouse.down();
await page.mouse.move(c3.x + c3.width / 2, c3.y + c3.height / 2, { steps: 8 });
await page.mouse.up();
await page.waitForTimeout(200);
await page.getByRole('button', { name: '合并后居中' }).first().click();
await page.waitForTimeout(400);
report.合并居中 = await page.evaluate(() => {
  const merged = document.querySelector('[data-cell="1:1"]');
  const swallowed = document.querySelector('[data-cell="1:2"]');
  return {
    合并后左上单元格: !!merged,
    被合并单元格已隐藏: !swallowed,
    水平居中: merged ? getComputedStyle(merged).textAlign : null,
    垂直居中: merged ? getComputedStyle(merged).alignItems : null,
  };
});

// 边框粗细（外框粗线）
await page.locator('button[aria-label="边框"]').first().click();
await page.waitForTimeout(300);
const thick = page.getByRole('menuitem', { name: /粗外框线/ }).first();
if (await thick.count()) { await thick.click(); } else { await page.keyboard.press('Escape'); }
await page.waitForTimeout(400);
report.边框 = await page.evaluate(() => {
  const outer = document.querySelector('[data-cell="1:1"]');
  const inner = document.querySelector('[data-cell="2:2"]');
  const width = (el, side) => (el ? getComputedStyle(el)[`border${side}Width`] : null);
  return { 外框上边: width(outer, 'Top'), 内部上边: width(inner, 'Top') };
});
await page.screenshot({ path: path.join(OUT, 'audit-sheet.png') });

// 冻结 + 选区框
await cellAt(8, 0).click();
await page.getByRole('button', { name: '冻结窗格' }).first().click();
await page.waitForTimeout(500);
report.冻结 = { 选区框数量: await page.locator('[data-range-box="selection"]').count() };

// 打印预览缩放选项
await page.getByRole('button', { name: '打印', exact: true }).first().click();
await page.waitForSelector('[aria-label="工作表打印预览"]', { timeout: 10000 });
await page.waitForTimeout(700);
report.打印预览 = await page.evaluate(() => {
  const selects = Array.from(document.querySelectorAll('select'));
  const zoom = selects.find((s) => Array.from(s.options).some((o) => o.textContent?.includes('将所有列打印在一页')));
  const host = document.querySelector('[aria-label="工作表打印预览"]');
  return {
    缩放选项: zoom ? Array.from(zoom.options).map((o) => o.textContent) : null,
    预览高度: host ? Math.round(host.getBoundingClientRect().height) : null,
    页边距参考线: document.querySelectorAll('[data-preview-margins]').length,
  };
});
await page.keyboard.press('Escape');
await page.waitForTimeout(300);

// 颜色记忆：先选一次色，再点主按钮应直接套用
await page.locator('[aria-label="字体颜色调色板"]').first().click();
await page.waitForTimeout(250);
await page.locator('[aria-label="#D92D20"]').first().click();
await page.waitForTimeout(250);
report.颜色记忆 = await page.evaluate(() => {
  const swatch = document.querySelector('[aria-label="字体颜色"] [aria-hidden="true"] > span:last-child');
  return { 当前色条: swatch ? getComputedStyle(swatch).backgroundColor : null };
});

// ---------------- Word ----------------
await page.reload({ waitUntil: 'domcontentloaded' });
await page.waitForTimeout(2200);
await page.getByRole('button', { name: /Word|文档/ }).first().click();
await page.waitForTimeout(1600);
await page.locator('[data-word-page="1"] [data-paragraph-id]').first().click();
await page.keyboard.type('alignment test');
await page.keyboard.press('Meta+e');
await page.waitForTimeout(400);
report.Word对齐 = await page.evaluate(() => {
  const p = document.querySelector('[data-word-page="1"] [data-paragraph-id]');
  return p ? { textAlign: getComputedStyle(p).textAlign, align: p.getAttribute('data-align') } : null;
});
await page.screenshot({ path: path.join(OUT, 'audit-word.png') });

console.log(JSON.stringify(report, null, 2));
console.log('错误:', JSON.stringify(errors.slice(0, 6)));
await browser.close();
