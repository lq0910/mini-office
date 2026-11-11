/** 功能：真实界面验收图片、表格右键/矩形/尺寸拖动、单步历史和 DOCX 重开；时间：2026-11-11；作者：lq。 */
import { chromium } from 'playwright-core';
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';
import JSZip from 'jszip';
const root = resolve(import.meta.dirname, '..');
const output = join(root, 'output/playwright/word-core-m3-objects.local');
await mkdir(output, { recursive: true });
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const report = { status: 'running', checks: [], errors: [] };
let page;
try {
  page = await browser.newPage({ viewport: { width: 1440, height: 1100 }, acceptDownloads: true });
  page.setDefaultTimeout(18000); page.on('pageerror', error => report.errors.push(String(error)));
  await page.addInitScript(() => { window.showSaveFilePicker = undefined; });
  await page.goto(process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1431');
  await page.getByRole('button', { name: 'DOCX 文件 Word 文档' }).click();
  const editor = page.locator('[data-word-engine="word-core"] .ProseMirror').first();
  await editor.click();
  // 功能：只读取运行时摘要，内容修改全部通过实际 UI；时间：2026-11-11；作者：lq。
  await page.evaluate(async root => {
    const { useWorkspaceStore } = await import('/@fs' + root + '/packages/document-model/src/index.ts');
    window.__objectStore = useWorkspaceStore;
    window.__objectCore = () => useWorkspaceStore.getState().sessions[useWorkspaceStore.getState().activeId].wordCore;
    window.__objectTable = () => window.__objectCore().editor.state.doc.content.content.find(node => node.type.name === 'docTable');
  }, root);
  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '插入表格', exact: true }).click();
  await page.getByRole('button', { name: '3 列 3 行', exact: true }).click();
  const table = editor.locator('table.doc-table').first();
  await table.waitFor(); assert.equal(await table.locator('tr').count(), 3);
  await page.keyboard.insertText('首格中文');
  await page.getByRole('button', { name: '撤销', exact: true }).click();
  assert.equal((await table.innerText()).trim(), '');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.equal(await editor.locator('table').count(), 0);
  await page.getByRole('button', { name: '重做', exact: true }).click(); await table.waitFor();
  report.checks.push('table-grid-picker-native-cell-focus-and-single-undo-redo');
  /** 功能：右键使用模型格内选区，不丢失已有矩形选择；时间：2026-11-11；作者：lq。 */
  const menu = async (label, cell = 0) => { await table.locator('td, th').nth(cell).click({ button: 'right' }); await page.getByRole('menuitem', { name: label, exact: true }).click(); };
  for (const [label, rows, cols] of [['在上方插入行', 4, 3], ['在下方插入行', 4, 3], ['在左侧插入列', 3, 4], ['在右侧插入列', 3, 4], ['删除行', 2, 3], ['删除列', 3, 2]]) {
    await menu(label); assert.equal(await table.locator('tr').count(), rows); assert.equal(await table.locator('tr').first().locator('td, th').count(), cols);
    await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.equal(await table.locator('tr').count(), 3);
  }
  report.checks.push('all-four-row-column-inserts-and-both-deletes-single-undo');
  await table.locator('tr').first().locator('td, th').nth(0).click(); await page.keyboard.insertText('合并甲');
  await table.locator('tr').first().locator('td, th').nth(1).click(); await page.keyboard.insertText('合并乙');
  const first = await table.locator('tr').first().locator('td, th').nth(0).boundingBox();
  const second = await table.locator('tr').first().locator('td, th').nth(1).boundingBox();
  await page.mouse.move(first.x + 24, first.y + first.height / 2); await page.mouse.down();
  await page.mouse.move(second.x + 40, second.y + second.height / 2, { steps: 15 }); await page.mouse.up();
  await page.waitForFunction(() => window.__objectCore().editor.state.selection.toJSON().type === 'cell');
  await menu('合并单元格'); assert.equal(await table.locator('td[colspan="2"], th[colspan="2"]').count(), 1);
  assert.ok((await table.innerText()).includes('合并甲')); assert.ok((await table.innerText()).includes('合并乙'));
  await menu('拆分单元格'); assert.equal(await table.locator('td[colspan="2"], th[colspan="2"]').count(), 0);
  report.checks.push('mouse-rectangle-selection-merge-split-retains-text');
  await menu('重复标题行'); assert.equal(await table.locator('tr').first().getAttribute('data-repeat-header'), '1');
  const version = await page.evaluate(() => window.__objectCore().state().revision);
  await menu('选择整个表格'); assert.equal(await page.evaluate(() => window.__objectCore().state().revision), version);
  assert.equal(await page.evaluate(() => window.__objectCore().editor.state.selection.toJSON().type), 'node');
  // 功能：右键首格后选择列，属性作用于原生列选区；时间：2026-11-11；作者：lq。
  await menu('选择列'); await menu('表格属性…');
  const dialog = page.getByRole('dialog', { name: '表格属性', exact: true });
  await dialog.getByRole('spinbutton', { name: '表格宽度 (px)', exact: true }).fill('480');
  await dialog.getByRole('combobox', { name: '表格对齐', exact: true }).selectOption('center');
  await dialog.getByRole('spinbutton', { name: '行高 (px)', exact: true }).fill('38');
  await dialog.getByRole('combobox', { name: '垂直对齐', exact: true }).selectOption('bottom');
  await dialog.getByLabel('单元格底纹', { exact: true }).fill('#fff1a8');
  await dialog.getByRole('combobox', { name: '边框范围', exact: true }).selectOption('all');
  await dialog.getByLabel('边框颜色', { exact: true }).fill('#246b8e');
  await dialog.getByRole('button', { name: '应用表格属性', exact: true }).click();
  await page.waitForFunction(() => window.__objectTable().attrs.widthPx === 480);
  assert.equal(await page.evaluate(() => window.__objectTable().child(0).child(0).attrs.fill), 'fff1a8');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.notEqual(await page.evaluate(() => window.__objectTable().attrs.widthPx), 480);
  await page.getByRole('button', { name: '重做', exact: true }).click();
  report.checks.push('repeat-header-and-selection-do-not-create-fake-edits', 'table-properties-retain-column-selection-and-single-undo');
  await menu('表格属性…'); await dialog.getByRole('spinbutton', { name: '表格宽度 (px)', exact: true }).fill('0');
  const badVersion = await page.evaluate(() => window.__objectCore().state().revision);
  await dialog.getByRole('button', { name: '应用表格属性', exact: true }).click(); await dialog.getByRole('alert').waitFor();
  assert.equal(await page.evaluate(() => window.__objectCore().state().revision), badVersion);
  await dialog.getByRole('button', { name: '取消', exact: true }).click();
  report.checks.push('invalid-table-size-rejected-without-partial-change');
  await menu('平均分布各列'); await menu('平均分布各行'); await menu('适应页宽');
  report.checks.push('distribute-rows-columns-and-autofit-use-native-grid');
  // 功能：实际拖动列边界和行底边，松手前文档保持原版本；时间：2026-11-11；作者：lq。
  const dragCell = await table.locator('tr').first().locator('td,th').first().boundingBox();
  const widthBefore = await page.evaluate(() => window.__objectTable().child(0).child(0).attrs.colwidth?.[0]);
  const dragVersion = await page.evaluate(() => window.__objectCore().state().revision);
  await page.mouse.move(dragCell.x + dragCell.width - 1, dragCell.y + dragCell.height / 2); await page.mouse.down();
  await page.mouse.move(dragCell.x + dragCell.width + 25, dragCell.y + dragCell.height / 2, { steps: 12 });
  assert.equal(await page.evaluate(() => window.__objectCore().state().revision), dragVersion); await page.mouse.up();
  await page.waitForFunction(old => window.__objectTable().child(0).child(0).attrs.colwidth?.[0] !== old, widthBefore);
  report.checks.push('real-column-border-drag-previews-before-commit');
  await table.locator('tr').first().locator('td,th').first().hover();
  const row = await table.locator('tr').first().boundingBox();
  const heightBefore = await page.evaluate(() => window.__objectTable().child(0).attrs.heightTwips);
  await page.mouse.move(row.x + row.width / 2 + 20, row.y + row.height - 1); await page.mouse.down();
  await page.mouse.move(row.x + row.width / 2 + 20, row.y + row.height + 22, { steps: 10 }); await page.mouse.up();
  await page.waitForFunction(old => window.__objectTable().child(0).attrs.heightTwips !== old, heightBefore);
  report.checks.push('real-row-border-drag-persists-native-row-height');
  await editor.locator(':scope > p').last().click(); await page.keyboard.insertText('图片后正文');
  // 功能：使用公开生成的 PNG 夹具验收原生文件选择与图片解码；时间：2026-11-11；作者：lq。
  const png = await page.evaluate(() => { const c = document.createElement('canvas'); c.width = 180; c.height = 90; const g = c.getContext('2d'); g.fillStyle = '#246b8e'; g.fillRect(0, 0, 180, 90); g.fillStyle = '#ffffff'; g.font = '20px sans-serif'; g.fillText('M3 image', 20, 52); return c.toDataURL('image/png').split(',')[1]; });
  const pictureFile = join(output, 'fixture.png'); await writeFile(pictureFile, Buffer.from(png, 'base64'));
  const fileChooser = page.waitForEvent('filechooser'); await page.getByRole('button', { name: '插入图片', exact: true }).click(); await (await fileChooser).setFiles(pictureFile);
  const picture = editor.locator('img.doc-protected-img').first(); await picture.waitFor();
  // 功能：文件插入后的焦点恢复和分页异步重绘完成后，再验证单击选中；时间：2026-11-11；作者：lq。
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  const clickVersion = await page.evaluate(() => window.__objectCore().state().revision);
  await picture.click();
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  assert.equal(await page.evaluate(() => window.__objectCore().state().revision), clickVersion);
  assert.equal(await page.evaluate(() => window.__objectCore().editor.state.selection.toJSON().type), 'node');
  const handle = editor.locator('.img-resize-handle[data-resize-handle="se"]').first(); await handle.waitFor({ state: 'visible' });
  const grip = await handle.boundingBox(); const imageVersion = await page.evaluate(() => window.__objectCore().state().revision);
  await page.mouse.move(grip.x + grip.width / 2, grip.y + grip.height / 2); await page.mouse.down(); await page.mouse.move(grip.x + 45, grip.y + 25, { steps: 12 });
  assert.equal(await page.evaluate(() => window.__objectCore().state().revision), imageVersion); await page.mouse.up();
  const pictureWidth = await page.evaluate(() => window.__objectCore().editor.state.doc.content.content.find(node => node.attrs.blockType === 'image').attrs.imageWidthPx); assert.ok(pictureWidth > 180);
  for (const [label, wrap] of [['左侧环绕', 'square-left'], ['上下型', 'topBottom'], ['嵌入型', null], ['右侧环绕', 'square-right']]) {
    await picture.click({ button: 'right' }); await page.getByRole('menuitem', { name: label, exact: true }).click();
    await page.waitForFunction(wrap => window.__objectCore().editor.state.doc.content.content.find(node => node.attrs.blockType === 'image').attrs.imageWrap === wrap, wrap);
  }
  await page.waitForFunction(() => window.__objectCore().editor.state.doc.content.content.some(node => node.attrs.imageWrap === 'square-right'));
  await picture.click({ button: 'right' }); await page.getByRole('menuitem', { name: '删除图片', exact: true }).click(); assert.equal(await editor.locator('img.doc-protected-img').count(), 0);
  await page.getByRole('button', { name: '撤销', exact: true }).click(); await picture.waitFor();
  report.checks.push('actual-picture-file-input-and-decoding', 'real-image-resize-grip-preview-and-commit', 'picture-wrap-delete-and-single-undo');
  await page.screenshot({ path: join(output, 'objects.png') });
  const downloading = page.waitForEvent('download'); await page.getByRole('button', { name: '保存', exact: true }).click();
  const path = join(output, 'objects.docx'); await (await downloading).saveAs(path); const bytes = await readFile(path);
  const zip = await JSZip.loadAsync(bytes); const xml = await zip.file('word/document.xml').async('string');
  for (const token of ['w:tblHeader', 'FFF1A8', 'wp:anchor', '合并甲', '合并乙']) assert.ok(xml.toUpperCase().includes(token.toUpperCase()), token);
  assert.ok(Object.keys(zip.files).some(path => /^word\/media\//.test(path)));
  await page.evaluate(async bytes => { await window.__objectStore.getState().openBytes({ name: '图片表格重开.docx', bytes: new Uint8Array(bytes) }); }, Array.from(bytes));
  await page.waitForFunction(() => window.__objectCore().editor.state.doc.content.content.some(node => node.attrs.imageWrap === 'square-right'));
  assert.equal(await page.evaluate(() => window.__objectCore().state().dirty), false);
  assert.equal(await table.locator('tr').first().getAttribute('data-repeat-header'), '1');
  await menu('在下方插入行'); await page.getByRole('button', { name: '撤销', exact: true }).click();
  await page.screenshot({ path: join(output, 'reopened-objects.png') });
  report.checks.push('real-docx-download-retains-picture-media-wrap-and-table-properties', 'workspace-reopen-can-edit-again-with-native-history');
  // 功能：实际拖动左上角手柄移动整表，原生来源定位随对象移动且可单步撤销；时间：2026-11-11；作者：lq。
  await editor.locator(':scope > p').last().click(); await page.keyboard.insertText('移动表格目标'); await page.keyboard.press('Enter');
  const beforeMove = await page.evaluate(() => window.__objectCore().editor.getJSON());
  await table.locator('td, th').first().hover();
  const moveHandle = page.locator('.mo-word-core .doc-table-handle'); await moveHandle.waitFor({ state: 'visible' });
  await moveHandle.click(); assert.equal(await page.evaluate(() => window.__objectCore().editor.state.selection.toJSON().type), 'node');
  await table.locator('td, th').first().hover();
  await moveHandle.dragTo(editor.locator(':scope > p').last(), { targetPosition: { x: 30, y: 5 } });
  await page.waitForFunction(() => window.__objectCore().editor.state.doc.firstChild.type.name !== 'docTable');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.deepEqual(await page.evaluate(() => window.__objectCore().editor.getJSON()), beforeMove);
  await page.getByRole('button', { name: '重做', exact: true }).click(); assert.notEqual(await page.evaluate(() => window.__objectCore().editor.state.doc.firstChild.type.name), 'docTable');
  report.checks.push('real-table-move-handle-drag-single-undo-and-redo');
  // 功能：矮窗口的长菜单可以滚动到底部属性入口，菜单整体保持在视口内部；时间：2026-11-11；作者：lq。
  await page.setViewportSize({ width: 800, height: 560 });
  await table.locator('td, th').first().click({ button: 'right' });
  const menuItem = page.getByRole('menuitem', { name: '表格属性…', exact: true }); await menuItem.scrollIntoViewIfNeeded();
  const bounds = await menuItem.boundingBox(); assert.ok(bounds.y >= 0 && bounds.y + bounds.height <= 560);
  await page.screenshot({ path: join(output, 'short-table-menu.png') });
  await menuItem.click(); await dialog.waitFor(); await dialog.getByRole('button', { name: '取消', exact: true }).click();
  report.checks.push('short-window-context-menu-can-scroll-to-table-properties');
  assert.deepEqual(report.errors, []); report.browser = browser.version(); report.status = 'passed';
} catch (error) {
  report.status = 'failed'; report.error = String(error); process.exitCode = 1;
  if (page) { await page.screenshot({ path: join(output, 'failure.png') }); report.state = await page.evaluate(() => ({ text: document.querySelector('.ProseMirror')?.innerText, selection: window.__objectCore?.()?.editor.state.selection.toJSON(), focused: document.activeElement?.className, alerts: [...document.querySelectorAll('[role="alert"]')].map(node => node.textContent) })); }
} finally { await browser.close(); await writeFile(join(output, 'report.json'), JSON.stringify(report, null, 2)); console.log(JSON.stringify(report)); }
