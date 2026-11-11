/** 功能：实际剪贴板/三模式/延迟切换历史/属性面板与 DOCX 重开验证；时间：2026-11-11；作者：lq。 */
import { chromium } from 'playwright-core';
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';
import JSZip from 'jszip';
const root = resolve(import.meta.dirname, '..');
const output = join(root, 'output/playwright/word-core-m3-paste-properties.local');
await mkdir(output, { recursive: true });
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const report = { status: 'running', checks: [], errors: [] }; let page;
try {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true, permissions: ['clipboard-read', 'clipboard-write'] });
  page = await context.newPage(); page.setDefaultTimeout(18000); page.on('pageerror', error => report.errors.push(String(error)));
  await page.addInitScript(() => { window.showSaveFilePicker = undefined; });
  await page.goto(process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1431');
  await page.getByRole('button', { name: 'DOCX 文件 Word 文档' }).click();
  const editor = page.locator('[data-word-engine="word-core"] .ProseMirror').first(); await editor.click();
  await page.evaluate(async root => {
    const { useWorkspaceStore } = await import('/@fs' + root + '/packages/document-model/src/index.ts');
    window.__pasteStore = useWorkspaceStore;
    window.__pasteCore = () => useWorkspaceStore.getState().sessions[useWorkspaceStore.getState().activeId].wordCore;
  }, root);
  await page.keyboard.insertText('目标中文 hello ');
  // 功能：使用浏览器真实 Clipboard API 和系统粘贴快捷键，不调用编辑核心插入方法；时间：2026-11-11；作者：lq。
  const html = '<p style="font-family:Georgia;font-size:24pt;color:#cc2200"><strong>来源中文</strong><em>斜体</em></p><p>第二段</p>';
  const paste = async (html, text) => {
    await page.evaluate(async ({ html, text }) => navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([html], { type: 'text/html' }), 'text/plain': new Blob([text], { type: 'text/plain' }) })]), { html, text });
    await page.keyboard.press('Meta+V');
  };
  await paste(html, '来源中文斜体\n第二段');
  const options = page.getByRole('group', { name: '粘贴选项', exact: true }); await options.waitFor();
  assert.ok((await editor.innerText()).includes('来源中文'));
  const style = () => page.evaluate(() => {
    let result; window.__pasteCore().editor.state.doc.descendants(node => { if (node.isText && node.text.includes('来源中文')) result = { attrs: node.marks.find(mark => mark.type.name === 'docTextStyle')?.attrs, marks: node.marks.map(mark => mark.type.name) }; }); return result;
  });
  assert.equal((await style()).attrs.sizeHalfPoints, 48); assert.ok((await style()).marks.includes('bold'));
  report.checks.push('real-browser-clipboard-keeps-source-font-size-color-and-emphasis');
  // 功能：等待超过原生连续输入分组时限，确认模式切换仍属于原粘贴的一次历史；时间：2026-11-11；作者：lq。
  await page.waitForTimeout(700);
  await options.getByRole('button', { name: '匹配目标格式', exact: true }).click(); assert.notEqual((await style()).attrs?.sizeHalfPoints, 48); assert.ok((await style()).marks.includes('bold'));
  await options.getByRole('button', { name: '仅保留文本', exact: true }).click(); assert.ok(!(await style()).marks.includes('bold')); assert.ok(!(await style()).marks.includes('italic'));
  assert.equal((await editor.innerText()).split('来源中文').length, 2);
  await options.getByRole('button', { name: '保留源格式', exact: true }).click(); assert.equal((await style()).attrs.sizeHalfPoints, 48);
  await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.equal((await editor.innerText()).trim(), '目标中文 hello');
  await page.getByRole('button', { name: '重做', exact: true }).click(); assert.ok((await editor.innerText()).includes('来源中文'));
  report.checks.push('three-paste-options-do-not-duplicate-content', 'delayed-multiple-mode-changes-single-undo-and-redo');
  await editor.click(); await page.keyboard.press('Meta+End'); await page.keyboard.insertText('后续');
  assert.equal(await options.count(), 0);
  report.checks.push('later-edits-dismiss-stale-paste-options');
  await page.getByRole('tab', { name: '布局', exact: true }).click(); await page.getByRole('button', { name: '布局与属性', exact: true }).click();
  const panel = page.getByRole('complementary', { name: '文档属性' });
  const stats = await page.evaluate(() => window.__pasteCore().statistics());
  const statsText = await panel.getByRole('region', { name: '文档统计' }).innerText();
  assert.ok(statsText.includes(String(stats.words))); assert.ok(statsText.includes(String(stats.characters)));
  const revision = await page.evaluate(() => window.__pasteCore().state().revision);
  await panel.getByRole('button', { name: '打开段落设置…' }).click(); await page.getByRole('dialog', { name: '段落', exact: true }).waitFor();
  await page.getByRole('button', { name: '取消', exact: true }).click();
  await panel.getByRole('button', { name: '查看兼容性报告' }).click(); const compatibility = page.getByRole('dialog', { name: '兼容性报告', exact: true }); await compatibility.waitFor();
  await compatibility.getByRole('button', { name: '关闭', exact: true }).last().click();
  assert.equal(await page.evaluate(() => window.__pasteCore().state().revision), revision);
  await page.screenshot({ path: join(output, 'properties.png') });
  await panel.getByRole('button', { name: '关闭属性面板' }).click();
  report.checks.push('property-statistics-match-live-core-and-statusbar', 'paragraph-and-compatibility-entries-open-without-dirty-change');
  const downloading = page.waitForEvent('download'); await page.getByRole('button', { name: '保存', exact: true }).click();
  const path = join(output, 'paste-properties.docx'); await (await downloading).saveAs(path); const bytes = await readFile(path);
  const zip = await JSZip.loadAsync(bytes); const xml = await zip.file('word/document.xml').async('string');
  for (const token of ['来源中文', '第二段', 'w:sz w:val="48"', 'CC2200']) assert.ok(xml.toUpperCase().includes(token.toUpperCase()), token);
  await page.evaluate(async bytes => window.__pasteStore.getState().openBytes({ name: '粘贴重开.docx', bytes: new Uint8Array(bytes) }), Array.from(bytes));
  await page.waitForFunction(() => window.__pasteCore().editor.state.doc.textContent.includes('来源中文'));
  assert.equal((await style()).attrs.sizeHalfPoints, 48); assert.equal(await options.count(), 0);
  assert.equal(await page.evaluate(() => window.__pasteCore().state().dirty), false);
  await page.screenshot({ path: join(output, 'reopened-paste.png') });
  report.checks.push('real-docx-download-and-reopen-preserve-paste-format');
  // 功能：浏览器真实 PNG 剪贴板经过文件读取/解码，在正文保存为实际媒体；复制剪切继续走原生剪贴板；时间：2026-11-11；作者：lq。
  await page.evaluate(() => window.__pasteStore.getState().createDocument('docx')); await editor.click();
  await page.evaluate(async () => {
    const canvas = document.createElement('canvas'); canvas.width = 120; canvas.height = 60;
    const context = canvas.getContext('2d'); context.fillStyle = '#246b8e'; context.fillRect(0, 0, 120, 60);
    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
    await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
  });
  await page.keyboard.press('Meta+V');
  const image = editor.locator('img.doc-protected-img'); await image.waitFor();
  report.checks.push('real-png-clipboard-file-decoded-and-inserted');
  assert.equal(await page.getByRole('group', { name: '粘贴选项', exact: true }).count(), 0);
  await image.click(); await page.keyboard.press('Meta+C');
  await editor.locator(':scope > p').last().click(); await page.keyboard.press('Meta+V');
  await page.waitForFunction(() => document.querySelectorAll('[data-word-engine="word-core"] img.doc-protected-img').length === 2);
  report.checks.push('real-image-native-copy-paste');
  await image.last().click(); await page.keyboard.press('Meta+X');
  await page.waitForFunction(() => document.querySelectorAll('[data-word-engine="word-core"] img.doc-protected-img').length === 1);
  await page.getByRole('button', { name: '撤销', exact: true }).click();
  await page.waitForFunction(() => document.querySelectorAll('[data-word-engine="word-core"] img.doc-protected-img').length === 2);
  const imageDownload = page.waitForEvent('download'); await page.getByRole('button', { name: '保存', exact: true }).click();
  const imagePath = join(output, 'clipboard-images.docx'); await (await imageDownload).saveAs(imagePath); const imageBytes = await readFile(imagePath);
  const imageZip = await JSZip.loadAsync(imageBytes); assert.ok(Object.keys(imageZip.files).some(name => /^word\/media\//.test(name)));
  await page.evaluate(async bytes => window.__pasteStore.getState().openBytes({ name: '截图重开.docx', bytes: new Uint8Array(bytes) }), Array.from(imageBytes));
  await page.waitForFunction(() => document.querySelectorAll('[data-word-engine="word-core"] img.doc-protected-img').length === 2);
  report.checks.push('real-png-clipboard-decoding-native-image-copy-cut-undo-and-docx-reopen');
  // 功能：实际产品打开含批注/缺失页眉字体的包，报告与状态栏读取相同检测结果；时间：2026-11-11；作者：lq。
  zip.file('word/comments.xml', '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>');
  zip.file('word/headerM3.xml', '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:rPr><w:rFonts w:ascii="DefinitelyMissingM3Font20261111"/></w:rPr><w:t>页眉字体</w:t></w:r></w:p></w:hdr>');
  const complex = await zip.generateAsync({ type: 'uint8array' });
  await page.evaluate(async bytes => window.__pasteStore.getState().openBytes({ name: '兼容性检测.docx', bytes: new Uint8Array(bytes) }), Array.from(complex));
  await page.waitForFunction(() => window.__pasteStore.getState().sessions[window.__pasteStore.getState().activeId].capability.level === 2);
  await compatibility.waitFor();
  assert.ok((await compatibility.innerText()).includes('批注部件和引用保留'));
  assert.ok((await compatibility.innerText()).includes('DefinitelyMissingM3Font20261111'));
  // 功能：等待报告入场动画完成，截图能评估真实文字与遮罩对比度；时间：2026-11-11；作者：lq。
  await compatibility.evaluate(dialog => Promise.all(dialog.parentElement.getAnimations({ subtree: true }).map(animation => animation.finished.catch(() => {}))));
  await page.screenshot({ path: join(output, 'compatibility.png') });
  await compatibility.getByRole('button', { name: '关闭', exact: true }).last().click();
  report.checks.push('real-compatibility-report-detects-comments-and-header-fonts-as-l2');
  // 功能：真实包含宏时保存和插入入口禁用，键盘输入不能绕过只读；时间：2026-11-11；作者：lq。
  zip.file('word/vbaProject.bin', 'M3 macro fixture');
  const macro = await zip.generateAsync({ type: 'uint8array' });
  await page.evaluate(async bytes => window.__pasteStore.getState().openBytes({ name: '宏只读.docm', bytes: new Uint8Array(bytes) }), Array.from(macro));
  await page.waitForFunction(() => !window.__pasteCore().editor.isEditable);
  if (await compatibility.isVisible()) await compatibility.getByRole('button', { name: '关闭', exact: true }).last().click();
  assert.ok(await page.getByRole('button', { name: '保存', exact: true }).isDisabled());
  const macroText = await editor.innerText(); await editor.click(); await page.keyboard.insertText('不应写入'); assert.equal(await editor.innerText(), macroText);
  assert.equal(await page.evaluate(() => window.__pasteCore().state().dirty), false);
  await page.screenshot({ path: join(output, 'macro-readonly.png') });
  report.checks.push('macro-docm-is-readonly-and-cannot-save-or-type');
  assert.deepEqual(report.errors, []); report.browser = browser.version(); report.status = 'passed';
} catch (error) { report.status = 'failed'; report.error = String(error); report.stack = error.stack; process.exitCode = 1; if (page) { await page.screenshot({ path: join(output, 'failure.png') }); report.state = await page.evaluate(() => ({ selection: window.__pasteCore?.().editor.state.selection.toJSON(), images: document.querySelectorAll('[data-word-engine="word-core"] img.doc-protected-img').length, focused: document.activeElement?.className, errors: [...document.querySelectorAll('[role="alert"]')].map(node => node.textContent) })); } }
finally { await browser.close(); await writeFile(join(output, 'report.json'), JSON.stringify(report, null, 2)); console.log(JSON.stringify(report)); }
