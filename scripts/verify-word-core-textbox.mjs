/** 功能：真实产品文本框的工具栏、主历史、粘贴、查找及下载重开验收；时间：2026-11-11；作者：lq。 */
import { chromium } from 'playwright-core';
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';
import JSZip from 'jszip';
const root = resolve(import.meta.dirname, '..'), output = join(root, 'output/playwright/word-core-m3-textbox.local'); await mkdir(output, { recursive: true });
const fixture = await JSZip.loadAsync(await readFile(join(root, 'tests/fixtures/docx/contract.docx')));
fixture.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word"><w:body><w:p><w:r><w:t>正文隔离验收</w:t></w:r></w:p><w:p><w:r><w:pict><v:shape id="box" style="position:absolute;margin-left:40pt;margin-top:50pt;width:300pt;height:180pt;z-index:1"><v:textbox><w:txbxContent><w:p><w:pPr><w:suppressOverlap/></w:pPr><w:r><w:t>文本框初始内容</w:t></w:r></w:p></w:txbxContent></v:textbox><w10:wrap type="square"/></v:shape></w:pict></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const report = { status: 'running', checks: [], errors: [] }; let page;
try {
  page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true }); page.setDefaultTimeout(20000);
  page.on('pageerror', error => report.errors.push(String(error))); await page.addInitScript(() => { window.showSaveFilePicker = undefined; });
  await page.goto(process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1431'); await page.getByRole('button', { name: 'DOCX 文件 Word 文档' }).click();
  await page.evaluate(async ({ root, bytes }) => {
    const { useWorkspaceStore } = await import('/@fs' + root + '/packages/document-model/src/index.ts'); window.__boxStore = useWorkspaceStore;
    window.__boxCore = () => useWorkspaceStore.getState().sessions[useWorkspaceStore.getState().activeId].wordCore;
    await useWorkspaceStore.getState().openBytes({ name: '文本框联合验收.docx', bytes: new Uint8Array(bytes) });
  }, { root, bytes: Array.from(await fixture.generateAsync({ type: 'uint8array' })) });
  const box = page.locator('[data-word-engine="word-core"] .doc-textbox-editor'); await box.waitFor(); await box.dblclick(); await box.press('Meta+a'); await page.keyboard.insertText('文本框即时保存');
  await page.getByRole('button', { name: '加粗', exact: true }).click(); await box.press('Meta+a'); await page.getByRole('button', { name: '加粗', exact: true }).click();
  assert.ok(await box.locator('strong').count()); assert.equal(await page.evaluate(() => window.__boxCore().editor.state.doc.firstChild.textContent), '正文隔离验收');
  report.checks.push('doubleclick-text-mode-toolbar-and-body-isolation');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.equal(await box.locator('strong').count(), 0);
  await page.getByRole('button', { name: '重做', exact: true }).click(); assert.ok(await box.locator('strong').count()); report.checks.push('toolbar-single-main-undo-redo');
  await page.getByRole('button', { name: '样式库', exact: true }).click(); await page.getByRole('button', { name: '新建样式…', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: '新建样式', exact: true }); await dialog.getByPlaceholder('例如：合同条款').fill('文本框共享样式'); await dialog.getByRole('button', { name: '创建', exact: true }).click();
  const style = await page.evaluate(() => window.__boxCore().commandEditor().state.doc.firstChild.attrs.styleId); assert.ok(style.startsWith('MoStyle_'));
  await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.equal(await box.locator('[data-style="' + style + '"]').count(), 0);
  await page.getByRole('button', { name: '重做', exact: true }).click(); assert.equal(await box.locator('[data-style="' + style + '"]').count(), 1); assert.ok((await box.innerText()).includes('文本框即时保存')); report.checks.push('region-style-definition-and-content-atomic-history');
  // 功能：文本框标尺采用该形状净宽并共享主历史；时间：2026-11-11；作者：lq。
  await page.getByRole('tab', { name: '视图', exact: true }).click();
  if (!(await page.locator('.mo-ruler').count())) await page.getByRole('button', { name: '标尺', exact: true }).click();
  const indent = page.locator('.mo-ruler').getByRole('button', { name: '左缩进', exact: true }); assert.equal(await indent.isDisabled(), false);
  const indentBox = await indent.boundingBox(), originalIndent = await page.evaluate(() => window.__boxCore().commandEditor().state.doc.firstChild.attrs.indentLeft);
  await page.mouse.move(indentBox.x+indentBox.width/2, indentBox.y+indentBox.height/2); await page.mouse.down(); await page.mouse.move(indentBox.x+indentBox.width/2+24, indentBox.y+indentBox.height/2, { steps: 8 }); await page.mouse.up();
  assert.ok(await page.evaluate(value => window.__boxCore().commandEditor().state.doc.firstChild.attrs.indentLeft >= value+350, originalIndent));
  await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.equal(await page.evaluate(() => window.__boxCore().commandEditor().state.doc.firstChild.attrs.indentLeft), originalIndent);
  await page.getByRole('tab', { name: '开始', exact: true }).click(); report.checks.push('textbox-local-width-ruler-and-main-undo');
  await box.press('ArrowRight'); await page.waitForFunction(() => window.__boxCore().commandEditor().state.selection.empty); await box.press('Meta+ArrowDown');
  report.beforeEnter = await page.evaluate(() => ({ selection: window.__boxCore().commandEditor().state.selection.toJSON(), text: window.__boxCore().commandEditor().getText(), dom: getSelection()?.toString() }));
  await box.press('Enter');
  report.afterEnter = await page.evaluate(() => ({ selection: window.__boxCore().commandEditor().state.selection.toJSON(), text: window.__boxCore().commandEditor().getText() })); await page.keyboard.insertText('列表第一项');
  await page.getByRole('button', { name: '编号', exact: true }).click(); await box.press('End'); await box.press('Enter'); await page.keyboard.insertText('列表第二项');
  assert.equal(await box.locator('[data-marker]').count(), 2); await box.press('Tab'); assert.equal(await page.evaluate(() => window.__boxCore().commandEditor().state.selection.$from.parent.attrs.listMarker.ilvl), 1);
  await box.press('Shift+Tab'); await box.press('End'); await box.press('Enter'); await box.press('Enter');
  assert.equal(await page.evaluate(() => window.__boxCore().commandEditor().state.selection.$from.parent.attrs.listMarker), null); assert.ok((await box.innerText()).includes('文本框即时保存'), '列表后应保留首段'); report.checks.push('ordered-enter-tab-shifttab-and-empty-exit');
  await box.press('Meta+f'); await page.getByRole('textbox', { name: '查找内容', exact: true }).fill('第一项');
  await page.getByRole('textbox', { name: '替换为', exact: true }).fill('替换项'); await page.getByRole('button', { name: '全部替换', exact: true }).click();
  assert.ok((await box.innerText()).includes('列表替换项')); await page.getByRole('search', { name: '查找替换' }).getByRole('button', { name: '关闭', exact: true }).click(); assert.ok((await box.innerText()).includes('文本框即时保存'), '替换后应保留首段'); report.checks.push('textbox-find-replace-isolated');
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']); await box.press('Meta+ArrowDown');
  await page.evaluate(async () => navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob(['<p><b style="color:#D92D20">框内粘贴</b></p>'], { type: 'text/html' }), 'text/plain': new Blob(['框内粘贴'], { type: 'text/plain' }) })]));
  await box.press('Meta+v'); const options = page.getByRole('group', { name: '粘贴选项', exact: true }); await options.waitFor(); await options.getByRole('button', { name: '仅保留文本', exact: true }).click();
  assert.ok((await box.innerText()).includes('框内粘贴')); await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.ok(!(await box.innerText()).includes('框内粘贴'));
  await page.getByRole('button', { name: '重做', exact: true }).click(); assert.ok((await box.innerText()).includes('框内粘贴')); report.checks.push('real-clipboard-mode-replay-single-main-history');
  // 功能：真实文件选择在文本框插图，媒体进入下载包并在工作区重开后完成解码；时间：2026-11-11；作者：lq。
  await box.press('ArrowRight'); await page.waitForFunction(() => window.__boxCore().commandEditor().state.selection.empty); await box.press('Meta+ArrowDown');
  const png = await page.evaluate(() => { const canvas = document.createElement('canvas'); canvas.width = 90; canvas.height = 30; const context = canvas.getContext('2d'); context.fillStyle = '#246b8e'; context.fillRect(0,0,90,30); return canvas.toDataURL('image/png').split(',')[1]; });
  const pictureFile = join(output, 'textbox-picture.png'); await writeFile(pictureFile, Buffer.from(png, 'base64'));
  await page.getByRole('tab', { name: '插入', exact: true }).click();
  const chooser = page.waitForEvent('filechooser'); await page.getByRole('button', { name: '插入图片', exact: true }).click(); await (await chooser).setFiles(pictureFile);
  await box.locator('img').waitFor(); await box.locator('img').evaluate(image => image.decode());
  await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.equal(await box.locator('img').count(), 0);
  await page.getByRole('button', { name: '重做', exact: true }).click(); await box.locator('img').waitFor();
  report.checks.push('textbox-picture-file-input-and-single-main-history');
  await box.scrollIntoViewIfNeeded(); await page.screenshot({ path: join(output, 'textbox-combined.png') });
  const download = page.waitForEvent('download'); await page.getByRole('button', { name: '保存', exact: true }).click(); const savedPath = join(output, 'textbox-saved.docx'); await (await download).saveAs(savedPath);
  const savedBytes = await readFile(savedPath), saved = await JSZip.loadAsync(savedBytes), xml = await saved.file('word/document.xml').async('string');
  assert.ok(xml.includes('文本框即时保存') && xml.includes('列表替换项') && xml.replace(/<[^>]+>/g, '').includes('框内粘贴') && xml.includes('w:suppressOverlap') && xml.includes('<w:numPr>') && xml.includes(style));
  assert.ok(!xml.includes('aidocs-new-pic-')); const pictureMedia = Object.keys(saved.files).filter(path => /^word\/media\/.*\.png$/.test(path)); assert.ok(pictureMedia.length > 0);
  assert.ok((await Promise.all(pictureMedia.map(path => saved.file(path).async('base64')))).includes(png));
  await page.evaluate(async bytes => window.__boxStore.getState().openBytes({ name: '文本框重开.docx', bytes: new Uint8Array(bytes) }), Array.from(savedBytes));
  await box.waitFor(); assert.ok((await box.innerText()).includes('列表替换项')); await box.dblclick(); await box.locator('img').evaluate(image => image.decode()); assert.ok(await box.locator('img').evaluate(image => image.naturalWidth > 0)); await box.press('Meta+ArrowDown'); await page.keyboard.insertText('重开可续写');
  assert.equal(await page.evaluate(() => window.__boxCore().editor.state.doc.firstChild.textContent), '正文隔离验收'); report.checks.push('save-while-open-docx-reopen-and-continue');
  assert.deepEqual(report.errors, []); report.status = 'passed';
} catch (error) { report.status = 'failed'; report.error = String(error.stack ?? error); if (page) report.diagnostic = await page.evaluate(() => ({ doc: window.__boxCore().editor.getJSON(), active: window.__boxCore().commandEditor().getJSON(), state: window.__boxCore().state(), focus: document.activeElement?.outerHTML.slice(0, 200) })); if (page) await page.screenshot({ path: join(output, 'failed.png') }).catch(() => {}); throw error; }
finally { await writeFile(join(output, 'report.json'), JSON.stringify(report, null, 2)); await browser.close(); console.log(JSON.stringify(report, null, 2)); }
