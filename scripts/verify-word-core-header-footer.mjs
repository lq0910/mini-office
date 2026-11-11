/** 功能：真实双击区域、富文本/布局表格、变体/继承及保存重开的浏览器验收；时间：2026-11-11；作者：lq。 */
import { chromium } from 'playwright-core';
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';
import JSZip from 'jszip';
const root = resolve(import.meta.dirname, '..'), output = join(root, 'output/playwright/word-core-m3-header-footer.local'); await mkdir(output, { recursive: true });
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const report = { status: 'running', checks: [], errors: [] }; let page;
try {
  page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true }); page.setDefaultTimeout(25000);
  page.on('pageerror', error => report.errors.push(String(error))); await page.addInitScript(() => { window.showSaveFilePicker = undefined; window.showOpenFilePicker = undefined; });
  await page.goto(process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1431'); await page.getByRole('button', { name: 'DOCX 文件 Word 文档' }).click();
  await page.evaluate(async root => { const { useWorkspaceStore } = await import('/@fs' + root + '/packages/document-model/src/index.ts'); window.__hfStore = useWorkspaceStore; window.__hfCore = () => useWorkspaceStore.getState().sessions[useWorkspaceStore.getState().activeId].wordCore; }, root);
  const body = page.locator('[data-word-engine="word-core"] .doc-page').first(); await body.click(); await page.keyboard.insertText('正文不应被区域操作改动');
  const before = await page.evaluate(() => ({ revision: window.__hfCore().state().revision, text: window.__hfCore().editor.getText() }));
  await page.locator('[data-hf-kind="header"][data-hf-page="0"]').dblclick();
  let region = page.getByRole('textbox', { name: '直接编辑页眉' }); await page.locator('.mo-core-hf-editing .ProseMirror').waitFor();
  assert.equal(await page.evaluate(() => window.__hfCore().state().revision), before.revision);
  await page.keyboard.insertText('区域页眉中文');
  await page.getByRole('button', { name: '区域插入页码', exact: true }).click(); await page.keyboard.insertText(' / ');
  await page.getByRole('button', { name: '区域插入总页数', exact: true }).click();
  await page.waitForFunction(() => document.querySelector('.mo-core-hf-editing .hf-field')?.getAttribute('data-hf-num') === '1');
  const regionEditor = page.locator('.mo-core-hf-editing .ProseMirror'); await regionEditor.press('Meta+a');
  await page.getByRole('button', { name: '加粗', exact: true }).click();
  assert.ok(await regionEditor.locator('strong').count()); assert.equal(await page.evaluate(() => window.__hfCore().editor.getText()), before.text);
  report.checks.push('empty-strip-doubleclick-no-dirty-and-native-rich-text-toolbar');
  // 功能：区域保持打开时下载，证明保存快照不依赖失焦提交；时间：2026-11-11；作者：lq。
  let download = page.waitForEvent('download'); await page.getByRole('button', { name: '保存', exact: true }).click(); let savedPath = join(output, 'open-region-saved.docx'); await (await download).saveAs(savedPath);
  const saved = await JSZip.loadAsync(await readFile(savedPath)); const headerPath = Object.keys(saved.files).find(path => /^word\/header\d+\.xml$/.test(path)); const header = await saved.file(headerPath).async('string'); assert.ok(header.includes('区域页眉中文') && header.includes('NUMPAGES') && header.includes('<w:b'));
  assert.equal(await regionEditor.count(), 1); report.checks.push('save-while-region-open-includes-text-format-and-live-fields');
  await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.equal(await regionEditor.locator('strong').count(), 0);
  await page.getByRole('button', { name: '重做', exact: true }).click(); assert.ok(await regionEditor.locator('strong').count());
  await regionEditor.press('Escape'); assert.equal(await page.locator('.mo-core-hf-editing').count(), 0); report.checks.push('one-main-history-undo-redo-and-escape');

  // 功能：实际样式库在区域创建/套用并复用到正文，格式刷通过鼠标跨区域应用；时间：2026-11-11；作者：lq。
  await page.getByRole('tab', { name: '开始', exact: true }).click();
  await page.locator('[data-hf-kind="header"][data-hf-page="0"]').dblclick();
  await regionEditor.press('Meta+a');
  await page.getByRole('button', { name: '样式库', exact: true }).click();
  await page.getByRole('button', { name: '新建样式…', exact: true }).click();
  const styleDialog = page.getByRole('dialog', { name: '新建样式', exact: true });
  await styleDialog.getByPlaceholder('例如：合同条款').fill('页眉共享样式'); await styleDialog.getByRole('button', { name: '创建', exact: true }).click();
  const regionStyle = await page.evaluate(() => window.__hfCore().commandEditor().state.doc.firstChild.attrs.styleId);
  assert.ok(regionStyle.startsWith('MoStyle_'));
  await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.notEqual(await regionEditor.locator('[data-style="' + regionStyle + '"]').count(), 1);
  await page.getByRole('button', { name: '重做', exact: true }).click();

  // 功能：页眉标尺真实拖动只修改该区域，撤销仍走主文档；时间：2026-11-11；作者：lq。
  await page.getByRole('tab', { name: '视图', exact: true }).click(); await page.getByRole('button', { name: '标尺', exact: true }).click();
  const rulerHandle = page.locator('.mo-ruler').getByRole('button', { name: '左缩进', exact: true }); assert.equal(await rulerHandle.isDisabled(), false);
  const rulerOriginal = await page.evaluate(() => window.__hfCore().commandEditor().state.doc.firstChild.attrs.indentLeft);
  const rulerBefore = await page.evaluate(() => window.__hfCore().state().revision);
  const rulerBox = await rulerHandle.boundingBox(); await page.mouse.move(rulerBox.x + rulerBox.width / 2, rulerBox.y + rulerBox.height / 2); await page.mouse.down(); await page.mouse.move(rulerBox.x + rulerBox.width / 2 + 48, rulerBox.y + rulerBox.height / 2, { steps: 12 });
  assert.equal(await page.evaluate(() => window.__hfCore().state().revision), rulerBefore); await page.mouse.up();
  assert.ok(await page.evaluate(() => window.__hfCore().commandEditor().state.doc.firstChild.attrs.indentLeft >= 700));
  assert.equal(await page.evaluate(() => window.__hfCore().editor.state.doc.firstChild.attrs.indentLeft), null);
  await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.equal(await page.evaluate(() => window.__hfCore().commandEditor().state.doc.firstChild.attrs.indentLeft), rulerOriginal);
  await page.getByRole('tab', { name: '开始', exact: true }).click();
  report.checks.push('header-ruler-drag-preview-main-history-and-body-isolation');
  await page.getByRole('button', { name: '字号', exact: true }).click(); await page.getByRole('textbox', { name: '字号磅值' }).fill('18'); await page.getByRole('textbox', { name: '字号磅值' }).press('Enter');
  await page.getByRole('button', { name: '格式刷', exact: true }).click();
  const bodyBounds = await body.locator(':scope > p').first().evaluate(element => { const range = document.createRange(); range.selectNodeContents(element); const r = range.getBoundingClientRect(); return { x: r.x, y: r.y + r.height / 2, width: r.width }; });
  await page.mouse.move(bodyBounds.x + 1, bodyBounds.y); await page.mouse.down(); await page.mouse.move(bodyBounds.x + bodyBounds.width - 1, bodyBounds.y, { steps: 12 }); await page.mouse.up();
  await page.waitForFunction(() => document.querySelector('button[aria-label="格式刷"]')?.getAttribute('aria-pressed') === 'false');
  assert.equal(await page.evaluate(() => window.__hfCore().editor.state.doc.firstChild.firstChild.marks.find(mark => mark.type.name === 'docTextStyle').attrs.sizeHalfPoints), 36);
  assert.equal(await page.evaluate(() => window.__hfCore().editor.getText()), before.text);
  await page.getByRole('button', { name: '撤销', exact: true }).click();
  assert.notEqual(await page.evaluate(() => window.__hfCore().editor.state.doc.firstChild.firstChild.marks.find(mark => mark.type.name === 'docTextStyle')?.attrs.sizeHalfPoints), 36);
  await page.getByRole('button', { name: '样式库', exact: true }).click(); await page.locator(`[data-style-card-host="${regionStyle}"] button[data-style-card]`).click();
  assert.equal(await page.evaluate(() => window.__hfCore().editor.state.doc.firstChild.attrs.styleId), regionStyle);
  report.checks.push('region-style-creation-main-undo-redo-and-body-reuse', 'mouse-format-painter-from-header-to-body-single-undo');

  // 功能：区域内列表续项/退出、查找替换、真实剪贴板和图片文件共用当前区域命令；时间：2026-11-11；作者：lq。
  await page.locator('[data-hf-kind="header"][data-hf-page="0"]').dblclick();
  await regionEditor.press('Meta+End'); await regionEditor.press('Enter'); await page.keyboard.insertText('区域列表甲');
  await page.getByRole('button', { name: '编号', exact: true }).click(); assert.equal(await page.getByRole('button', { name: '编号', exact: true }).getAttribute('aria-pressed'), 'true');
  await regionEditor.press('End'); await regionEditor.press('Enter'); await page.keyboard.insertText('区域列表乙');
  assert.deepEqual(await regionEditor.locator('[data-marker]').evaluateAll(elements => elements.map(element => element.dataset.marker)), ['1.', '2.']);
  await regionEditor.press('Tab'); assert.equal(await page.evaluate(() => window.__hfCore().commandEditor().state.selection.$from.parent.attrs.listMarker.ilvl), 1);
  await regionEditor.press('Shift+Tab'); await regionEditor.press('End'); await regionEditor.press('Enter'); await regionEditor.press('Enter');
  assert.equal(await page.evaluate(() => window.__hfCore().commandEditor().state.selection.$from.parent.attrs.listMarker), null);
  await page.getByRole('button', { name: '查找替换', exact: true }).click();
  await page.getByRole('textbox', { name: '查找内容', exact: true }).fill('区域列表甲'); assert.equal(await page.getByRole('status', { name: '查找结果', exact: true }).innerText(), '1/1');
  await page.getByRole('textbox', { name: '替换为', exact: true }).fill('区域替换甲'); await page.getByRole('button', { name: '全部替换', exact: true }).click();
  assert.ok(await regionEditor.innerText().then(text => text.includes('区域替换甲'))); await page.getByRole('search', { name: '查找替换' }).getByRole('button', { name: '关闭', exact: true }).click();
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await regionEditor.press('Meta+End');
  await page.evaluate(async () => navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob(['<p><b style="color:#D92D20;font-size:24pt">区域粘贴</b></p>'], { type: 'text/html' }), 'text/plain': new Blob(['区域粘贴'], { type: 'text/plain' }) })]));
  await regionEditor.press('Meta+v'); const pasteOptions = page.getByRole('group', { name: '粘贴选项', exact: true }); await pasteOptions.waitFor();
  await pasteOptions.getByRole('button', { name: '仅保留文本', exact: true }).click();
  assert.ok(await regionEditor.innerText().then(text => text.includes('区域粘贴')));
  const chooser = page.waitForEvent('filechooser'); await page.getByRole('tab', { name: '插入', exact: true }).click(); await page.getByRole('button', { name: '插入图片', exact: true }).click();
  await (await chooser).setFiles({ name: '区域图片.png', mimeType: 'image/png', buffer: Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jBf8AAAAASUVORK5CYII=', 'base64') });
  await regionEditor.locator('img.doc-inline-img').waitFor();
  await body.evaluate(element => { element.closest('.mo-scroll').scrollTop = 0; });
  await page.screenshot({ path: join(output, 'region-combined-editing.png') });
  download = page.waitForEvent('download'); await page.getByRole('button', { name: '保存', exact: true }).click(); const combinedPath = join(output, 'combined-region-saved.docx'); await (await download).saveAs(combinedPath);
  const combined = await JSZip.loadAsync(await readFile(combinedPath)); const combinedHeader = await combined.file(headerPath).async('string');
  assert.ok(combinedHeader.includes('区域替换甲') && combinedHeader.includes('区域粘贴') && combinedHeader.includes('<w:numPr>') && /<w:drawing[\s>]/.test(combinedHeader));
  assert.equal(await page.evaluate(() => window.__hfCore().editor.getText()), before.text);
  await regionEditor.press('Escape');
  report.checks.push('header-list-enter-tab-empty-exit-and-region-search-replace', 'header-real-html-clipboard-mode-file-picture-and-docx-save');
  await page.getByRole('tab', { name: '布局', exact: true }).click(); await page.getByRole('button', { name: '页眉、页脚与页码', exact: true }).click();
  let dialog = page.getByRole('dialog', { name: '页眉、页脚与页码', exact: true }); await dialog.getByRole('checkbox', { name: '首页不同', exact: true }).check(); await dialog.getByRole('checkbox', { name: '奇偶页不同', exact: true }).check();
  await dialog.getByRole('combobox', { name: '页眉页脚页面类型' }).selectOption('first'); await dialog.getByRole('textbox', { name: '页脚内容' }).fill('首页页脚'); await dialog.getByRole('button', { name: '应用到文档', exact: true }).click();
  await page.waitForFunction(() => document.querySelector('.mo-core-footer')?.textContent.includes('首页页脚')); assert.equal(await page.evaluate(() => window.__hfCore().evenAndOddHeaders()), true); report.checks.push('first-page-and-even-odd-controls-repaginate');
  // 功能：两节继承的真实布局表格包含未知段落属性；解除继承后只改第二节第一格；时间：2026-11-11；作者：lq。
  const zip = await JSZip.loadAsync(await readFile(join(root, 'tests/fixtures/docx/contract.docx')));
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const geometry = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="480" w:footer="720"/>';
  const p = text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdRichHeader"/>${geometry}</w:sectPr></w:pPr><w:r><w:t>第一节正文</w:t></w:r></w:p>${p('第二节正文')}<w:sectPr>${geometry}</w:sectPr></w:body></w:document>`);
  zip.file('word/headers/rich.xml', `<w:hdr xmlns:w="${W}" xmlns:r="${R}"><w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr>${p('左侧合同编号')}</w:tc><w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:contextualSpacing/></w:pPr><w:hyperlink r:id="rIdLink"><w:r><w:t>右侧保留链接</w:t></w:r></w:hyperlink></w:p></w:tc></w:tr></w:tbl>${p('表格下方文字')}</w:hdr>`);
  zip.file('word/headers/_rels/rich.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="${R}/hyperlink" Target="https://example.com/contract" TargetMode="External"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', (await zip.file('word/_rels/document.xml.rels').async('string')).replace('</Relationships>', `<Relationship Id="rIdRichHeader" Type="${R}/header" Target="headers/rich.xml"/></Relationships>`));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml').async('string')).replace('</Types>', '<Override PartName="/word/headers/rich.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
  const rich = await zip.generateAsync({ type: 'uint8array' }); await writeFile(join(output, 'complex-regions.docx'), rich);
  await page.evaluate(async bytes => window.__hfStore.getState().openBytes({ name: '复杂区域验收.docx', bytes: new Uint8Array(bytes) }), Array.from(rich));
  await page.waitForFunction(() => document.querySelector('[data-hf-section="1"][data-hf-kind="header"]')?.textContent.includes('左侧合同编号'));
  await page.getByRole('tab', { name: '布局', exact: true }).click(); await page.getByRole('button', { name: '页眉、页脚与页码', exact: true }).click(); dialog = page.getByRole('dialog', { name: '页眉、页脚与页码', exact: true });
  await dialog.getByRole('combobox', { name: '页眉页脚分节' }).selectOption('1'); await dialog.getByRole('button', { name: '页眉', exact: true }).click(); assert.equal(await dialog.getByRole('checkbox', { name: '链接到前一节' }).isChecked(), true);
  await dialog.getByRole('checkbox', { name: '链接到前一节' }).uncheck(); await dialog.getByRole('button', { name: '应用到文档', exact: true }).click();
  const strip = page.locator('[data-hf-section="1"][data-hf-kind="header"]'); await strip.dblclick();
  // 功能：双列必须真实并排，功能点击通过不能代替编辑态布局验收；时间：2026-11-11；作者：lq。
  const cells = await page.locator('.mo-core-hf-editing .page-hf-cell').evaluateAll(elements => elements.map(element => { const r = element.getBoundingClientRect(); return { x: r.x, y: r.y, width: r.width }; }));
  assert.equal(cells.length, 2); assert.ok(Math.abs(cells[0].y - cells[1].y) < 2 && cells[1].x > cells[0].x + 100);
  const firstCell = page.locator('.mo-core-hf-editing .page-hf-cell').first(); await firstCell.click(); await page.keyboard.press('Home'); await page.keyboard.insertText('第二节独立：');
  await page.waitForFunction(() => document.querySelector('.mo-core-hf-editing')?.textContent.includes('第二节独立：'));
  assert.ok(!(await page.locator('[data-hf-section="0"][data-hf-kind="header"]').innerText()).includes('第二节独立：'));
  // 功能：页眉表格格内标尺以当前单元格净宽为原点，单步撤销不改相邻格；时间：2026-11-11；作者：lq。
  const cellRuler = page.locator('.mo-ruler').getByRole('button', { name: '左缩进', exact: true }); assert.equal(await cellRuler.isDisabled(), false);
  const cellIndent = await page.evaluate(() => window.__hfCore().commandEditor().state.selection.$from.parent.attrs.indentLeft), cellHandle = await cellRuler.boundingBox();
  await page.mouse.move(cellHandle.x+cellHandle.width/2, cellHandle.y+cellHandle.height/2); await page.mouse.down(); await page.mouse.move(cellHandle.x+cellHandle.width/2+24, cellHandle.y+cellHandle.height/2, { steps: 8 }); await page.mouse.up();
  assert.ok(await page.evaluate(value => window.__hfCore().commandEditor().state.selection.$from.parent.attrs.indentLeft >= (value ?? 0)+350, cellIndent));
  await page.getByRole('button', { name: '撤销', exact: true }).click(); assert.equal(await page.evaluate(() => window.__hfCore().commandEditor().state.selection.$from.parent.attrs.indentLeft), cellIndent);
  report.checks.push('header-table-cell-local-ruler-single-undo');
  // 功能：格内段落对话框真实修改右缩进，退出编辑及保存重开后纸面仍一致；时间：2026-11-11；作者：lq。
  await page.getByRole('tab', { name: '布局', exact: true }).click(); await page.getByRole('button', { name: '段落设置', exact: true }).click();
  const paragraph = page.getByRole('dialog', { name: '段落', exact: true }); await paragraph.getByLabel('右侧（字符）').fill('1'); await paragraph.getByRole('button', { name: '确定', exact: true }).click();
  const rightIndent = await page.evaluate(() => window.__hfCore().commandEditor().state.selection.$from.parent.attrs.indentRight); assert.ok(rightIndent > 0);
  await page.locator('.mo-core-hf-editing .ProseMirror').press('Escape');
  const paperCell = strip.locator('.page-hf-cell-para').first(); await page.waitForFunction(({ indent }) => Number.parseFloat(getComputedStyle(document.querySelector('[data-hf-section="1"][data-hf-kind="header"] .page-hf-cell-para')).marginInlineEnd) > indent / 15 - 1, { indent: rightIndent });
  assert.ok(Math.abs(await paperCell.evaluate(element => Number.parseFloat(getComputedStyle(element).marginInlineEnd)) - rightIndent / 15) < 1);
  report.checks.push('header-table-cell-right-indent-dialog-and-paper-render');
  await page.screenshot({ path: join(output, 'rich-table-direct-edit.png') }); report.checks.push('layout-table-cells-direct-edit-and-independent-section-inheritance');
  download = page.waitForEvent('download'); await page.getByRole('button', { name: '保存', exact: true }).click(); savedPath = join(output, 'complex-region-saved.docx'); await (await download).saveAs(savedPath);
  const result = await JSZip.loadAsync(await readFile(savedPath)); assert.equal(await result.file('word/headers/rich.xml').async('string'), await zip.file('word/headers/rich.xml').async('string'));
  const clonePath = Object.keys(result.files).find(path => /^word\/header\d+\.xml$/.test(path) && !zip.file(path)); const cloned = await result.file(clonePath).async('string');
  assert.ok(cloned.includes('第二节独立：') && cloned.includes('w:right=') && cloned.includes('contextualSpacing') && cloned.includes('右侧保留链接')); assert.ok((await result.file(clonePath.replace('word/', 'word/_rels/') + '.rels').async('string')).includes('https://example.com/contract'));
  await page.evaluate(async bytes => window.__hfStore.getState().openBytes({ name: '区域重开.docx', bytes: new Uint8Array(bytes) }), Array.from(await readFile(savedPath)));
  await page.waitForFunction(() => document.querySelector('[data-hf-section="1"][data-hf-kind="header"]')?.textContent.includes('第二节独立：')); report.checks.push('download-and-reopen-preserve-original-part-table-link-and-unknown-properties');
  assert.deepEqual(report.errors, []); report.status = 'passed'; report.browser = browser.version();
} catch (error) { report.status = 'failed'; report.error = String(error); process.exitCode = 1; if (page) await page.screenshot({ path: join(output, 'failure.png') }); }
finally { await browser.close(); await writeFile(join(output, 'report.json'), JSON.stringify(report, null, 2)); console.log(JSON.stringify(report)); }
