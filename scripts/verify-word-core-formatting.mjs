/** 功能：在实际产品中验收 M3 字体/段落浮层、混合格式、下载和重开；时间：2026-11-10；作者：lq。 */
import { chromium } from 'playwright-core';
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { resolve, join } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '..');
const output = join(root, 'output/playwright/word-core-m3.local');
await mkdir(output, { recursive: true });
const origin = process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1431';
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const report = { status: 'running', checks: [], errors: [] };
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true });
  page.setDefaultTimeout(20000);
  page.on('pageerror', error => report.errors.push(String(error)));
  // 功能：真实下载保存，不模拟编辑器、序列化器或写盘；时间：2026-11-10；作者：lq。
  await page.addInitScript(() => { window.showSaveFilePicker = undefined; });
  await page.goto(origin);
  await page.getByRole('button', { name: 'DOCX 文件 Word 文档' }).click();
  const editor = page.locator('[data-word-engine="word-core"] .ProseMirror');
  await editor.click();
  await page.keyboard.insertText('中文字体与段落设置验收');
  await page.keyboard.press('Meta+a');
  await page.getByRole('button', { name: '字号', exact: true }).click();
  await page.getByRole('textbox', { name: '字号磅值' }).fill('18');
  await page.getByRole('textbox', { name: '字号磅值' }).press('Enter');
  await page.getByRole('combobox', { name: '字体', exact: true }).selectOption('Arial');
  await page.getByRole('button', { name: '居中', exact: true }).click();
  await page.getByRole('button', { name: '字体颜色调色板', exact: true }).click();
  await page.getByRole('dialog', { name: '字体颜色', exact: true }).getByRole('button', { name: '#D92D20', exact: true }).click();
  await page.getByRole('button', { name: '文本背景颜色（突出显示）调色板', exact: true }).click();
  await page.getByRole('dialog', { name: '文本背景颜色（突出显示）', exact: true }).getByRole('button', { name: '#F79009', exact: true }).click();
  await page.getByRole('button', { name: '行距', exact: true }).click();
  await page.getByRole('button', { name: '固定值', exact: true }).click();
  await page.getByRole('textbox', { name: '固定值行距', exact: true }).fill('24');
  await page.getByRole('button', { name: '应用', exact: true }).click();
  // 功能：通过工作区只读检查同一核心的格式，不用检查入口替代真实 UI 命令；时间：2026-11-10；作者：lq。
  await page.evaluate(async root => {
    const { useWorkspaceStore } = await import('/@fs' + root + '/packages/document-model/src/index.ts');
    const { coreToolbarFormatting } = await import('/@fs' + root + '/packages/editor-word/src/core-formatting.ts');
    window.__m3Store = useWorkspaceStore;
    window.__m3Format = () => coreToolbarFormatting(useWorkspaceStore.getState().sessions[useWorkspaceStore.getState().activeId].wordCore);
  }, root);
  const formatted = await page.evaluate(() => window.__m3Format());
  assert.equal(formatted.activeStyle.fontSizePt, 18);
  assert.equal(formatted.activeStyle.fontFamily, 'Arial');
  assert.equal(formatted.activeStyle.color, '#D92D20');
  assert.equal(formatted.activeStyle.highlight, '#F79009');
  assert.equal(formatted.paragraph.alignment, 'center');
  assert.equal(formatted.paragraph.lineHeightExactPx, 32);
  report.checks.push('font-size-popover-retains-selection', 'font-selector-color-highlight', 'alignment-exact-line-spacing');
  // 功能：快捷键与按钮共用新核心事务，不能继续进入旧 DOM 命令；时间：2026-11-10；作者：lq。
  await page.keyboard.press('Meta+l');
  assert.equal((await page.evaluate(() => window.__m3Format())).paragraph.alignment, 'left');
  await page.keyboard.press('Meta+e');
  assert.equal((await page.evaluate(() => window.__m3Format())).paragraph.alignment, 'center');
  report.checks.push('alignment-keyboard-shortcuts');
  await page.getByRole('tab', { name: '布局', exact: true }).click();
  await page.getByRole('button', { name: '段落设置', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: '段落', exact: true });
  await dialog.getByLabel('段前（磅）').fill('6');
  await dialog.getByRole('button', { name: '确定', exact: true }).click();
  const after = await page.evaluate(() => window.__m3Format());
  assert.equal(after.paragraph.spaceBeforePx, 8);
  assert.equal(after.paragraph.alignment, 'center');
  assert.equal(after.paragraph.lineHeightExactPx, 32);
  await page.getByRole('button', { name: '撤销', exact: true }).click();
  assert.equal((await page.evaluate(() => window.__m3Format())).paragraph.spaceBeforePx, formatted.paragraph.spaceBeforePx);
  await page.getByRole('button', { name: '重做', exact: true }).click();
  report.checks.push('paragraph-dialog-partial-patch-and-single-undo');
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: '保存', exact: true }).click();
  const savedPath = join(output, 'formatted.docx');
  await (await downloading).saveAs(savedPath);
  const saved = await readFile(savedPath);
  const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml').async('text');
  for (const fragment of ['中文字体与段落设置验收', 'D92D20', 'F79009', 'w:val="36"', 'w:line="480"', 'w:before="120"']) assert.ok(xml.includes(fragment), fragment);
  const id = await page.evaluate(bytes => window.__m3Store.getState().openBytes({ name: '格式重开.docx', bytes: new Uint8Array(bytes) }), [...saved]);
  await page.waitForFunction(id => document.querySelector('[data-word-engine="word-core"] .ProseMirror') === window.__m3Store.getState().sessions[id]?.wordCore.editor.view.dom, id);
  await editor.click(); await page.keyboard.press('Meta+a');
  const reopened = await page.evaluate(() => window.__m3Format());
  assert.deepEqual(reopened.activeStyle, after.activeStyle);
  assert.equal(reopened.paragraph.spaceBeforePx, 8);
  assert.equal(reopened.paragraph.lineHeightExactPx, 32);
  report.checks.push('actual-download-docx-and-reopen-format-state');
  // 功能：通过实际键盘形成不同字号选区，工具栏应显示混合而非默认字号；时间：2026-11-10；作者：lq。
  await page.keyboard.press('ArrowRight'); await page.keyboard.press('End');
  await page.keyboard.insertText('新增');
  await page.keyboard.press('Shift+ArrowLeft'); await page.keyboard.press('Shift+ArrowLeft');
  await page.getByRole('tab', { name: '开始', exact: true }).click();
  await page.getByRole('button', { name: '字号', exact: true }).click();
  await page.getByRole('textbox', { name: '字号磅值' }).fill('12');
  await page.getByRole('textbox', { name: '字号磅值' }).press('Enter');
  await page.keyboard.press('Meta+a');
  // 功能：等待 React 完成原生选区摘要回填，再检查混合格式；时间：2026-11-10；作者：lq。
  await page.waitForFunction(() => document.querySelector('button[aria-label="字号"]')?.textContent.includes('混合字号'));
  assert.ok((await page.getByRole('button', { name: '字号', exact: true }).innerText()).includes('混合字号'));
  report.checks.push('mixed-size-state-from-real-keyboard-selection');
  await page.keyboard.press('ArrowRight');
  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '插入分页符', exact: true }).click();
  assert.ok(await page.evaluate(() => {
    const core = window.__m3Store.getState().sessions[window.__m3Store.getState().activeId].wordCore;
    return core.editor.getJSON().content.some(node => node.attrs?.pageBreakBefore);
  }));
  await page.getByRole('button', { name: '撤销', exact: true }).click();
  await page.getByRole('tab', { name: '开始', exact: true }).click();
  await page.keyboard.press('Meta+a');
  report.checks.push('page-break-button-and-undo');
  // 功能：截图前回到文档顶部，确保检查真实正文而不是滚动后的空白页尾；时间：2026-11-10；作者：lq。
  await editor.evaluate(element => { element.closest('.mo-scroll').scrollTop = 0; });
  await page.waitForFunction(() => {
    const paragraph = document.querySelector('[data-word-engine="word-core"] .ProseMirror p');
    if (!paragraph) return false;
    const box = paragraph.getBoundingClientRect();
    return box.top > 190 && box.bottom < innerHeight;
  });
  await page.screenshot({ path: join(output, 'formatting.png') });
  assert.deepEqual(report.errors, []);
  report.status = 'passed'; report.browser = browser.version();
} catch (error) { report.status = 'failed'; report.error = String(error); report.diagnostic = await browser.contexts()[0]?.pages()[0]?.evaluate(() => ({ format: window.__m3Format?.(), active: document.activeElement?.outerHTML, selection: window.getSelection()?.toString() })); await browser.contexts()[0]?.pages()[0]?.screenshot({ path: join(output, 'failure.png') }); process.exitCode = 1; }
finally { await browser.close(); await writeFile(join(output, 'report.json'), JSON.stringify(report, null, 2)); console.log(JSON.stringify(report)); }
