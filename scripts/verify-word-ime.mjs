/** 功能：通过真实 Chromium 输入法事件验证表格中文确认、取消及保存属性保持；时间：2026-11-07；作者：lq。 */
import { chromium } from 'playwright-core';
import JSZip from 'jszip';
import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { homedir } from 'node:os';

const [file, target = 'zhi h'] = process.argv.slice(2);
if (!file) throw new Error('用法：node scripts/verify-word-ime.mjs 文件.docx [目标段落文字]');
const output = resolve('output/playwright/word-ime.local');
await mkdir(output, { recursive: true });
const original = await readFile(file);
const originalZip = await JSZip.loadAsync(original);
const originalXml = await originalZip.file('word/document.xml').async('text');
const finalText = '智慧物业信息系统部署实施';
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const report = { checks: [] };
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  page.setDefaultTimeout(15000);
  const errors = [];
  page.on('pageerror', error => errors.push(String(error)));
  await page.addInitScript(() => { window.showOpenFilePicker = undefined; window.showSaveFilePicker = undefined; });
  await page.goto(process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1420/');
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: '打开文件', exact: true }).first().click();
  await (await chooser).setFiles(resolve(file));
  const targetParagraph = page.locator('[data-paragraph-id][role="textbox"]').filter({ hasText: target }).first();
  await targetParagraph.waitFor({ state: 'visible', timeout: 60000 });
  const id = await targetParagraph.getAttribute('data-paragraph-id');
  const paragraph = page.locator(`[data-paragraph-id="${id}"][role="textbox"]`).first();
  /** 功能：只读观察产品模型，所有正文修改均通过浏览器输入或 DOM 事件完成；时间：2026-11-07；作者：lq。 */
  const moduleUrl = `/@fs${resolve('packages/document-model/src/index.ts')}`;
  const model = () => page.evaluate(async ({ id, moduleUrl }) => {
    const { useWorkspaceStore } = await import(moduleUrl);
    const state = useWorkspaceStore.getState();
    const paragraphs = [];
    const visit = blocks => blocks.forEach(block => {
      if (block.type === 'paragraph') paragraphs.push(block);
      if (block.type === 'table') block.rows.forEach(row => row.cells.forEach(cell => visit(cell.blocks)));
    });
    visit(state.sessions[state.activeId].data.blocks);
    return { text: paragraphs.find(p => p.id === id)?.runs.map(run => run.text ?? '').join(''), count: paragraphs.length };
  }, { id, moduleUrl });
  /** 功能：仅设置目标格文字选区，避免全选覆盖整篇正文；时间：2026-11-07；作者：lq。 */
  const selectText = async () => {
    await paragraph.click();
    await paragraph.evaluate(element => {
      const range = document.createRange();
      range.selectNodeContents(element);
      const selection = window.getSelection();
      selection.removeAllRanges(); selection.addRange(range);
    });
  };
  const before = await model();
  const cdp = await page.context().newCDPSession(page);
  await selectText();
  await cdp.send('Input.imeSetComposition', { text: 'zhi hui wu ye', selectionStart: 12, selectionEnd: 12 });
  assert.deepEqual(await model(), before, '候选拼音不得进入正文模型');
  // 功能：同时覆盖标准 isComposing 和 WebKit 仅给出 keyCode 229 的候选确认回车；时间：2026-11-07；作者：lq。
  for (const properties of [{ isComposing: true }, { keyCode: 229 }]) {
    const prevented = await paragraph.evaluate((element, properties) => {
      const event = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true, ...properties });
      element.dispatchEvent(event);
      return event.defaultPrevented;
    }, properties);
    assert.equal(prevented, false, '候选确认回车不得被分段逻辑拦截');
    assert.deepEqual(await model(), before, '候选确认不得新增段落');
  }
  await cdp.send('Input.insertText', { text: finalText });
  await page.waitForFunction(async ({ moduleUrl, finalText }) => {
    const { useWorkspaceStore } = await import(moduleUrl);
    return JSON.stringify(useWorkspaceStore.getState().sessions[useWorkspaceStore.getState().activeId].data.blocks).includes(finalText);
  }, { moduleUrl, finalText });
  assert.deepEqual(await model(), { text: finalText, count: before.count });
  report.checks.push('native-composition-and-confirmation');
  // 功能：一次候选确认只对应一次撤销，撤销与重做中不出现拼音中间态；时间：2026-11-07；作者：lq。
  await page.keyboard.press('ControlOrMeta+z');
  assert.deepEqual(await model(), before);
  await page.keyboard.press('ControlOrMeta+Shift+z');
  assert.deepEqual(await model(), { text: finalText, count: before.count });
  report.checks.push('composition-undo-redo');
  // 功能：覆盖 compositionend 后 DOM 才更新、且缺少最终 input 的宿主事件顺序；时间：2026-11-07；作者：lq。
  await paragraph.evaluate((element, finalText) => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    element.querySelector('span').textContent = 'qu xiao';
    element.dispatchEvent(new InputEvent('input', { bubbles: true, isComposing: true, data: 'qu xiao', inputType: 'insertCompositionText' }));
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '' }));
    element.querySelector('span').textContent = finalText;
  }, finalText);
  await page.waitForTimeout(100);
  assert.deepEqual(await model(), { text: finalText, count: before.count });
  report.checks.push('cancel-and-end-without-input');
  // 功能：最终 DOM 内容确实变化且没有 input 时，也必须经 compositionend 写回；时间：2026-11-07；作者：lq。
  await paragraph.evaluate((element, finalText) => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    element.querySelector('span').textContent = 'ce shi';
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: finalText }));
    element.querySelector('span').textContent = finalText;
  }, finalText + '测试');
  await page.waitForTimeout(100);
  assert.deepEqual(await model(), { text: finalText + '测试', count: before.count });
  // 功能：恢复目标中文，继续检查保存；时间：2026-11-07；作者：lq。
  await selectText();
  await page.keyboard.insertText(finalText);
  report.checks.push('end-without-final-input');
  const downloadPromise = page.waitForEvent('download');
  await page.keyboard.press('ControlOrMeta+s');
  const download = await downloadPromise;
  const savedPath = join(output, 'ime-saved.docx');
  await download.saveAs(savedPath);
  const saved = await readFile(savedPath);
  const savedZip = await JSZip.loadAsync(saved);
  const savedXml = await savedZip.file('word/document.xml').async('text');
  assert.ok(savedXml.includes(finalText), '保存必须包含最终中文');
  /** 功能：剔除文本节点后逐字比较，合并续格边框和其他表格属性必须完全一致；时间：2026-11-07；作者：lq。 */
  const withoutText = xml => xml.replace(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/g, '<w:t/>');
  assert.equal(withoutText(savedXml), withoutText(originalXml));
  assert.deepEqual(Object.keys(savedZip.files).filter(name => !savedZip.files[name].dir).sort(), Object.keys(originalZip.files).filter(name => !originalZip.files[name].dir).sort());
  for (const [name, entry] of Object.entries(originalZip.files)) {
    if (!entry.dir && name !== 'word/document.xml') assert.deepEqual(await savedZip.file(name).async('uint8array'), await entry.async('uint8array'), name);
  }
  report.checks.push('saved-chinese-and-all-nontext-properties');
  await page.reload();
  const reopen = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: '打开文件', exact: true }).first().click();
  await (await reopen).setFiles(savedPath);
  await page.locator('[data-paragraph-id][role="textbox"]').filter({ hasText: finalText }).first().waitFor({ state: 'visible' });
  report.checks.push('reopened-chinese');
  assert.deepEqual(errors, []);
  assert.deepEqual(await readFile(file), original, '原件必须保持不变');
  report.status = 'passed';
} catch (error) {
  report.status = 'failed'; report.error = String(error); process.exitCode = 1;
} finally {
  await browser.close();
  await writeFile(join(output, 'report.json'), JSON.stringify(report, null, 2));
  console.log(JSON.stringify(report));
}
