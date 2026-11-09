/** 功能：在真实 Mini Office 页面验收 M2 开发会话的输入、工具栏、下载保存与标签生命周期；时间：2026-11-09；作者：lq。 */
import { chromium } from 'playwright-core';
import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { homedir } from 'node:os';
import JSZip from 'jszip';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const output = join(root, 'output/playwright/word-core-m2-workspace.local');
await mkdir(output, { recursive: true });
const origin = process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1431';
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const report = { status: 'running', checks: [] };
const errors = [];
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true });
  page.setDefaultTimeout(20000);
  page.on('pageerror', error => errors.push(String(error)));
  // 功能：覆盖不提供文件系统选择器的浏览器环境，验证真实下载保存路径；不模拟解析、编辑或写盘结果；时间：2026-11-09；作者：lq。
  await page.addInitScript(() => { window.showSaveFilePicker = undefined; });
  await page.goto(origin);
  await page.evaluate(async url => { const { useWorkspaceStore } = await import(url); window.__workspace = useWorkspaceStore; }, `${origin}/@fs${root}/packages/document-model/src/index.ts`);
  const first = await page.evaluate(async () => {
    const id = await window.__workspace.getState().createDocument('docx');
    window.__firstCore = window.__workspace.getState().sessions[id].wordCore;
    return id;
  });
  await page.locator('[data-word-engine="word-core"] .ProseMirror').waitFor();
  await page.locator('[data-word-engine="word-core"] .ProseMirror').click();
  await page.keyboard.insertText('新建文档使用 GenOffice 定制内核。');
  await page.getByRole('button', { name: '加粗', exact: true }).click();
  await page.keyboard.insertText('加粗中文验收');
  assert.equal(await page.evaluate(() => window.__firstCore.editor.isActive('bold')), true);
  const pending = page.waitForEvent('download');
  await page.getByRole('button', { name: '保存', exact: true }).click();
  const downloaded = await pending;
  const savedPath = join(output, 'new-saved.docx');
  await downloaded.saveAs(savedPath);
  await page.waitForFunction(id => !window.__workspace.getState().sessions[id].dirty, first);
  const savedBytes = await readFile(savedPath);
  const savedZip = await JSZip.loadAsync(savedBytes);
  const xml = await savedZip.file('word/document.xml').async('text');
  assert.ok(xml.includes('加粗中文验收')); assert.match(xml, /<w:b\b/);
  report.checks.push('actual-app-new-core-editor', 'toolbar-bold-and-chinese-edit', 'actual-browser-download-save');
  // 功能：打开真实保存副本触发 React 标签切换，旧会话在界面卸载后仍可保存与撤销；时间：2026-11-09；作者：lq。
  const second = await page.evaluate(bytes => window.__workspace.getState().openBytes({ name: '重开.docx', bytes: new Uint8Array(bytes) }), Array.from(savedBytes));
  await page.waitForFunction(id => window.__workspace.getState().activeId === id && !!document.querySelector('[data-word-engine="word-core"] .ProseMirror'), second);
  const switched = await page.evaluate(first => {
    const core = window.__workspace.getState().sessions[first].wordCore;
    const active = window.__workspace.getState().sessions[window.__workspace.getState().activeId].wordCore;
    return { retained: core === window.__firstCore && !core.editor.isDestroyed, reopened: active.editor.getText() };
  }, first);
  assert.equal(switched.retained, true); assert.ok(switched.reopened.includes('加粗中文验收'));
  await page.evaluate(first => window.__workspace.getState().activateTab(first), first);
  await page.waitForFunction(() => document.querySelector('[data-word-engine="word-core"] .ProseMirror') === window.__firstCore.editor.view.dom);
  assert.equal(await page.getByRole('button', { name: '撤销', exact: true }).isEnabled(), true);
  await page.getByRole('button', { name: '撤销', exact: true }).click();
  assert.equal(await page.evaluate(() => window.__firstCore.editor.getText().includes('加粗中文验收')), false);
  await page.getByRole('button', { name: '重做', exact: true }).click();
  assert.equal(await page.evaluate(() => window.__firstCore.editor.getText().includes('加粗中文验收')), true);
  report.checks.push('reopen-through-real-workspace', 'tab-unmount-retains-session-and-history');
  // 功能：浏览器实际中文组合输入期间保存不能写出候选拼音；时间：2026-11-09；作者：lq。
  await page.locator('[data-word-engine="word-core"] .ProseMirror').click();
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Input.imeSetComposition', { text: 'shu ru fa', selectionStart: 8, selectionEnd: 8 });
  assert.equal(await page.evaluate(first => window.__workspace.getState().saveSession(first), first), false);
  await cdp.send('Input.insertText', { text: '输入法确认' });
  await page.waitForFunction(() => !window.__firstCore.editor.view.composing && window.__firstCore.editor.getText().includes('输入法确认'));
  report.checks.push('actual-app-ime-save-guard');
  // 功能：自动恢复快照经真实 IndexedDB 保存 DOCX 字节，不序列化编辑器对象；时间：2026-11-09；作者：lq。
  await page.evaluate(() => window.__workspace.getState().updateSettings({ autoSave: true, autoSaveDelayMs: 500 }));
  await page.keyboard.insertText('自动恢复快照验收');
  assert.equal(await page.evaluate(() => window.__firstCore.editor.getText().includes('自动恢复快照验收')), true);
  // 功能：异步快照读取显式轮询实际字节；waitForFunction 的同步谓词不能用 Promise 当完成条件；时间：2026-11-09；作者：lq。
  const snapshotBytes = await page.evaluate(async ({ first, root }) => {
    const module = await import('/@fs' + root + '/packages/platform-adapter/src/index.ts');
    const { readDocxPartBase64 } = await import('/@fs' + root + '/packages/word-core/src/index.ts');
    const deadline = Date.now() + 20000;
    while (Date.now() < deadline) {
      const snapshot = await module.getPlatformAdapter().loadSnapshot(first);
      if (snapshot?.dirty) {
        const encoded = await readDocxPartBase64(snapshot.bytes, 'word/document.xml');
        if (encoded) {
          const xml = new TextDecoder().decode(Uint8Array.from(atob(encoded), char => char.charCodeAt(0)));
          const document = new DOMParser().parseFromString(xml, 'application/xml');
          const text = Array.from(document.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 't')).map(node => node.textContent).join('');
          if (text.includes('自动恢复快照验收')) return Array.from(snapshot.bytes);
        }
      }
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    throw new Error('自动恢复快照在期限内没有包含最新事务');
  }, { first, root });
  const snapshot = await JSZip.loadAsync(new Uint8Array(snapshotBytes));
  const snapshotXml = await snapshot.file('word/document.xml').async('text');
  await writeFile(join(output, 'snapshot.xml'), snapshotXml);
  await writeFile(join(output, 'snapshot.docx'), new Uint8Array(snapshotBytes));
  const debug = await page.evaluate(async first => {
    const core = window.__workspace.getState().sessions[first].wordCore;
    const prepared = await core.prepareSave();
    return { state: core.state(), text: core.editor.getText(), json: core.editor.getJSON(), bytes: Array.from(prepared.bytes) };
  }, first);
  await writeFile(join(output, 'session.json'), JSON.stringify({ state: debug.state, text: debug.text, json: debug.json }, null, 2));
  await writeFile(join(output, 'latest.docx'), new Uint8Array(debug.bytes));
  assert.ok([...snapshotXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map(match => match[1]).join('').includes('自动恢复快照验收'), snapshotXml);
  report.checks.push('real-indexeddb-docx-recovery-snapshot');
  // 功能：已保存文档继续输入后再次真实下载，随后重开再编辑并第三次保存；时间：2026-11-09；作者：lq。
  const secondDownload = page.waitForEvent('download');
  await page.getByRole('button', { name: '保存', exact: true }).click();
  const secondPath = join(output, 'continued-saved.docx');
  await (await secondDownload).saveAs(secondPath);
  const continued = await readFile(secondPath);
  const third = await page.evaluate(bytes => window.__workspace.getState().openBytes({ name: '再次重开.docx', bytes: new Uint8Array(bytes) }), Array.from(continued));
  await page.waitForFunction(id => window.__workspace.getState().activeId === id && document.querySelector('[data-word-engine="word-core"] .ProseMirror') === window.__workspace.getState().sessions[id].wordCore?.editor.view.dom, third);
  await page.locator('[data-word-engine="word-core"] .ProseMirror').click();
  await page.keyboard.insertText('重开后再编辑');
  const thirdDownload = page.waitForEvent('download');
  await page.getByRole('button', { name: '保存', exact: true }).click();
  const thirdPath = join(output, 'reopened-edited-saved.docx');
  await (await thirdDownload).saveAs(thirdPath);
  const thirdZip = await JSZip.loadAsync(await readFile(thirdPath));
  const thirdXml = await thirdZip.file('word/document.xml').async('text');
  const thirdText = [...thirdXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map(match => match[1]).join('');
  for (const expected of ['输入法确认', '自动恢复快照验收', '重开后再编辑']) assert.ok(thirdText.includes(expected), expected);
  report.checks.push('continued-save-reopen-edit-and-third-save');
  await page.screenshot({ path: join(output, 'workspace.png') });
  assert.deepEqual(errors, []);
  report.status = 'passed'; report.browser = browser.version();
} catch (error) { report.status = 'failed'; report.error = String(error); process.exitCode = 1; }
finally { await browser.close(); await writeFile(join(output, 'report.json'), JSON.stringify({ ...report, errors }, null, 2)); console.log(JSON.stringify({ ...report, errors })); }
