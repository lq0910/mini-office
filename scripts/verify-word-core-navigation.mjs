/** 功能：真实产品验证目录单击/书签导航和 Ctrl+滚轮缩放，支持显式传入的本机只读样本；时间：2026-11-11；作者：lq。 */
import { chromium } from 'playwright-core';
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import JSZip from 'jszip';
const root = resolve(import.meta.dirname, '..'), output = join(root, 'output/playwright/word-core-navigation.local');
await mkdir(output, { recursive: true });
const zip = await JSZip.loadAsync(await readFile(join(root, 'tests/fixtures/docx/contract.docx')));
// 功能：导入目录引用普通段落的隐藏书签，缓存页码故意过期，验证导航按真实目标而非页码；时间：2026-11-11；作者：lq。
const run = text => `<w:r><w:t>${text}</w:t></w:r>`;
zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-3" \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="_TocTarget">${run('单击跳转测试章')}<w:r><w:tab/></w:r>${run('99')}</w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>${Array.from({ length: 70 }, (_, i) => `<w:p>${run(`正文第 ${i} 段`)}</w:p>`).join('')}<w:p><w:bookmarkStart w:id="1" w:name="_TocTarget"/>${run('书签目标正文，无标题样式')}<w:bookmarkEnd w:id="1"/></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
const samples = [{ name: '目录书签验收.docx', bytes: await zip.generateAsync({ type: 'uint8array' }) }];
for (const path of process.argv.slice(2)) samples.push({ name: basename(path), bytes: await readFile(path) });
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const report = { status: 'running', samples: [], errors: [] };
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  page.on('pageerror', error => report.errors.push(String(error))); page.setDefaultTimeout(20000);
  await page.goto(process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1432');
  await page.evaluate(async root => { window.__navStore = (await import('/@fs' + root + '/packages/document-model/src/index.ts')).useWorkspaceStore; }, root);
  for (const sample of samples) {
    await page.evaluate(async sample => { const store = window.__navStore.getState(); await store.closeAll(true); await store.openBytes({ name: sample.name, bytes: Uint8Array.from(atob(sample.bytes), ch => ch.charCodeAt(0)) }); }, { name: sample.name, bytes: Buffer.from(sample.bytes).toString('base64') });
    await page.locator('.doc-toc-line').first().waitFor(); await page.waitForTimeout(2000);
    const count = await page.locator('.doc-toc-line').count(), checks = [];
    // 功能：失败时保留具体目录条目和目标，便于区分链接缺失与滚动异常；时间：2026-11-11；作者：lq。
    const sampleReport = { name: sample.name, entries: count, checks };
    report.samples.push(sampleReport);
    for (const index of Array.from({ length: count }, (_, index) => index)) {
      const line = page.locator('.doc-toc-line').nth(index), title = await line.innerText();
      await line.click(); await page.waitForTimeout(150);
      const result = await page.evaluate(() => { const st = window.__navStore.getState(), s = st.sessions[st.activeId], e = s.wordCore.editor; return { text: e.state.selection.$from.parent.textContent, position: e.state.selection.from, dirty: s.wordCore.state().dirty, targetTop: e.view.nodeDOM(e.state.selection.$from.before(1))?.getBoundingClientRect().top, viewportTop: document.querySelector('.mo-word-canvas').getBoundingClientRect().top, viewportBottom: document.querySelector('.mo-word-canvas').getBoundingClientRect().bottom, error: document.querySelector('[role=alert]')?.textContent }; });
      checks.push({ title, ...result });
      assert.ok(!result.error, result.error); assert.ok(result.position > 1); assert.equal(result.dirty, false); assert.ok(result.targetTop >= result.viewportTop - 2 && result.targetTop < result.viewportBottom, 'target is visible');
      if (sample === samples[0]) assert.ok(result.text.includes('书签目标正文'));
    }
    // 功能：真实鼠标滚轮事件覆盖方向、比例同步、文档 dirty 和浏览器整窗比例；时间：2026-11-11；作者：lq。
    await page.locator('.mo-word-canvas').hover();
    await page.keyboard.down('Control'); await page.mouse.wheel(0, -120); await page.keyboard.up('Control');
    await page.waitForFunction(() => { const s = window.__navStore.getState(); return s.sessions[s.activeId].view.zoom === 1.1; });
    const enlarged = await page.evaluate(() => ({ zoom: getComputedStyle(document.querySelector('.doc-zoom')).zoom, scale: window.visualViewport.scale, dirty: window.__navStore.getState().sessions[window.__navStore.getState().activeId].wordCore.state().dirty }));
    assert.equal(Number(enlarged.zoom), 1.1); assert.equal(enlarged.scale, 1); assert.equal(enlarged.dirty, false);
    await page.keyboard.down('Control'); await page.mouse.wheel(0, 120); await page.keyboard.up('Control');
    await page.waitForFunction(() => { const s = window.__navStore.getState(); return s.sessions[s.activeId].view.zoom === 1; });
    await page.mouse.wheel(0, 120); await page.waitForTimeout(100);
    assert.equal(await page.evaluate(() => { const s = window.__navStore.getState(); return s.sessions[s.activeId].view.zoom; }), 1);
    Object.assign(sampleReport, { enlarged, normalWheelPreservesZoom: true });
  }
  assert.deepEqual(report.errors, []); report.status = 'passed';
} catch (error) { report.status = 'failed'; report.error = String(error); report.stack = error.stack; process.exitCode = 1; }
finally { await browser.close(); await writeFile(join(output, 'report.json'), JSON.stringify(report, null, 2)); console.log(JSON.stringify(report)); }
