/** 功能：复用仓库 Playwright 运行时，在真实浏览器打开指定样本并检查字节往返、版面和运行错误；时间：2026-11-03；作者：lq。 */
import { chromium } from 'playwright-core';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { homedir } from 'node:os';

const directory = process.argv[2];
if (!directory) throw new Error('用法：node scripts/verify-word-fidelity-browser.mjs 样本目录');
const output = resolve('output/playwright');
await mkdir(output, { recursive: true });
const browser = await chromium.launch({
  executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing'),
  headless: true,
});
const report = [];
try {
  for (const name of ['采购文件.doc', '测试分包合同.doc', '测试请假单.docx']) {
    const page = await browser.newPage({ viewport: { width: 1440, height: 960 } });
    // 功能：使用标准文件输入分支，便于自动化真实的“打开文件”流程；时间：2026-11-03；作者：lq。
    await page.addInitScript(() => { window.showOpenFilePicker = undefined; });
    const errors = [];
    page.on('pageerror', (error) => errors.push(String(error)));
    await page.goto(process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1421/', { waitUntil: 'domcontentloaded' });
    // 功能：记录初始页面再调用应用的真实字节打开入口，检查浏览器端转换而非 Node 专用路径；时间：2026-11-03；作者：lq。
    await page.locator('body').innerText();
    await page.screenshot({ path: join(output, 'fidelity-initial.png') });
    const chooser = page.waitForEvent('filechooser');
    await page.getByRole('button', { name: '打开文件', exact: true }).first().click();
    await (await chooser).setFiles(join(directory, name));
    await page.waitForSelector('[data-word-editor]', { timeout: 30000 });
    const base64 = (await readFile(join(directory, name))).toString('base64');
    const state = await page.evaluate(async ({ name, base64, root }) => {
      const { useWorkspaceStore } = await import(`/@fs${root}/packages/document-model/src/store.ts`);
      const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
      const id = useWorkspaceStore.getState().activeId;
      const session = useWorkspaceStore.getState().sessions[id];
      if (!session) throw new Error(`样本打开失败：${JSON.stringify(useWorkspaceStore.getState().notifications)}`);
      const { serializeDocx } = await import(`/@fs${root}/packages/editor-word/src/serialize.ts`);
      const { convertLegacyDoc } = await import(`/@fs${root}/packages/platform-adapter/src/legacy-doc.ts`);
      const source = name.endsWith('.doc') ? (await convertLegacyDoc(bytes, name)).bytes : bytes;
      const saved = await serializeDocx(session.data, { sourceBytes: source, name: name.replace(/\.doc$/, '.docx') });
      if (saved.length !== source.length || !saved.every((value, index) => value === source[index])) throw new Error('无操作保存改变了源包');
      return { dirty: session.dirty, warnings: session.capability.warnings, sourceGroups: session.data.source?.groups.length, tableWidths: session.data.blocks.filter((b) => b.type === 'table').map((b) => b.columnWidthsPx) };
    }, { name, base64, root: process.cwd() });
    await page.waitForSelector('[data-word-editor]', { timeout: 30000 });
    await page.evaluate(() => document.fonts.ready);
    await page.waitForTimeout(1200);
    await page.screenshot({ path: join(output, `fidelity-${name.replace(/\.docx?$/, '')}.png`) });
    const visual = await page.evaluate(() => ({ pages: document.querySelectorAll('.mo-word-page').length, paragraphs: document.querySelectorAll('[data-word-editor] [data-paragraph-id]').length, fontOptions: Array.from(document.querySelectorAll('select[aria-label="字体"] option'), (e) => e.textContent) }));
    if (errors.length) throw new Error(errors.join('\n'));
    report.push({ name, ...state, ...visual, noOpByteExact: true, errors });
    await page.close();
  }
  await writeFile(join(output, 'fidelity-browser-report.json'), JSON.stringify(report, null, 2));
  console.log(JSON.stringify(report.map(({ name, pages, noOpByteExact, errors }) => ({ name, pages, noOpByteExact, errors })), null, 2));
} finally {
  await browser.close();
}
