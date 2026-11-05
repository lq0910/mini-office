/** 功能：从只读用户文件建立逐页浏览器截图、布局和无操作保存报告；私有产物仅写入 .local 目录；时间：2026-11-05；作者：lq。 */
import { chromium } from 'playwright-core';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { resolve, join, basename } from 'node:path';
import { homedir } from 'node:os';

const [directory, ...names] = process.argv.slice(2);
if (!directory || !names.length || names.some((name) => basename(name) !== name)) {
  throw new Error('用法：node scripts/verify-word-user-fidelity.mjs 样本目录 文件名...');
}
const root = process.cwd();
const output = resolve(process.env.MINI_OFFICE_USER_FIDELITY_OUTPUT ?? 'output/word-user-fidelity.local/browser');
if (!output.split('/').some((part) => part.endsWith('.local'))) throw new Error('私人验收输出必须位于 .local 目录');
await mkdir(output, { recursive: true });
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME
  ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const report = { generatedAt: new Date().toISOString(), browser: browser.version(), samples: [] };
try {
  for (const name of names) {
    const input = await readFile(join(directory, name));
    const sample = { name, sourceSha256: hash(input), pages: [] };
    report.samples.push(sample);
    const page = await browser.newPage({ viewport: { width: 1440, height: 1800 }, deviceScaleFactor: 1, colorScheme: 'light' });
    const errors = [];
    page.on('pageerror', (error) => errors.push(String(error)));
    try {
      await page.addInitScript(() => { window.showOpenFilePicker = undefined; });
      await page.goto(process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:4196/', { waitUntil: 'domcontentloaded' });
      const chooser = page.waitForEvent('filechooser');
      await page.getByRole('button', { name: '打开文件', exact: true }).first().click();
      await (await chooser).setFiles(join(directory, name));
      await page.waitForSelector('[data-word-editor]', { timeout: 60000 });
      const opened = await page.evaluate(async ({ root, name, encoded }) => {
        const { useWorkspaceStore } = await import(`/@fs${root}/packages/document-model/src/index.ts`);
        const { serializeDocx } = await import(`/@fs${root}/packages/editor-word/src/serialize.ts`);
        const { convertLegacyDoc } = await import(`/@fs${root}/packages/platform-adapter/src/legacy-doc.ts`);
        const input = Uint8Array.from(atob(encoded), (char) => char.charCodeAt(0));
        const state = useWorkspaceStore.getState();
        const session = state.sessions[state.activeId];
        if (!session || session.kind !== 'docx') throw new Error('用户样本未进入 Word 编辑器');
        state.updateView(session.id, { zoom: 1 });
        const source = /\.doc$/i.test(name) ? (await convertLegacyDoc(input, name)).bytes : input;
        const saved = await serializeDocx(session.data, { sourceBytes: source, name: name.replace(/\.doc$/i, '.docx') });
        return { model: session.data, warnings: session.capability.warnings, dirty: session.dirty,
          source: btoa(Array.from(source, (byte) => String.fromCharCode(byte)).join('')),
          noOpByteExact: saved.length === source.length && saved.every((byte, index) => byte === source[index]) };
      }, { root, name, encoded: input.toString('base64') });
      if (!opened.noOpByteExact) throw new Error('无操作保存改变源包');
      sample.noOpByteExact = true;
      sample.dirty = opened.dirty;
      sample.warnings = opened.warnings;
      await writeFile(join(output, `${name}.model.json`), JSON.stringify(opened.model, null, 2));
      await writeFile(join(output, `${name}.converted.docx`), Buffer.from(opened.source, 'base64'));
      // 功能：采用与公开语料一致的白纸、缩放及隐藏辅助标记规则，保留实际文字和布局；时间：2026-11-05；作者：lq。
      await page.evaluate(() => { document.documentElement.dataset.theme = 'light'; });
      await page.addStyleTag({ content: '.mo-word-page{box-shadow:none!important;background:white!important;--mo-paper:#fff;--mo-paper-ink:#000} *{caret-color:transparent!important} [data-margin-guides],.mo-word-page [data-tab] [aria-hidden],.mo-word-page [data-line-break]{visibility:hidden!important}' });
      await page.evaluate(async () => { await document.fonts.ready; await Promise.all([...document.querySelectorAll('.mo-word-page img')].map((image) => image.decode())); });
      let previous = ''; let stable = 0;
      for (let attempt = 0; attempt < 100 && stable < 5; attempt++) {
        const current = await page.locator('.mo-word-page').evaluateAll((pages) => JSON.stringify(pages.map((paper) => [paper.offsetWidth, paper.offsetHeight, paper.textContent, paper.innerHTML.length])));
        stable = current === previous ? stable + 1 : 0; previous = current; await page.waitForTimeout(150);
      }
      if (stable < 5) throw new Error('用户样本分页未稳定');
      const pages = page.locator('.mo-word-page');
      sample.pageCount = await pages.count();
      for (let index = 0; index < sample.pageCount; index++) {
        const paper = pages.nth(index);
        const file = `${name}-page-${String(index + 1).padStart(3, '0')}.png`;
        const layout = await paper.evaluate((element) => ({ width: element.offsetWidth, height: element.offsetHeight,
          // 功能：记录真实表格边框、内边距与行盒，以原尺寸几何核对窄列换行；时间：2026-11-05；作者：lq。
          tables: [...element.querySelectorAll('[data-table-id]')].map((table) => {
            const rect = table.getBoundingClientRect(); const paper = element.getBoundingClientRect();
            return { leftPx: rect.left - paper.left, topPx: rect.top - paper.top, widthPx: rect.width,
              cells: [...table.querySelectorAll('[data-cell-id]')].map((cell) => {
                const style = getComputedStyle(cell); const paragraph = cell.querySelector('[data-paragraph-id]');
                return { text: cell.textContent, paddingLeftPx: Number.parseFloat(style.paddingLeft), paddingRightPx: Number.parseFloat(style.paddingRight),
                  paragraphHeightPx: paragraph?.getBoundingClientRect().height, lineHeightPx: paragraph ? Number.parseFloat(getComputedStyle(paragraph).lineHeight) : undefined };
              }) };
          }),
          // 功能：记录实际文本框字形位置和逐页域值，页码出现不等于锚定位置正确；时间：2026-11-05；作者：lq。
          textBoxes: [...element.querySelectorAll('[data-header-footer-text-box]')].map((box) => {
            const rect = box.getBoundingClientRect(); const paper = element.getBoundingClientRect();
            return { text: box.textContent, pageNumber: Number(box.getAttribute('data-page-number')),
              centerGapPx: (rect.left + rect.right - paper.left - paper.right) / 2,
              topPx: rect.top - paper.top, widthPx: rect.width, heightPx: rect.height };
          }),
          // 功能：跨页段落 DOM 含被裁掉的全文，报告只能提取裁剪窗口中可见的字符，避免把重复 DOM 当成重复正文；时间：2026-11-05；作者：lq。
          visibleText: (() => {
            const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
            const texts = []; let node;
            while ((node = walker.nextNode())) {
              if (getComputedStyle(node.parentElement).visibility === 'hidden') continue;
              const clip = { left: -Infinity, right: Infinity, top: -Infinity, bottom: Infinity };
              for (let parent = node.parentElement; parent; parent = parent.parentElement) {
                const style = getComputedStyle(parent);
                if (parent === element || /(hidden|clip)/.test(style.overflowY + style.overflowX)) {
                  const rect = parent.getBoundingClientRect();
                  clip.left = Math.max(clip.left, rect.left); clip.right = Math.min(clip.right, rect.right);
                  clip.top = Math.max(clip.top, rect.top); clip.bottom = Math.min(clip.bottom, rect.bottom);
                }
                if (parent === element) break;
              }
              const range = document.createRange(); range.selectNodeContents(node);
              const box = range.getBoundingClientRect();
              if (box.top >= clip.top && box.bottom <= clip.bottom && box.left >= clip.left && box.right <= clip.right) {
                texts.push(node.textContent); continue;
              }
              for (let index = 0; index < node.length; index++) {
                range.setStart(node, index); range.setEnd(node, index + 1);
                const rect = range.getBoundingClientRect(); const x = (rect.left + rect.right) / 2; const y = (rect.top + rect.bottom) / 2;
                if (x >= clip.left && x <= clip.right && y >= clip.top && y <= clip.bottom) texts.push(node.textContent[index]);
              }
            }
            return texts.join('');
          })(),
          text: element.textContent, paragraphs: [...element.querySelectorAll('[data-paragraph-id]')].map((paragraph) => {
            const box = paragraph.getBoundingClientRect(); const pageBox = element.getBoundingClientRect();
            const style = getComputedStyle(paragraph);
            return { id: paragraph.dataset.paragraphId, text: paragraph.textContent, top: box.top - pageBox.top,
              height: box.height, font: style.fontFamily, fontSize: style.fontSize, lineHeight: style.lineHeight };
          }) }));
        await paper.screenshot({ path: join(output, file), animations: 'disabled', caret: 'hide' });
        sample.pages.push({ number: index + 1, file, ...layout });
      }
      if (errors.length) throw new Error(errors.join('\n'));
      if (hash(await readFile(join(directory, name))) !== sample.sourceSha256) throw new Error('原文件哈希发生变化');
      sample.status = 'captured';
    } catch (error) { sample.status = 'failed'; sample.error = String(error); process.exitCode = 1; }
    finally { await page.close(); }
    console.log(JSON.stringify({ name, status: sample.status, pages: sample.pageCount, noOpByteExact: sample.noOpByteExact, error: sample.error }));
  }
} finally {
  await browser.close();
  await writeFile(join(output, 'report.json'), JSON.stringify(report, null, 2));
}
