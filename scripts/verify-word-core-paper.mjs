/** 功能：真实产品页面验证分页灰带、无框纸面及混合纸张，不修改公共原样本；时间：2026-11-11；作者：lq。 */
import { chromium } from 'playwright-core';
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { homedir } from 'node:os';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '..');
const output = join(root, 'output/playwright/word-core-m4-paper.local');
await mkdir(output, { recursive: true });
const sample = await readFile(join(root, 'tests/fixtures/docx/contract.docx'));
const paragraph = text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const portrait = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
const longBody = Array.from({ length: 75 }, (_, index) => paragraph(`第 ${index + 1} 段：分页间隔与无边框纸张显示验收。`)).join('');
/** 功能：基于公开 ZIP 外壳创建长段落、跨页表格和横纵混排样本；时间：2026-11-11；作者：lq。 */
async function fixture(body, section = portrait) {
  const zip = await JSZip.loadAsync(sample);
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}${section}</w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}
const samples = {
  contract: sample,
  paragraphs: await fixture(longBody),
  longParagraph: await fixture(paragraph('这是长段落中的实际换页，页间应显示灰色画布。'.repeat(250))),
  table: await fixture(`<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>${Array.from({ length: 80 }, (_, index) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>${paragraph(`表格第 ${index + 1} 行：分页灰带验收`)}</w:tc></w:tr>`).join('')}</w:tbl>`),
  // 功能：同一行两列长文本及不同段落高度连续跨页，检查伴随页缝不会盖字或累积错位；时间：2026-11-11；作者：lq。
  longCells: await fixture(`<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="5000"/></w:tblGrid><w:tr>${[3000, 5000].map((width, col) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${Array.from({ length: col ? 100 : 115 }, (_, row) => paragraph(`第${col + 1}列第${row + 1}段：单行多格跨页正文，必须完整显示。`)).join('')}</w:tc>`).join('')}</w:tr></w:tbl>`),
  // 功能：跨页重复表头仅为装饰，验证真实正文行不会落入页间画布；时间：2026-11-11；作者：lq。
  repeatedHeader: await fixture(`<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${Array.from({ length: 85 }, (_, row) => `<w:tr>${row === 0 ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${[1, 2].map(col => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${paragraph(row ? `第${row}行第${col}列：重复表头后正文` : `表头${col}`)}</w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>`),
  // 功能：跨页纵向合并单元格的边框不能穿过页缝，覆盖普通单列表格没有包含的绘制顺序；时间：2026-11-11；作者：lq。
  mergedTable: await fixture(`<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side => `<w:${side} w:val="single" w:sz="8" w:color="000000"/>`).join('')}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="6000"/></w:tblGrid>${Array.from({ length: 80 }, (_, index) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vMerge${index === 0 ? ' w:val="restart"' : ''}/></w:tcPr>${paragraph(index === 0 ? '跨页纵向合并' : '')}</w:tc><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>${paragraph(`合并表格第 ${index + 1} 行：页间不保留边框`)}</w:tc></w:tr>`).join('')}</w:tbl>`),
  mixed: await fixture(`${longBody}<w:p><w:pPr>${portrait}</w:pPr><w:r><w:t>横向分节</w:t></w:r></w:p>${longBody}`, portrait.replace('w:w="11906" w:h="16838"', 'w:w="16838" w:h="11906" w:orient="landscape"')),
};
// 功能：按需验收本机用户样本，文件和截图仍只写忽略目录，大文件使用紧凑字节传输；时间：2026-11-11；作者：lq。
for (const [index, path] of process.argv.slice(2).entries()) samples[`local-${index + 1}`] = await readFile(path);
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const report = { status: 'running', checks: [], samples: {}, errors: [] };
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  // 功能：保留懒加载模块的真实 URL，开发热更时仍查找同一实例而不创建第二份模块状态；时间：2026-11-11；作者：lq。
  await page.addInitScript(() => performance.setResourceTimingBufferSize(5000));
  page.setDefaultTimeout(20000);
  page.on('pageerror', error => report.errors.push(String(error)));
  await page.goto(process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1432');
  await page.evaluate(async root => {
    const { useWorkspaceStore } = await import('/@fs' + root + '/packages/document-model/src/index.ts');
    window.__paperStore = useWorkspaceStore;
  }, root);
  for (const [name, bytes] of Object.entries(samples)) {
    // 功能：修复回归时允许只重跑已失败的样本，完整验收默认仍覆盖全部；时间：2026-11-11；作者：lq。
    if (process.env.MINI_OFFICE_PAPER_SAMPLES && !process.env.MINI_OFFICE_PAPER_SAMPLES.split(',').includes(name)) continue;
    await page.evaluate(async ({ name, bytes }) => {
      const store = window.__paperStore;
      await store.getState().closeAll(true);
      await store.getState().openBytes({ name, bytes: Uint8Array.from(atob(bytes), char => char.charCodeAt(0)) });
    }, { name: process.argv.slice(2)[Number(name.replace('local-', '')) - 1]?.split('/').pop() ?? `${name}.docx`, bytes: Buffer.from(bytes).toString('base64') });
    await page.waitForFunction(() => document.querySelector('[data-word-engine="word-core"] .page-gap:not(.page-gap-carry)')?.getBoundingClientRect().height > 28);
    await page.waitForTimeout(1000);
    for (const [theme, zoom] of [['light', 1], ['light', 0.75], ['dark', 1]]) {
      // 功能：通过产品设置切换主题/缩放；截图像素检查真实灰带，而非只断言样式文本；时间：2026-11-11；作者：lq。
      await page.evaluate(({ theme, zoom }) => {
        const store = window.__paperStore.getState();
        store.updateSettings({ theme }); store.updateView(store.activeId, { zoom });
      }, { theme, zoom });
      await page.waitForTimeout(1000);
      // 功能：读取产品实际宿主而非第二份分页实例，逐行核对字形没有穿过物理纸边；时间：2026-11-11；作者：lq。
      const content = await page.evaluate(async () => {
        const store = window.__paperStore.getState(), core = store.sessions[store.activeId].wordCore;
        let surface;
        for (const resource of performance.getEntriesByType('resource').filter(r => /\/core-surface\.ts(?:\?|$)/.test(r.name)).reverse()) {
          surface = (await import(resource.name)).WordCoreSurface.current(core); if (surface) break;
        }
        if (!surface?.layout.current()) throw new Error('未找到实际文档分页宿主');
        const { frames } = surface.layout.current(), paper = core.editor.view.dom, box = paper.getBoundingClientRect(), factor = Number(getComputedStyle(paper.closest('.doc-zoom')).zoom) || 1;
        const walker = document.createTreeWalker(paper, NodeFilter.SHOW_TEXT), range = document.createRange();
        const violations = []; let node, rects = 0;
        while ((node = walker.nextNode())) {
          if (!node.textContent.trim() || node.parentElement.closest('.page-gap, .page-repeat-header, .doc-cell-boxes')) continue;
          range.selectNodeContents(node);
          for (const rect of range.getClientRects()) {
            if (rect.width < 0.1 || rect.height < 0.1) continue;
            rects++;
            const top = (rect.top - box.top) / factor, bottom = (rect.bottom - box.top) / factor;
            const frame = [...frames].reverse().find(frame => frame.top <= (top + bottom) / 2);
            if (frame && (top < frame.top - 2 || bottom > frame.top + frame.heightPx + 2)) {
              const cell = node.parentElement.closest('td,th');
              if (!violations.length) window.__paperViolation = node.parentElement;
              violations.push({ text: node.textContent.slice(0, 45), top, bottom, pageTop: frame.top, pageBottom: frame.top + frame.heightPx, cell: cell?.outerHTML.slice(0, 750), cellTop: cell && (cell.getBoundingClientRect().top - box.top) / factor, cellBottom: cell && (cell.getBoundingClientRect().bottom - box.top) / factor, align: cell && getComputedStyle(cell).verticalAlign, attrs: cell?.pmViewDesc?.node?.attrs });
            }
          }
        }
        // 功能：重复表头必须与真实首行逐列对齐，防止页缝扩列仍因纸边检查通过而漏检；时间：2026-11-11；作者：lq。
        const headerMismatches = [];
        for (const table of paper.querySelectorAll('table')) {
          const rows = [...table.rows].filter(row => !row.classList.contains('page-gap-inline'));
          for (const header of table.querySelectorAll(':scope > tbody > .mo-word-repeat-header')) {
            const source = rows[0]; if (!source) continue;
            [...header.cells].forEach((cell, index) => {
              const original = source.cells[index]; if (!original) { headerMismatches.push('extra column'); return; }
              const a = cell.getBoundingClientRect(), b = original.getBoundingClientRect();
              if (Math.abs(a.left - b.left) / factor > 1.5 || Math.abs(a.width - b.width) / factor > 1.5) headerMismatches.push({ index, left: (a.left - b.left) / factor, width: (a.width - b.width) / factor });
            });
          }
        }
        // 功能：页脚下沿锚定 DOCX 距底距离，不能把正文预留高度再次减到页面上方；时间：2026-11-11；作者：lq。
        const sections = core.sections(), footerMismatches = [];
        for (const footer of document.querySelectorAll('.mo-core-page-regions [data-hf-kind="footer"]')) {
          const frame = frames[Number(footer.dataset.hfPage)]; if (!frame) continue;
          const expected = frame.top + frame.heightPx - (sections[frame.section].settings.footerDist ?? 720) / 15;
          const actual = (footer.getBoundingClientRect().bottom - box.top) / factor;
          if (Math.abs(actual - expected) > 1.5) footerMismatches.push({ actual, expected });
        }
        return { rects, headerMismatches, footerMismatches, violations: violations.slice(0, 20), violationCount: violations.length, error: String(surface.layoutError ?? ''), ...(violations.length ? { alignments: core.editor.state.plugins.filter(p => p.key.includes('miniOfficeMergedCellAlignment')).flatMap(p => p.getState(core.editor.state).find().map(d => ({ from: d.from, to: d.to, text: core.editor.state.doc.textBetween(d.from, d.to), attrs: d.type.attrs }))), frames, cuts: surface.layout.current().cuts.map(cut => ({ lead: cut.lead, start: cut.start, gap: cut.gap })), gaps: Array.from(paper.querySelectorAll('.page-gap, .mo-word-cell-gap')).map(el => ({ type: el.className, page: el.dataset.page, top: (el.getBoundingClientRect().top - box.top) / factor, height: el.getBoundingClientRect().height / factor, shift: el.closest('.mo-word-merged-cell-text')?.outerHTML.slice(0, 200), text: el.parentElement.textContent.slice(0, 60) })) } : {}) };
      });
      report.samples[`${name}-${theme}-${zoom}-content`] = content;
      if (content.violationCount) { await page.evaluate(() => window.__paperViolation?.scrollIntoView({ block: 'center' })); await page.screenshot({ path: join(output, `${name}-${theme}-${zoom}-violation.png`) }); }
      assert.equal(content.error, '', `${name}: layout error`);
      assert.deepEqual(content.headerMismatches, [], `${name}: repeated header columns`);
      assert.deepEqual(content.footerMismatches, [], `${name}: footer bottom anchor`);
      assert.equal(content.violationCount, 0, `${name}: text crosses paper edge ${JSON.stringify(content.violations)}`);
      // 功能：本机长文档检查跨页合并表格的中段页缝，避免只检查封面后的普通段落分隔；时间：2026-11-11；作者：lq。
      const tableGaps = page.locator('[data-word-engine="word-core"] .page-gap-table');
      const tableGapCount = await tableGaps.count();
      const gap = name.startsWith('local-') && tableGapCount ? tableGaps.nth(Math.min(4, tableGapCount - 1)) : page.locator('[data-word-engine="word-core"] .page-gap:not(.page-gap-carry)').first();
      await gap.evaluate(element => element.scrollIntoView({ block: 'center' }));
      // 功能：滚动后的延迟布局/滚动锚定完成后再采集截图坐标，避免用旧坐标扫描新画面；时间：2026-11-11；作者：lq。
      await page.waitForTimeout(400);
      // 功能：超高行的补白可大于视口，定位真实灰带中心而非整个占位行中心，防止扫描视口外像素；时间：2026-11-11；作者：lq。
      await gap.evaluate(gap => {
        const scroller = gap.closest('.mo-word-canvas'), paper = gap.closest('.doc-page');
        const factor = paper.getBoundingClientRect().width / paper.offsetWidth;
        const center = gap.getBoundingClientRect().top + (parseFloat(getComputedStyle(gap).getPropertyValue('--gap-mb')) + 14) * factor;
        const viewport = scroller.getBoundingClientRect();
        scroller.scrollTop += center - (viewport.top + viewport.height / 2);
      });
      await page.waitForTimeout(150);
      const state = await gap.evaluate(gap => {
        const store = window.__paperStore.getState(), session = store.sessions[store.activeId];
        const root = document.querySelector('[data-word-engine="word-core"] .mo-word-core');
        const paper = root.querySelector('.doc-page'), wrap = root.querySelector('.page-wrap');
        const fill = gap.querySelector('.page-gap-table-fill');
        const box = gap.getBoundingClientRect(), factor = paper.getBoundingClientRect().width / paper.offsetWidth;
        const band = getComputedStyle(fill ?? gap, fill ? null : '::before');
        const mb = parseFloat(getComputedStyle(gap).getPropertyValue('--gap-mb'));
        const canvas = getComputedStyle(paper).getPropertyValue('--canvas').trim();
        const probe = document.createElement('div'); probe.style.color = canvas; document.body.append(probe);
        const color = getComputedStyle(probe).color; probe.remove();
        return { pages: session.view.pageCount, dirty: session.wordCore.state().dirty, canvas: color, bandColor: band.backgroundColor, bandImage: band.backgroundImage, bandShadow: band.boxShadow, paperShadow: getComputedStyle(paper).boxShadow, border: getComputedStyle(paper).borderTopWidth, mixed: wrap.classList.contains('paper-mixed'), isolation: getComputedStyle(wrap).isolation, sheets: Array.from(wrap.querySelectorAll('.page-sheet')).map(sheet => ({ position: getComputedStyle(sheet).position, background: getComputedStyle(sheet).backgroundColor, shadow: getComputedStyle(sheet).boxShadow, width: sheet.getBoundingClientRect().width })), point: { x: Math.round(box.left + box.width / 2), y: Math.round(box.top + (mb + 14) * factor) }, scan: { left: Math.ceil(paper.getBoundingClientRect().left + 2), right: Math.floor(paper.getBoundingClientRect().right - 2) }, factor };
      });
      const label = `${name}-${theme}-${zoom}`;
      assert.ok(state.pages > 1, label);
      assert.ok(Math.abs(state.factor - zoom) < 0.01, `${label}: actual zoom`);
      assert.equal(state.paperShadow, 'none', label);
      assert.equal(state.border, '0px', label);
      assert.equal(state.bandShadow, 'none', label);
      assert.equal(state.dirty, false, label);
      if (state.bandImage === 'none') assert.equal(state.bandColor, state.canvas, label);
      else assert.ok(state.bandImage.includes(state.canvas), label);
      if (name === 'mixed') {
        assert.equal(state.mixed, true); assert.equal(state.isolation, 'isolate');
        assert.equal(state.sheets.length, state.pages);
        assert.ok(state.sheets.every(sheet => sheet.position === 'absolute' && sheet.background === 'rgb(255, 255, 255)' && sheet.shadow === 'none'));
        assert.ok(new Set(state.sheets.map(sheet => Math.round(sheet.width))).size > 1);
      }
      const screenshot = await page.screenshot({ path: join(output, `${label}.png`) });
      const pixel = await page.evaluate(async ({ source, point, scan, expected }) => {
        const image = new Image(); image.src = source; await image.decode();
        const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
        const context = canvas.getContext('2d'); context.drawImage(image, 0, 0);
        // 功能：横向扫描整条页缝的三个高度，能捕获合并格边框穿透，避免仅中点像素假通过；时间：2026-11-11；作者：lq。
        for (const y of [point.y - 4, point.y, point.y + 4]) {
          const data = context.getImageData(scan.left, y, scan.right - scan.left, 1).data;
          for (let x = 0; x < data.length; x += 4) if (expected.some((value, channel) => value !== data[x + channel])) throw new Error(`页缝背景不连续：x=${scan.left + x / 4}, y=${y}, rgb=${Array.from(data.slice(x, x + 3))}`);
        }
        return Array.from(context.getImageData(point.x, point.y, 1, 1).data).slice(0, 3);
      }, { source: `data:image/png;base64,${screenshot.toString('base64')}`, point: state.point, scan: state.scan, expected: state.canvas.match(/\d+/g).slice(0, 3).map(Number) });
      assert.deepEqual(pixel, state.canvas.match(/\d+/g).slice(0, 3).map(Number), `${label}: gap pixel`);
      report.samples[label] = { ...state, pixel }; report.checks.push(label);
    }
  }
  assert.deepEqual(report.errors, []); report.status = 'passed'; report.browser = browser.version();
} catch (error) { report.status = 'failed'; report.error = String(error); report.stack = error.stack; process.exitCode = 1; }
finally { await browser.close(); await writeFile(join(output, 'report.json'), JSON.stringify(report, null, 2)); console.log(JSON.stringify(report)); }
