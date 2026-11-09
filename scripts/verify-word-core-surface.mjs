/** 功能：用真实浏览器验收 M2 产品正文宿主的分页、中文事务和保存重开；时间：2026-11-09；作者：lq。 */
import { chromium } from 'playwright-core';
import JSZip from 'jszip';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { homedir } from 'node:os';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const output = join(root, 'output/playwright/word-core-m2.local');
await mkdir(output, { recursive: true });
const require = createRequire(join(root, 'package.json'));
const { build } = createRequire(require.resolve('vite/package.json'))('esbuild');
// 功能：只打包真实产品宿主及其核心，不复制测试编辑器或以旧模型替代；时间：2026-11-09；作者：lq。
await build({ absWorkingDir: root, entryPoints: { surface: 'packages/editor-word/src/core-surface.ts', styles: 'packages/word-core/src/styles/document.css', fonts: 'packages/word-core/src/styles/fonts.css' }, outdir: output, bundle: true, splitting: true, format: 'esm', platform: 'browser', target: 'es2022', loader: { '.ttf': 'file', '.otf': 'file', '.woff': 'file', '.woff2': 'file' } });
const zip = await JSZip.loadAsync(await readFile(join(root, 'tests/fixtures/docx/contract.docx')));
const paragraph = text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const sections = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
/** 功能：公共样本只提供 ZIP 结构，验收正文由脚本构造，不修改任何用户文件；时间：2026-11-09；作者：lq。 */
async function fixture(body, sectionXml = sections, parts = []) {
  const copy = zip.clone();
  copy.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word"><w:body>${body}${sectionXml}</w:body></w:document>`);
  let relationships = await copy.file('word/_rels/document.xml.rels').async('text');
  let types = await copy.file('[Content_Types].xml').async('text');
  for (const part of parts) {
    copy.file(`word/${part.name}.xml`, part.xml);
    relationships = relationships.replace('</Relationships>', `<Relationship Id="m2-${part.name}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${part.kind}" Target="${part.name}.xml"/></Relationships>`);
    types = types.replace('</Types>', `<Override PartName="/word/${part.name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${part.kind}+xml"/></Types>`);
  }
  copy.file('word/_rels/document.xml.rels', relationships); copy.file('[Content_Types].xml', types);
  return Array.from(await copy.generateAsync({ type: 'uint8array' }));
}
// 功能：复杂页面样本保留公共包结构，覆盖 M2 的合并格、嵌套表格、浮动框和附属部件；时间：2026-11-09；作者：lq。
const longBody = Array.from({ length: 80 }, (_, index) => paragraph(`第${index + 1}段：智慧物业信息系统部署实施，验证同一核心分页与保存。`)).join('');
const cellTable = text => `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>${text}</w:tc></w:tr></w:tbl>`;
const fixtures = {
  paragraphs: await fixture(Array.from({ length: 80 }, (_, index) => paragraph(`第${index + 1}段：智慧物业信息系统部署实施，验证同一核心分页与保存。`)).join('')),
  longParagraph: await fixture(paragraph('跨页正文连续编辑与中文保存验证。'.repeat(280))),
  mergedCell: await fixture(`<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>${Array.from({ length: 70 }, (_, index) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/>${index === 26 ? '<w:vMerge w:val="restart"/>' : index === 27 ? '<w:vMerge/>' : ''}</w:tcPr>${paragraph(index === 27 ? '' : `第${index + 1}行，合并单元格分页验收`)}</w:tc></w:tr>`).join('')}</w:tbl>`),
  nestedTable: await fixture(cellTable(paragraph('嵌套表格之前') + cellTable(longBody) + paragraph('嵌套表格之后'))),
  multiSection: await fixture(longBody.slice(0, longBody.length) + `<w:p><w:pPr>${sections}</w:pPr><w:r><w:t>分节边界</w:t></w:r></w:p>` + longBody, sections.replace('w:w="11906" w:h="16838"', 'w:w="16838" w:h="11906" w:orient="landscape"')),
  columns: await fixture(longBody + longBody, sections.replace('</w:sectPr>', '<w:cols w:num="2" w:space="720"/></w:sectPr>')),
  floatingBox: await fixture(`<w:p><w:r><w:pict><v:shape id="m2-float" style="position:absolute;margin-left:50pt;margin-top:30pt;width:100pt;height:60pt;z-index:1"><v:textbox><w:txbxContent>${paragraph('浮动文本框保留内容')}</w:txbxContent></v:textbox><w10:wrap type="square"/></v:shape></w:pict></w:r></w:p>${longBody}`),
  headerFootnote: await fixture(`<w:p><w:r><w:t>脚注引用正文</w:t><w:footnoteReference w:id="1"/></w:r></w:p>${longBody}`, sections.replace('</w:sectPr>', '<w:headerReference w:type="default" r:id="m2-header1"/><w:footerReference w:type="default" r:id="m2-footer1"/></w:sectPr>'), [
    { name: 'header1', kind: 'header', xml: `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${paragraph('定制内核页眉验收')}</w:hdr>` },
    { name: 'footer1', kind: 'footer', xml: '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>' },
    { name: 'footnotes', kind: 'footnotes', xml: `<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1">${paragraph('页底脚注验收内容')}</w:footnote></w:footnotes>` },
  ]),
  endnotes: await fixture(`<w:p><w:r><w:t>尾注引用正文</w:t><w:endnoteReference w:id="1"/></w:r></w:p>${longBody}`, sections, [
    { name: 'endnotes', kind: 'endnotes', xml: `<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="1"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t>文末尾注验收内容</w:t></w:r></w:p></w:endnote></w:endnotes>` },
  ]),
  table: await fixture(`<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>${Array.from({ length: 70 }, (_, index) => `<w:tr>${index === 0 ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}<w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>${paragraph(`表格第${index + 1}行，验证跨页行定位。`)}</w:tc></w:tr>`).join('')}</w:tbl>`),
};
const origin = process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1431';
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const errors = [];
const report = { status: 'running', checks: [], samples: {} };
try {
  const page = await browser.newPage({ viewport: { width: 1100, height: 900 } });
  page.on('pageerror', error => errors.push(String(error)));
  await page.route('**/__word_core_m2__', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"></head><body style="margin:0;background:#eee"><div id="core" class="mo-word-core"></div></body></html>' }));
  await page.goto(`${origin}/__word_core_m2__`);
  const base = `${origin}/@fs${output}`;
  await page.addStyleTag({ url: `${base}/styles.css` });
  await page.addStyleTag({ url: `${base}/fonts.css` });
  await page.evaluate(async url => { window.__surfaceModule = await import(url); }, `${base}/surface.js`);
  for (const [name, bytes] of Object.entries(fixtures)) {
    const result = await page.evaluate(async bytes => {
      window.__surface?.dispose();
      window.__layoutErrors = [];
      const surface = await window.__surfaceModule.WordCoreSurface.mount(document.getElementById('core'), new Uint8Array(bytes), { onLayoutError: error => window.__layoutErrors.push(String(error)) });
      window.__surface = surface;
      await document.fonts.ready;
      // 功能：栏宽及分节宽度需要同源的后续测量；要求在有界轮数内收敛，再验证不会漂移；时间：2026-11-09；作者：lq。
      const settling = [];
      for (let pass = 0; pass < 10; pass++) {
        surface.layout.flush(); settling.push(surface.layout.current()?.pageCount);
        if (settling.length >= 3 && settling.slice(-3).every(count => count === settling.at(-1))) break;
      }
      if (!settling.slice(-3).every(count => count === settling.at(-1))) throw new Error(`分页未在有界轮数内收敛：${settling.join(',')}`);
      const initial = surface.layout.current();
      const noOp = await surface.session.prepareSave();
      // 功能：反复重排不能累加页间空白，不能把显示装饰记入 dirty 或保存字节；时间：2026-11-09；作者：lq。
      const counts = [];
      for (let pass = 0; pass < 3; pass++) { surface.layout.flush(); counts.push(surface.layout.current()?.pageCount); }
      const paperRect = surface.session.editor.view.dom.getBoundingClientRect();
      return { pages: initial?.pageCount, counts, gaps: document.querySelectorAll('.page-gap:not(.page-gap-carry)').length, overlays: document.querySelectorAll('.page-gap-cut').length, endnoteText: document.querySelector('.mo-core-endnotes')?.textContent, endnoteMarker: document.querySelector('.mo-core-endnotes sup')?.textContent, headers: document.querySelectorAll('.mo-core-header').length, notes: document.querySelectorAll('.page-gap-note').length, repeatedHeaders: document.querySelectorAll('.page-repeat-header').length, paperHeight: paperRect.height, sheetEdges: Array.from(document.querySelectorAll('.page-gap:not(.page-gap-carry)')).map(gap => ({ top: gap.getBoundingClientRect().top - paperRect.top, height: gap.getBoundingClientRect().height, mt: getComputedStyle(gap).getPropertyValue('--gap-mt'), mb: getComputedStyle(gap).getPropertyValue('--gap-mb') })), dirty: surface.session.state().dirty, sameBytes: bytes.every((byte, index) => noOp.bytes[index] === byte) && noOp.bytes.length === bytes.length, errors: window.__layoutErrors };
    }, bytes);
    assert.ok(result.pages > 1, `${name}: expected real multi-page layout`);
    assert.deepEqual(result.errors, [], `${name}: layout error`);
    assert.ok(result.counts.every(count => count === result.pages), `${name}: unstable page count: ${JSON.stringify(result)}`);
    assert.equal(result.gaps + result.overlays, result.pages - 1, `${name}: missing page decoration`);
    // 功能：按真实末页页首验证完整纸张高度，不把前面各页的行盒溢出误算成末页高度；时间：2026-11-09；作者：lq。
    const lastEdge = result.sheetEdges[result.sheetEdges.length - 1];
    if (lastEdge && ['paragraphs', 'longParagraph', 'table'].includes(name)) {
      const lastPaperTop = lastEdge.top + lastEdge.height - parseFloat(lastEdge.mt);
      assert.ok(Math.abs(result.paperHeight - lastPaperTop - 16838 / 15) < 1, `${name}: last page is not a full sheet: ${JSON.stringify(result)}`);
    }
    if (name === 'endnotes') { assert.ok(result.endnoteText?.includes('文末尾注验收内容')); assert.equal(result.endnoteMarker, 'i'); }
    if (name === 'headerFootnote') { assert.equal(result.headers, result.pages); assert.equal(result.notes, 1); }
    if (name === 'table') assert.equal(result.repeatedHeaders, result.pages - 1, 'table: missing repeated header');
    assert.equal(result.dirty, false);
    assert.equal(result.sameBytes, true);
    report.samples[name] = result;
    await page.screenshot({ path: join(output, `${name}.png`), fullPage: true });
    report.checks.push(`${name}:real-pagination-stable-and-lossless`);
    if (name === 'paragraphs') {
      // 功能：跨实际页界选择并粘贴，断言真实事务内容与保存结果，再撤销恢复原文；时间：2026-11-09；作者：lq。
      const pasted = await page.evaluate(async () => {
        const surface = window.__surface;
        const editor = surface.session.editor;
        const original = editor.getText();
        const boundary = surface.layout.current().cuts[1].lead.block;
        const blocks = [];
        editor.state.doc.forEach((node, pos) => blocks.push({ node, pos }));
        const start = blocks[Math.max(0, boundary - 1)];
        const end = blocks[boundary + 1];
        editor.commands.setTextSelection({ from: start.pos + 1, to: end.pos + end.node.nodeSize - 1 });
        const clipboardData = new DataTransfer();
        clipboardData.setData('text/plain', '跨页粘贴验证');
        editor.view.dom.dispatchEvent(new ClipboardEvent('paste', { clipboardData, bubbles: true, cancelable: true }));
        const contains = editor.getText().includes('跨页粘贴验证');
        const saved = await surface.session.prepareSave();
        editor.commands.undo();
        const restored = editor.getText() === original;
        surface.layout.flush();
        return { contains, restored, bytes: saved.bytes.length, pages: surface.layout.current().pageCount };
      });
      assert.equal(pasted.contains, true); assert.equal(pasted.restored, true); assert.ok(pasted.bytes > 0); assert.equal(pasted.pages, result.pages);
      report.checks.push('cross-page-selection-paste-and-undo');
    }
  }
  // 功能：在真实表格中通过浏览器组合输入提交中文，再使用同一宿主写盘并重开；时间：2026-11-09；作者：lq。
  await page.evaluate(() => {
    const editor = window.__surface.session.editor;
    let from;
    editor.state.doc.descendants((node, pos) => { if (from === undefined && node.isText) from = pos; });
    editor.commands.setTextSelection(from);
    editor.view.focus();
  });
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Input.imeSetComposition', { text: 'zhong wen', selectionStart: 9, selectionEnd: 9 });
  const guarded = await page.evaluate(async () => { try { await window.__surface.save(async () => {}); return false; } catch (error) { return String(error).includes('输入尚未确认'); } });
  assert.equal(guarded, true);
  await cdp.send('Input.insertText', { text: '中文输入验收' });
  await page.waitForFunction(() => !window.__surface.session.editor.view.composing && window.__surface.session.editor.getText().includes('中文输入验收'));
  const saved = await page.evaluate(async () => {
    const surface = window.__surface;
    const editor = surface.session.editor;
    editor.commands.insertContent('撤销重做');
    editor.commands.undo();
    const undone = !editor.getText().includes('撤销重做');
    editor.commands.redo();
    const redone = editor.getText().includes('撤销重做');
    let written;
    await surface.save(async bytes => { written = Array.from(bytes); });
    return { bytes: written, undone, redone, dirty: surface.session.state().dirty };
  });
  assert.equal(saved.undone, true); assert.equal(saved.redone, true); assert.equal(saved.dirty, false);
  await writeFile(join(output, 'saved.docx'), new Uint8Array(saved.bytes));
  const reopened = await page.evaluate(async bytes => {
    window.__surface.dispose();
    const surface = await window.__surfaceModule.WordCoreSurface.mount(document.getElementById('core'), new Uint8Array(bytes));
    window.__surface = surface;
    surface.layout.flush();
    const prepared = await surface.session.prepareSave();
    return { text: surface.session.editor.getText(), pages: surface.layout.current()?.pageCount, bytes: Array.from(prepared.bytes) };
  }, saved.bytes);
  assert.ok(reopened.text.includes('中文输入验收')); assert.ok(reopened.text.includes('撤销重做')); assert.ok(reopened.pages > 1);
  assert.deepEqual(reopened.bytes, saved.bytes);
  report.checks.push('real-ime-save-guard', 'table-chinese-input', 'transaction-undo-redo', 'save-reopen-and-second-noop-save');
  // 功能：多栏页面使用同源布局，正文与来源保持原样；时间：2026-11-09；作者：lq。
  const unsupported = await page.evaluate(async bytes => {
    window.__surface.dispose();
    const surface = await window.__surfaceModule.WordCoreSurface.mount(document.getElementById('core'), new Uint8Array(bytes));
    window.__surface = surface;
    let error;
    try { surface.layout.flush(); } catch (caught) { error = String(caught); }
    return { error, hasLayout: !!surface.layout.current(), dirty: surface.session.state().dirty };
  }, await fixture(paragraph('多栏文档'), sections.replace('</w:sectPr>', '<w:cols w:num="2"/></w:sectPr>')));
  assert.equal(unsupported.error, undefined); assert.equal(unsupported.hasLayout, true); assert.equal(unsupported.dirty, false);
  report.checks.push('column-layout-without-changing-document');
  assert.deepEqual(errors, []);
  report.status = 'passed'; report.browser = browser.version();
} catch (error) {
  report.status = 'failed'; report.error = String(error); process.exitCode = 1;
} finally {
  await browser.close();
  await writeFile(join(output, 'report.json'), JSON.stringify({ ...report, errors }, null, 2));
  console.log(JSON.stringify({ ...report, errors }));
}
