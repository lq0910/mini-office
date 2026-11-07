/** 功能：只读 DOCX 经真实浏览器输入和下载保存，验证未编辑表格属性与部件保持；时间：2026-11-07；作者：lq。 */
import { chromium } from 'playwright-core';
import JSZip from 'jszip';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { homedir } from 'node:os';
import { createHash } from 'node:crypto';

const [file, target = '建设内容'] = process.argv.slice(2);
if (!file || !/\.docx$/i.test(file)) throw new Error('用法：node scripts/verify-word-save.mjs 文件.docx [目标段落文字]');
const output = resolve(process.env.MINI_OFFICE_SAVE_OUTPUT ?? 'output/playwright/word-save.local');
if (!output.split('/').some(part => part.endsWith('.local'))) throw new Error('私人验收产物必须写入 .local 目录');
await mkdir(output, { recursive: true });
const input = await readFile(resolve(file));
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
/** 功能：读取 OOXML 字符串、表格物理属性及全部部件，避免回读模型掩盖损失；时间：2026-11-07；作者：lq。 */
async function inspect(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file('word/document.xml').async('text');
  const properties = {};
  for (const tag of ['tblPr', 'tblGrid', 'trPr', 'tcPr']) {
    properties[tag] = [...xml.matchAll(new RegExp(`<w:${tag}\\b[^>]*(?:/>|>[\\s\\S]*?</w:${tag}>)`, 'g'))].map(match => match[0]);
  }
  return { zip, xml, properties, borders: (xml.match(/<w:tcBorders\b/g) ?? []).length };
}
/** 功能：对普通文字输入要求除 w:t 外的整个 document.xml 完全保持；时间：2026-11-07；作者：lq。 */
function withoutText(xml) { return xml.replace(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/g, '<w:t/>'); }
const browser = await chromium.launch({ headless: true, executablePath: process.env.MINI_OFFICE_CHROME ?? join(homedir(), 'Library/Caches/ms-playwright/chromium-1217/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing') });
const report = { sourceSha256: hash(input), browser: browser.version(), stages: [] };
const original = await inspect(input);
try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  const errors = [];
  page.on('pageerror', error => errors.push(String(error)));
  await page.addInitScript(() => { window.showOpenFilePicker = undefined; window.showSaveFilePicker = undefined; });
  await page.goto(process.env.MINI_OFFICE_TEST_URL ?? 'http://127.0.0.1:1420/', { waitUntil: 'domcontentloaded' });
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: '打开文件', exact: true }).first().click();
  await (await chooser).setFiles(resolve(file));
  await page.waitForSelector('[data-word-editor]', { timeout: 60000 });
  // 功能：所有编辑通过浏览器键盘输入，保存走产品快捷键和下载链路，不直接修改文档模型；时间：2026-11-07；作者：lq。
  const paragraph = page.locator('[data-paragraph-id][role="textbox"]').filter({ hasText: target }).first();
  await paragraph.waitFor({ state: 'visible' });
  const stages = ['no-op', 'text', 'line-break', 'paragraph'];
  for (const stage of stages) {
    if (stage !== 'no-op') {
      await paragraph.click();
      await paragraph.press('End');
      if (stage === 'line-break') await page.keyboard.press('Shift+Enter');
      if (stage === 'paragraph') await page.keyboard.press('Enter');
      await page.keyboard.insertText(stage === 'text' ? '保存验证' : '追加文字');
    }
    const downloadPromise = page.waitForEvent('download', { timeout: 15000 });
    await page.keyboard.press('ControlOrMeta+s');
    const download = await downloadPromise;
    const savedPath = join(output, `${stage}.docx`);
    await download.saveAs(savedPath);
    const bytes = await readFile(savedPath);
    const saved = await inspect(bytes);
    const propertiesUnchanged = JSON.stringify(original.properties) === JSON.stringify(saved.properties);
    if (!propertiesUnchanged) throw new Error(`${stage}: 未编辑的表格属性变化`);
    if (stage === 'no-op' && !input.equals(bytes)) throw new Error('无操作保存改变原文件字节');
    if (stage === 'text' && withoutText(original.xml) !== withoutText(saved.xml)) throw new Error('普通文字编辑改变了目标文字以外的 XML');
    if (stage === 'text' && !saved.xml.includes('保存验证')) throw new Error('输入未进入保存文件');
    const originalParts = Object.keys(original.zip.files).filter(name => !original.zip.files[name].dir).sort();
    const savedParts = Object.keys(saved.zip.files).filter(name => !saved.zip.files[name].dir).sort();
    if (JSON.stringify(originalParts) !== JSON.stringify(savedParts)) throw new Error(`${stage}: ZIP 部件集合变化`);
    for (const name of originalParts) {
      if (name === 'word/document.xml') continue;
      const a = await original.zip.file(name).async('nodebuffer');
      const b = await saved.zip.file(name).async('nodebuffer');
      if (!a.equals(b)) throw new Error(`${stage}: 非目标部件 ${name} 变化`);
    }
    report.stages.push({ stage, sha256: hash(bytes), borders: saved.borders, tablePropertiesUnchanged: propertiesUnchanged, otherPartsUnchanged: true });
  }
  await page.screenshot({ path: join(output, 'saved.png') });
  if (errors.length) throw new Error(errors.join('\n'));
  report.sourceUnchanged = hash(await readFile(resolve(file))) === report.sourceSha256;
  if (!report.sourceUnchanged) throw new Error('源文件哈希变化');
  report.status = 'passed';
} catch (error) {
  report.status = 'failed';
  report.error = String(error);
  process.exitCode = 1;
} finally {
  await browser.close();
  await writeFile(join(output, 'report.json'), JSON.stringify(report, null, 2));
  console.log(JSON.stringify(report));
}
