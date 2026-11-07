/** 功能：验证粘贴格式、安全解析及选区替换时保留段外内容；时间：2026-11-07；作者：lq。 */
import { describe, it, expect } from 'vitest';
import { createBlankDoc } from '@mini-office/document-model';
import { clipboardParagraphs } from '../packages/editor-word/src/clipboard';
import { collectParagraphs, makeParagraph } from '../packages/editor-word/src/doc-utils';
import { runsPlainText } from '../packages/editor-word/src/editing';
import { replaceDocumentSelection } from '../packages/editor-word/src/document-selection';

const html = '<p style="font-family:Arial;font-size:24pt;color:red;text-align:center"><b>标题</b></p>\n<p>正文<script>bad()</script></p>';
const destination = { fontFamily: '宋体', fontSizePt: 12, color: '#000000' };
describe('粘贴格式', () => {
  it('保留源字体颜色强调和段落对齐，丢弃脚本', () => {
    const parts = clipboardParagraphs(html, '标题\n正文', 'source', destination, {});
    expect(parts.map((p) => runsPlainText(p.runs))).toEqual(['标题', '正文']);
    expect(parts[0].runs[0].style).toMatchObject({ fontFamily: 'Arial', fontSizePt: 24, color: '#FF0000', bold: true });
    expect(parts[0].style.alignment).toBe('center');
  });
  it('匹配目标字体颜色，保留加粗；纯文本清除源强调', () => {
    const matched = clipboardParagraphs(html, '标题\n正文', 'destination', destination, { alignment: 'left' });
    expect(matched[0].runs[0].style).toMatchObject({ ...destination, bold: true });
    expect(matched[0].style.alignment).toBe('left');
    const plain = clipboardParagraphs(html, '标题\n正文', 'text', destination, {});
    expect(plain[0].runs[0].style).toEqual(destination);
  });
  it('多段粘贴替换选区，保留前后文本及其样式', () => {
    const doc = createBlankDoc();
    const paragraph = makeParagraph([{ text: '前选中后', style: { italic: true } }]);
    doc.blocks = [paragraph];
    const rich = clipboardParagraphs(html, '标题\n正文', 'source', destination, {});
    const result = replaceDocumentSelection(doc, [{ paragraphId: paragraph.id, start: 1, end: 3 }], '', rich)!;
    const parts = collectParagraphs(result.doc);
    expect(parts.map((p) => runsPlainText(p.runs))).toEqual(['前标题', '正文后']);
    expect(parts[0].runs[0].style.italic).toBe(true);
    expect(parts[0].runs[1].style.bold).toBe(true);
    expect(parts[1].runs.at(-1)?.style.italic).toBe(true);
    expect(result.offset).toBe(2);
    expect(runsPlainText(paragraph.runs)).toBe('前选中后');
  });
});
