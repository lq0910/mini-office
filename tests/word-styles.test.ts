/**
 * 功能：验证 Word 样式库——预设完整性、样式套用语义、样式识别、样式集持久化与 DOCX 往返；
 * 时间：2026-10-29；作者：lq。
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createBlankDoc, type DocModel, type DocParagraph } from '@mini-office/document-model';
import {
  addStyleSet,
  applyStyleToParagraph,
  BUILTIN_STYLES,
  clearParagraphFormatting,
  collectUsedStyleDefinitions,
  createStyleFromParagraph,
  detectStyleId,
  documentStyles,
  effectiveStyles,
  findStyle,
  importStyleSet,
  isBuiltinStyle,
  readStyleSets,
  removeStyleSet,
  resolveStyleRun,
  STYLE_SET_STORAGE_KEY,
} from '@mini-office/editor-word';
import { mergeStyleDefinitions, serializeDocx } from '@mini-office/editor-word';
import { parseDocx } from '@mini-office/editor-word';

function paragraph(overrides: Partial<DocParagraph> = {}): DocParagraph {
  return {
    type: 'paragraph',
    id: 'p-1',
    runs: [{ text: '示例文字', style: { fontSizePt: 10.5 } }],
    style: { alignment: 'left', lineHeight: 1.6 },
    ...overrides,
  };
}

function memoryStorage(): Storage {
  const map = new Map<string, string>();
  return {
    get length() {
      return map.size;
    },
    clear: () => map.clear(),
    getItem: (key: string) => map.get(key) ?? null,
    key: (index: number) => Array.from(map.keys())[index] ?? null,
    removeItem: (key: string) => void map.delete(key),
    setItem: (key: string, value: string) => void map.set(key, String(value)),
  } as Storage;
}

describe('样式预设', () => {
  it('预设 ID 唯一，且标题层级连续', () => {
    const ids = BUILTIN_STYLES.map((style) => style.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(BUILTIN_STYLES.length).toBeGreaterThanOrEqual(16);
    const levels = ['Heading1', 'Heading2', 'Heading3', 'Heading4'].map(
      (id) => findStyle(BUILTIN_STYLES, id)?.paragraph.outlineLevel,
    );
    expect(levels).toEqual([0, 1, 2, 3]);
  });

  it('正文样式不锁死字体与字号，跟随文档默认值', () => {
    const normal = findStyle(BUILTIN_STYLES, 'Normal');
    expect(normal).toBeDefined();
    const doc = { defaultFont: '仿宋', defaultFontSizePt: 12 };
    expect(resolveStyleRun(normal!, doc)).toMatchObject({ fontFamily: '仿宋', fontSizePt: 12 });
  });

  it('isBuiltinStyle 只认内置预设', () => {
    expect(isBuiltinStyle('Heading1')).toBe(true);
    expect(isBuiltinStyle('MoStyle_abc')).toBe(false);
  });
});

describe('样式套用', () => {
  const doc = { defaultFont: '宋体', defaultFontSizePt: 10.5 };

  it('套用标题 1 会写入字号、加粗、大纲级别与样式关联', () => {
    const heading = findStyle(BUILTIN_STYLES, 'Heading1')!;
    const next = applyStyleToParagraph(paragraph(), heading, { doc });
    expect(next.styleId).toBe('Heading1');
    expect(next.style.outlineLevel).toBe(0);
    expect(next.style.keepNext).toBe(true);
    expect(next.runs[0].style).toMatchObject({ fontSizePt: 18, bold: true, fontFamily: '宋体' });
  });

  it('回到正文会清掉加粗并解除样式关联', () => {
    const heading = findStyle(BUILTIN_STYLES, 'Heading1')!;
    const normal = findStyle(BUILTIN_STYLES, 'Normal')!;
    const styled = applyStyleToParagraph(paragraph(), heading, { doc });
    const back = applyStyleToParagraph(styled, normal, { doc });
    expect(back.styleId).toBeUndefined();
    expect(back.runs[0].style.bold).toBeUndefined();
    expect(back.runs[0].style.fontSizePt).toBe(10.5);
    expect(back.style.outlineLevel).toBeUndefined();
  });

  it('套用样式保留用户的下划线、突出显示等强调格式', () => {
    const heading = findStyle(BUILTIN_STYLES, 'Heading2')!;
    const source = paragraph({ runs: [{ text: '重点', style: { highlight: '#FFF3B0', underline: true } }] });
    const next = applyStyleToParagraph(source, heading, { doc });
    expect(next.runs[0].style.highlight).toBe('#FFF3B0');
    expect(next.runs[0].style.underline).toBe(true);
    expect(next.runs[0].style.bold).toBe(true);
  });

  it('套用样式会重写段落级排版并清掉旧边框与底纹', () => {
    const quote = findStyle(BUILTIN_STYLES, 'Quote')!;
    const source = paragraph({
      style: {
        alignment: 'right',
        backgroundColor: '#FFE4E4',
        borders: { top: { style: 'single', widthPx: 1, color: '#000000' } },
        indentLeftPx: 40,
      },
    });
    const next = applyStyleToParagraph(source, quote, { doc });
    expect(next.style.alignment).toBe('left');
    expect(next.style.backgroundColor).toBeUndefined();
    expect(next.style.borders?.top).toBeUndefined();
    expect(next.style.borders?.left?.widthPx).toBe(3);
    expect(next.style.indentLeftPx).toBeGreaterThan(0);
  });

  it('列表样式写入标记，正文样式清除标记', () => {
    const bullet = findStyle(BUILTIN_STYLES, 'MoListBullet')!;
    const number = findStyle(BUILTIN_STYLES, 'MoListNumber')!;
    const normal = findStyle(BUILTIN_STYLES, 'Normal')!;
    expect(applyStyleToParagraph(paragraph(), bullet, { doc }).list).toMatchObject({
      ordered: false,
      marker: '•',
    });
    expect(
      applyStyleToParagraph(paragraph(), number, { doc, marker: '3.' }).list,
    ).toMatchObject({ ordered: true, marker: '3.' });
    expect(applyStyleToParagraph(paragraph({ list: { level: 0, marker: '•', ordered: false } }), normal, { doc }).list).toBeUndefined();
  });

  it('清除格式回到文档默认字体与无缩进状态', () => {
    const heading = findStyle(BUILTIN_STYLES, 'Heading1')!;
    const styled = applyStyleToParagraph(paragraph(), heading, { doc });
    const cleared = clearParagraphFormatting(styled, doc);
    expect(cleared.styleId).toBeUndefined();
    expect(cleared.style).toMatchObject({ alignment: 'left', spaceBeforePx: 0, firstLineIndentPx: 0 });
    expect(cleared.runs[0].style).toEqual({ fontFamily: '宋体', fontSizePt: 10.5 });
  });
});

describe('样式识别', () => {
  const doc = { defaultFont: '宋体', defaultFontSizePt: 10.5, styles: [] as DocModel['styles'] };

  it('优先按 pStyle 识别', () => {
    const heading = findStyle(BUILTIN_STYLES, 'Heading2')!;
    const styled = applyStyleToParagraph(paragraph(), heading, { doc });
    expect(detectStyleId(styled, BUILTIN_STYLES, doc)).toBe('Heading2');
  });

  it('没有 pStyle 时按排版特征识别', () => {
    const manual = paragraph({
      runs: [{ text: '手调标题', style: { fontSizePt: 18, bold: true } }],
      style: {
        alignment: 'left',
        outlineLevel: 0,
        keepNext: true,
        lineHeight: 1.3,
        spaceBeforePx: 16,
        spaceAfterPx: 8,
        firstLineIndentPx: 0,
        indentLeftPx: 0,
      },
    });
    expect(detectStyleId(manual, BUILTIN_STYLES, doc)).toBe('Heading1');
  });

  it('无法匹配任何样式时返回 undefined', () => {
    const odd = paragraph({
      runs: [{ text: '奇怪格式', style: { fontSizePt: 13.5, color: '#00FF00' } }],
      style: { alignment: 'right', lineHeight: 1.7 },
    });
    expect(detectStyleId(odd, BUILTIN_STYLES, doc)).toBeUndefined();
  });
});

describe('文档样式与样式集', () => {
  const doc = { defaultFont: '宋体', defaultFontSizePt: 10.5, styles: [] as DocModel['styles'] };

  it('文档内定义覆盖同名预设，但保留中文名', () => {
    const custom = {
      id: 'Heading1',
      name: 'heading 1',
      paragraph: { fontSizePt: 20 } as DocParagraph['style'],
      run: { fontSizePt: 20 },
      origin: 'imported' as const,
    };
    const merged = effectiveStyles({ styles: [custom] });
    const heading = findStyle(merged, 'Heading1')!;
    expect(heading.name).toBe('标题 1');
    expect(heading.run.fontSizePt).toBe(20);
  });

  it('以段落格式创建自定义样式', () => {
    const source = paragraph({
      runs: [{ text: '条款', style: { fontSizePt: 12, bold: true, color: '#B42318' } }],
      style: { alignment: 'center', spaceBeforePx: 8 },
    });
    const created = createStyleFromParagraph('合同条款', source, doc);
    expect(created.id.startsWith('MoStyle_')).toBe(true);
    expect(created.name).toBe('合同条款');
    expect(created.origin).toBe('custom');
    expect(created.run).toMatchObject({ fontSizePt: 12, bold: true, color: '#B42318' });
    expect(created.paragraph.alignment).toBe('center');
    expect(documentStyles({ styles: [created] })).toHaveLength(1);
  });

  it('样式集保存、覆盖、读取与删除', () => {
    const storage = memoryStorage();
    const created = createStyleFromParagraph('公文标题', paragraph(), doc);
    const first = addStyleSet('公文', [created], storage);
    expect(first).toHaveLength(1);
    expect(first[0].styles[0].name).toBe('公文标题');

    const second = addStyleSet('公文', [created, created], storage);
    expect(second).toHaveLength(1);
    expect(second[0].styles).toHaveLength(2);

    expect(readStyleSets(storage)[0].name).toBe('公文');
    expect(removeStyleSet(second[0].id, storage)).toHaveLength(0);
    expect(readStyleSets(storage)).toHaveLength(0);
  });

  it('样式集数据损坏时返回空列表而不是抛错', () => {
    const storage = memoryStorage();
    storage.setItem(STYLE_SET_STORAGE_KEY, '{ not json');
    expect(readStyleSets(storage)).toEqual([]);
  });

  it('导入样式集时冲突的 ID 会重新分配', () => {
    const existing = createStyleFromParagraph('已有样式', paragraph(), doc);
    const incoming = { ...existing, name: '外来样式' };
    const set = { id: 'set_1', name: '外部', styles: [incoming], savedAt: new Date().toISOString() };
    const merged = importStyleSet(set, [existing]);
    expect(merged).toHaveLength(2);
    expect(merged[1].id).not.toBe(existing.id);
    expect(merged[1].name).toBe('外来样式');
  });

  it('只收集文档里真正用到的样式', () => {
    const docModel = createBlankDoc();
    docModel.blocks = [
      { type: 'paragraph', id: 'p1', runs: [{ text: '标题', style: {} }], style: {}, styleId: 'Heading1' },
      { type: 'paragraph', id: 'p2', runs: [{ text: '正文', style: {} }], style: {} },
    ];
    const used = collectUsedStyleDefinitions(docModel);
    expect(used.map((style) => style.id)).toEqual(['Heading1']);
  });
});

describe('styles.xml 写入', () => {
  const base = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

  it('追加缺失的样式，元素顺序符合 CT_Style', () => {
    const doc = createBlankDoc();
    const heading = findStyle(BUILTIN_STYLES, 'Heading1')!;
    const xml = mergeStyleDefinitions(base, [heading], doc);
    expect(xml).toContain('w:styleId="Heading1"');
    expect(xml).toContain('<w:name w:val="heading 1"/>');
    const nameAt = xml.indexOf('<w:name w:val="heading 1"/>');
    const basedAt = xml.indexOf('<w:basedOn w:val="Normal"/>');
    const nextAt = xml.indexOf('<w:next w:val="Normal"/>');
    const qFormatAt = xml.indexOf('<w:qFormat/>', nameAt);
    const pPrAt = xml.indexOf('<w:pPr>', nameAt);
    const rPrAt = xml.indexOf('<w:rPr>', pPrAt);
    expect(nameAt).toBeGreaterThan(-1);
    expect(basedAt).toBeGreaterThan(nameAt);
    expect(nextAt).toBeGreaterThan(basedAt);
    expect(qFormatAt).toBeGreaterThan(nextAt);
    expect(pPrAt).toBeGreaterThan(qFormatAt);
    expect(rPrAt).toBeGreaterThan(pPrAt);
    // 已存在的 Normal 不重复写入
    expect(xml.match(/w:styleId="Normal"/g)).toHaveLength(1);
  });

  it('重复合并保持幂等', () => {
    const doc = createBlankDoc();
    const styles = [findStyle(BUILTIN_STYLES, 'Heading1')!, findStyle(BUILTIN_STYLES, 'Quote')!];
    const once = mergeStyleDefinitions(base, styles, doc);
    const twice = mergeStyleDefinitions(once, styles, doc);
    expect(twice).toBe(once);
  });
});

describe('DOCX 往返', () => {
  it('套用标题后保存，styles.xml 与段落 pStyle 都能还原', async () => {
    const doc = createBlankDoc();
    const heading = findStyle(BUILTIN_STYLES, 'Heading1')!;
    doc.blocks = [applyStyleToParagraph({ ...paragraph({ id: 'p1', runs: [{ text: '第一章 总则', style: {} }] }) }, heading, { doc })];
    const bytes = await serializeDocx(doc, { name: 'styles.docx' });

    const zip = await JSZip.loadAsync(bytes);
    const stylesXml = await zip.file('word/styles.xml')!.async('text');
    expect(stylesXml).toContain('w:styleId="Heading1"');

    const reopened = await parseDocx(bytes, { fileName: 'styles.docx' });
    const first = reopened.blocks[0];
    expect(first.type).toBe('paragraph');
    if (first.type !== 'paragraph') return;
    expect(first.styleId).toBe('Heading1');
    expect(first.style.outlineLevel).toBe(0);
    expect(first.runs[0].style.fontSizePt).toBe(18);
    expect(reopened.styles?.some((style) => style.id === 'Heading1')).toBe(true);
  });
});

describe('docx 下划线解析回归', () => {
  // 回归：<w:u> 存在但缺省 w:val（如 <w:u w:color="000000"/>）按 OOXML 不产生下划线；
  // Word/WPS 不显示，mini-office 此前误判为下划线。时间：2026-10-30；作者：hou0。
  it('w:u 缺省 w:val 时视为无下划线，与 Word/WPS 一致', async () => {
    const doc = createBlankDoc();
    doc.blocks = [
      { type: 'paragraph', id: 'p1', runs: [{ text: '字段值', style: { underline: true } }], style: {} },
    ];
    const bytes = await serializeDocx(doc, { name: 'u.docx' });
    const zip = await JSZip.loadAsync(bytes);
    const docXml = (await zip.file('word/document.xml')!.async('text')).replace(
      '<w:u w:val="single"/>',
      '<w:u w:color="000000"/>',
    );
    zip.file('word/document.xml', docXml);
    const reopened = await parseDocx(await zip.generateAsync({ type: 'uint8array' }), { fileName: 'u.docx' });
    const run = reopened.blocks[0];
    expect(run.type).toBe('paragraph');
    if (run.type !== 'paragraph') return;
    expect(run.runs[0].style.underline).toBe(false);
  });
});
