/**
 * 功能：验证 Word 制表位定位/前导符、目录条目采集与页码计算、以及二者写回 docx 的结果。
 * 时间：2026-10-30；作者：lq。
 *
 * 覆盖用户反馈的两个缺口：
 * 1. 正文 Tab 与标尺制表位此前没有任何结构化数据，渲染只能退化成空格，右对齐与点线前导无法实现；
 * 2. 目录此前只能靠手动敲点线和页码，增删内容后页码不会跟随。
 */
import { afterEach, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { createBlankDoc, type DocBlock, type DocModel, type DocParagraph } from '@mini-office/document-model';
import { DEFAULT_TAB_STOP_PX, applyTabStops, nextTabStop } from '../packages/editor-word/src/tab-stops';
import { buildTocEntries, collectTocHeadings, headingLevelOf, paragraphOutlineText } from '../packages/editor-word/src/toc';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import { parseDocxPackage } from '../packages/editor-word/src/parse';
import { serializeDomRuns } from '../packages/editor-word/src/editing';
import { loadOoxmlPackage } from '@mini-office/ooxml';
import type { PageSlice } from '../packages/editor-word/src/pagination';
import { renderRunsHtml } from '../packages/editor-word/src/html';

// 功能：恢复新增几何测量回归中的 DOM mock，防止影响其他制表位用例；时间：2026-11-03；作者：lq。
afterEach(() => vi.restoreAllMocks());

function paragraph(overrides: Partial<DocParagraph> & { id: string; text?: string }): DocParagraph {
  const { text, ...rest } = overrides;
  return {
    type: 'paragraph',
    runs: [{ text: text ?? '', style: { fontSizePt: 10.5 } }],
    style: { alignment: 'left', lineHeight: 1.6 },
    ...rest,
  } as DocParagraph;
}

function docWith(blocks: DocBlock[]): DocModel {
  const doc = createBlankDoc();
  return { ...doc, blocks };
}

describe('制表位定位', () => {
  it('没有自定义制表位时按默认间距等距推进', () => {
    expect(DEFAULT_TAB_STOP_PX).toBe(48);
    expect(nextTabStop(undefined, 0)).toMatchObject({ posPx: 48, alignment: 'left' });
    expect(nextTabStop([], 40)).toMatchObject({ posPx: 48 });
    expect(nextTabStop([], 48)).toMatchObject({ posPx: 96 });
    expect(nextTabStop([], 10, 60)).toMatchObject({ posPx: 60 });
  });

  it('优先使用右侧最近的自定义制表位，并带回对齐方式与前导符', () => {
    const stops = [
      { posPx: 400, alignment: 'right' as const, leader: 'dot' as const },
      { posPx: 120, alignment: 'left' as const },
      { posPx: 240, alignment: 'center' as const, leader: 'underscore' as const },
    ];
    expect(nextTabStop(stops, 0)).toMatchObject({ posPx: 120, alignment: 'left' });
    expect(nextTabStop(stops, 130)).toMatchObject({ posPx: 240, alignment: 'center', leader: 'underscore' });
    expect(nextTabStop(stops, 260)).toMatchObject({ posPx: 400, alignment: 'right', leader: 'dot' });
  });

  it('越过最后一个自定义制表位后回到默认间距，与 Word 一致', () => {
    const stops = [{ posPx: 100, alignment: 'left' as const }];
    expect(nextTabStop(stops, 101)).toMatchObject({ posPx: 144, alignment: 'left' });
    expect(nextTabStop(stops, 101)).not.toHaveProperty('leader');
  });

  it('非法间距回落到默认值，避免出现除零与负宽度', () => {
    expect(nextTabStop([], 0, 0)).toMatchObject({ posPx: 48 });
    expect(nextTabStop([], 0, -10)).toMatchObject({ posPx: 48 });
  });
});

describe('制表符渲染后置处理', () => {
  // 功能：覆盖真实语料的文字内 Tab，保留字符样式和编辑后的 Tab 语义；时间：2026-11-03；作者：lq。
  it('文字内 Tab 与结构化 Tab 都输出可测量的占位，且往返编辑不丢失', () => {
    const runs = [{ text: '标题\t12', style: { fontSizePt: 11, bold: true } }];
    const root = document.createElement('p');
    root.innerHTML = renderRunsHtml(runs);
    expect(root.querySelectorAll('[data-tab]')).toHaveLength(1);
    expect(root.querySelector('[data-tab]')?.parentElement?.style.fontWeight).toBe('700');
    const restored = serializeDomRuns(root, {}, new Map());
    expect(restored.map((run) => run.tab ? '\t' : run.text ?? '').join('')).toBe('标题\t12');
    expect(runs[0].text).toBe('标题\t12');
  });

  /** 功能：构造带嵌套文字矩形的真实制表位几何，区分同一行多矩形与多行；时间：2026-11-03；作者：lq。 */
  function tabGeometry(multiline = false) {
    const root = document.createElement('div');
    root.innerHTML = '<p data-paragraph-id="p-indent"><span>标题</span><span data-tab="1"></span><span><b>12</b></span></p>';
    const paragraph = root.firstElementChild as HTMLElement;
    const tab = root.querySelector<HTMLElement>('[data-tab]')!;
    Object.defineProperty(paragraph, 'offsetWidth', { value: 400 });
    vi.spyOn(paragraph, 'getBoundingClientRect').mockReturnValue({ left: 32, width: 400 } as DOMRect);
    vi.spyOn(tab, 'getBoundingClientRect').mockReturnValue({ left: 132 } as DOMRect);
    const create = document.createRange.bind(document);
    vi.spyOn(document, 'createRange').mockImplementation(() => {
      const range = create();
      Object.assign(range, { getClientRects: () => [
        { left: 132, right: 152, top: 0, bottom: 20, width: 20, height: 20 },
        { left: 132, right: 152, top: multiline ? 24 : 0, bottom: multiline ? 44 : 20, width: 20, height: 20 },
      ] });
      return range;
    });
    return { root, tab };
  }

  it('右对齐按同一行矩形并集扣除页码宽度，并消除段落缩进漂移', () => {
    const { root, tab } = tabGeometry();
    applyTabStops(root, () => ({ stops: [{ posPx: 400, alignment: 'right', leader: 'dot' }], indentLeftPx: 32 }));
    expect(parseFloat(tab.style.width)).toBe(248);
    expect(tab.style.borderBottom).toContain('dotted');
  });

  it('目录标题与页码分属嵌套样式节点时仍测量页码宽度，避免右对齐 Tab 将页码挤到下一行', () => {
    const root = document.createElement('div');
    root.innerHTML = '<p data-paragraph-id="toc"><span>较长的目录标题</span><span><span data-tab="1"></span></span><span><b>12</b></span></p>';
    const paragraph = root.firstElementChild as HTMLElement;
    const tab = root.querySelector<HTMLElement>('[data-tab]')!;
    Object.defineProperty(paragraph, 'offsetWidth', { value: 500 });
    vi.spyOn(paragraph, 'getBoundingClientRect').mockReturnValue({ left: 0, width: 500 } as DOMRect);
    vi.spyOn(tab, 'getBoundingClientRect').mockReturnValue({ left: 300 } as DOMRect);
    const create = document.createRange.bind(document);
    vi.spyOn(document, 'createRange').mockImplementation(() => {
      const range = create();
      Object.assign(range, { getClientRects: () => [
        { left: 300, right: 320, top: 0, bottom: 20, width: 20, height: 20 },
      ] });
      return range;
    });

    applyTabStops(root, () => ({ stops: [{ posPx: 480, alignment: 'right', leader: 'dot' }] }));

    expect(Number.parseFloat(tab.style.width)).toBeCloseTo(160, 2);
    expect(tab.style.borderBottom).toContain('dotted');
  });

  it('真正跨行的后续内容仍回退左对齐，避免错误扣除整段宽度', () => {
    const { root, tab } = tabGeometry(true);
    applyTabStops(root, () => ({ stops: [{ posPx: 400, alignment: 'right' }], indentLeftPx: 32 }));
    expect(parseFloat(tab.style.width)).toBe(268);
  });

  it('按实测位置补宽度，并按前导符画下边框', () => {
    const root = document.createElement('div');
    root.innerHTML =
      '<p data-paragraph-id="p-1"><span>甲方：</span>' +
      '<span data-tab="1" style="display:inline-block;width:0;height:1em"></span><span>深圳市XXX有限公司</span></p>';
    const changed = applyTabStops(root, (id) =>
      id === 'p-1' ? { stops: [{ posPx: 160, alignment: 'left' }] } : undefined,
    );
    expect(changed).toBe(true);
    const tab = root.querySelector<HTMLElement>('[data-tab]')!;
    // jsdom 不做排版，所有矩形都是 0，因此宽度直接等于制表位位置。
    expect(Number.parseFloat(tab.style.width)).toBeCloseTo(160, 2);
    expect(tab.style.borderBottom).toBe('');
  });

  it('点线制表位输出虚线边框，未知段落不做任何改动', () => {
    const root = document.createElement('div');
    root.innerHTML =
      '<p data-paragraph-id="p-2"><span data-tab="1"></span></p>' +
      '<p data-paragraph-id="p-3"><span data-tab="1"></span></p>';
    applyTabStops(root, (id) =>
      id === 'p-2' ? { stops: [{ posPx: 200, alignment: 'left', leader: 'dot' }] } : undefined,
    );
    const [handled, untouched] = Array.from(root.querySelectorAll<HTMLElement>('[data-tab]'));
    expect(handled.style.borderBottom).toContain('dotted');
    expect(untouched.style.width).toBe('');
  });

  it('没有制表符占位块时不做无谓的重排', () => {
    const root = document.createElement('div');
    root.innerHTML = '<p data-paragraph-id="p-1"><span>纯文字</span></p>';
    expect(applyTabStops(root, () => ({ stops: [{ posPx: 100, alignment: 'left' }] }))).toBe(false);
  });
});

describe('目录条目采集', () => {
  it('大纲级别与内置标题样式都能识别为标题', () => {
    expect(headingLevelOf(paragraph({ id: 'a', style: { outlineLevel: 0 } }))).toBe(1);
    expect(headingLevelOf(paragraph({ id: 'b', style: { outlineLevel: 2 } }))).toBe(3);
    expect(headingLevelOf(paragraph({ id: 'c', styleId: 'Heading2' }))).toBe(2);
    expect(headingLevelOf(paragraph({ id: 'd', styleId: '标题3' }))).toBe(3);
    // 非标题、以及越界的大纲级别都不进目录。
    expect(headingLevelOf(paragraph({ id: 'e', style: { outlineLevel: 9 } }))).toBe(0);
    expect(headingLevelOf(paragraph({ id: 'f', styleId: 'BodyText' }))).toBe(0);
  });

  it('条目文本忽略图片与域，制表符归一为空格', () => {
    const target = paragraph({
      id: 'g',
      runs: [
        { text: '第一章', style: {} },
        { text: '', style: {}, tab: true },
        { text: '  总则  ', style: {} },
        { text: '', style: {}, image: { id: 'img-1', src: 'data:image/png;base64,AA==', widthPx: 10, heightPx: 10 } },
        { text: '页码', style: {}, field: 'PAGE' },
      ],
    });
    expect(paragraphOutlineText(target)).toBe('第一章 总则 页码');
  });

  it('按正文顺序采集，跳过目录段落自身、表格与空标题', () => {
    const blocks: DocBlock[] = [
      paragraph({ id: 'h1', text: '第一章 总则', style: { outlineLevel: 0 } }),
      paragraph({ id: 'toc', toc: { levels: 3 }, style: { outlineLevel: 0 } }),
      {
        type: 'table',
        id: 't1',
        rows: [{ cells: [{ runs: [{ text: '表内标题', style: {} }], style: { outlineLevel: 0 } }] }],
        style: {},
      } as unknown as DocBlock,
      paragraph({ id: 'h2', text: '   ', style: { outlineLevel: 0 } }),
      paragraph({ id: 'h3', text: '1.1 范围', styleId: 'Heading2' }),
      paragraph({ id: 'h4', text: '第四级标题', style: { outlineLevel: 3 } }),
    ];
    const headings = collectTocHeadings(blocks, 3);
    expect(headings.map((item) => item.paragraphId)).toEqual(['h1', 'h3']);
    expect(headings[1]).toMatchObject({ level: 2, text: '1.1 范围' });
  });

  it('页码由分页结果反查，并叠加节的起始页码', () => {
    const doc = docWith([
      paragraph({ id: 'h1', text: '第一章 总则', style: { outlineLevel: 0 } }),
      paragraph({ id: 'h2', text: '1.1 范围', styleId: 'Heading2' }),
      paragraph({ id: 'h3', text: '第二章 附则', style: { outlineLevel: 0 } }),
    ]);
    const pages: PageSlice[] = [
      { index: 0, blocks: [doc.blocks[0]] },
      { index: 1, blocks: [doc.blocks[1]] },
      { index: 2, blocks: [doc.blocks[2]], section: { page: { ...doc.page, pageNumberStart: 10 } }, sectionPage: 1 },
    ];
    const entries = buildTocEntries(doc, pages, 3);
    expect(entries.map((entry) => [entry.text, entry.page])).toEqual([
      ['第一章 总则', 1],
      ['1.1 范围', 2],
      ['第二章 附则', 11],
    ]);
  });

  it('分页结果缺失时页码回落到第 1 页，不产生 NaN', () => {
    const doc = docWith([paragraph({ id: 'h1', text: '第一章 总则', style: { outlineLevel: 0 } })]);
    expect(buildTocEntries(doc, [], 3)[0]).toMatchObject({ page: 1 });
  });
});

describe('制表位与目录写回 docx', () => {
  it('段落制表位按 CT_PPr 顺序写在 shd 之后、spacing 之前', async () => {
    const doc = docWith([
      paragraph({
        id: 'p-1',
        text: '甲方',
        style: {
          alignment: 'left',
          backgroundColor: 'FFFF00',
          spaceAfterPx: 8,
          tabStops: [
            { posPx: 480, alignment: 'right', leader: 'dot' },
            { posPx: 96, alignment: 'left' },
          ],
        },
      }),
    ]);
    const bytes = await serializeDocx(doc, { name: 'tabs.docx' });
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('text');
    const tabs = xml.slice(xml.indexOf('<w:tabs>'), xml.indexOf('</w:tabs>') + '</w:tabs>'.length);
    // 7200 twips = 480px，1440 twips = 96px，按位置升序输出。
    expect(tabs).toBe(
      '<w:tabs><w:tab w:val="left" w:pos="1440"/><w:tab w:val="right" w:leader="dot" w:pos="7200"/></w:tabs>',
    );
    expect(xml.indexOf('<w:tabs>')).toBeGreaterThan(xml.indexOf('<w:shd'));
    expect(xml.indexOf('<w:tabs>')).toBeLessThan(xml.indexOf('<w:spacing'));
  });

  it('目录段落写成真正的 TOC 域，并写入实测页码与更新标记', async () => {
    const doc = docWith([
      paragraph({ id: 'toc', toc: { levels: 3 } }),
      paragraph({ id: 'h1', text: '第一章 总则', style: { outlineLevel: 0 } }),
    ]);
    const bytes = await serializeDocx(doc, {
      name: 'toc.docx',
      tocEntries: [
        { level: 1, text: '第一章 总则', page: 3, paragraphId: 'h1' },
        { level: 4, text: '被级别过滤掉的条目', page: 9, paragraphId: 'h9' },
      ],
    });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')!.async('text');
    expect(xml).toContain('w:fldCharType="begin"');
    // 引号在 XML 文本里被转义，Word 读到的是 TOC \o "1-3" \h \z \u。
    expect(xml).toContain(' TOC \\o &quot;1-3&quot; \\h \\z \\u ');
    expect(xml).toContain('<w:t xml:space="preserve">第一章 总则</w:t>');
    expect(xml).toContain('<w:tab/><w:t xml:space="preserve">3</w:t>');
    const begin = xml.indexOf('w:fldCharType="separate"');
    const fieldResult = xml.slice(begin, xml.indexOf('w:fldCharType="end"', begin));
    expect(fieldResult).not.toContain('被级别过滤掉的条目');
    // 打开文档时让 Word / WPS 自行刷新域，页码不会停在缓存值上。
    const settings = await zip.file('word/settings.xml')!.async('text');
    expect(settings).toContain('<w:updateFields w:val="true"/>');
  });

  it('没有传入实测页码时按标题重采，条目不会为空', async () => {
    const doc = docWith([
      paragraph({ id: 'toc', toc: { levels: 2 } }),
      paragraph({ id: 'h1', text: '第一章 总则', style: { outlineLevel: 0 } }),
      paragraph({ id: 'h3', text: '第三级不进目录', style: { outlineLevel: 2 } }),
    ]);
    const bytes = await serializeDocx(doc, { name: 'toc-fallback.docx' });
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('text');
    const begin = xml.indexOf('w:fldCharType="separate"');
    const fieldResult = xml.slice(begin, xml.indexOf('w:fldCharType="end"', begin));
    expect(fieldResult).toContain('第一章 总则');
    expect(fieldResult).not.toContain('第三级不进目录');
  });

  it('保存后重新打开，制表位、默认间距与目录级别都还在', async () => {
    const doc = docWith([
      paragraph({ id: 'toc', toc: { levels: 2 } }),
      paragraph({
        id: 'h1',
        text: '第一章 总则',
        style: { outlineLevel: 0, tabStops: [{ posPx: 480, alignment: 'right', leader: 'dot' }] },
      }),
    ]);
    const bytes = await serializeDocx(doc, {
      name: 'round-trip.docx',
      tocEntries: [{ level: 1, text: '第一章 总则', page: 1, paragraphId: 'h1' }],
    });
    const pkg = await loadOoxmlPackage(bytes);
    try {
      const parsed = await parseDocxPackage(pkg, { fileName: 'round-trip.docx' });
      // 默认制表位写 720 twips，解析回 48px，与渲染默认值一致，重开不会跳位。
      expect(parsed.defaultTabStopPx).toBe(DEFAULT_TAB_STOP_PX);
      const tocBlock = parsed.blocks.find((block) => block.type === 'paragraph' && block.toc);
      expect(tocBlock && tocBlock.type === 'paragraph' ? tocBlock.toc : undefined).toEqual({ levels: 2 });
      const heading = parsed.blocks.find(
        (block) => block.type === 'paragraph' && block.runs.some((run) => run.text === '第一章 总则'),
      );
      expect(heading && heading.type === 'paragraph' ? heading.style.tabStops : undefined).toEqual([
        { posPx: 480, alignment: 'right', leader: 'dot' },
      ]);
    } finally {
      pkg.dispose();
    }
  });
});

describe('编辑区 DOM 回写 runs', () => {
  it('浏览器插入的块级包裹转成段内换行，不会把两行挤成一行', () => {
    const root = document.createElement('p');
    root.innerHTML = '<span>甲方：</span><div>深圳市XXX有限公司</div>';
    const runs = serializeDomRuns(root, { fontSizePt: 10.5 }, new Map());
    expect(runs.map((run) => (run.lineBreak ? '\n' : run.text)).join('')).toBe('甲方：\n深圳市XXX有限公司');
  });

  it('首个块级包裹前不补多余换行，<br> 与制表符占位各自成 run', () => {
    const root = document.createElement('p');
    root.innerHTML =
      '<div>第一行</div><br><span data-tab="1"></span><span>右对齐内容</span>';
    const runs = serializeDomRuns(root, { fontSizePt: 10.5 }, new Map());
    expect(runs.map((run) => (run.tab ? '[tab]' : run.lineBreak ? '\n' : run.text)).join('')).toBe(
      '第一行\n[tab]右对齐内容',
    );
  });
});
