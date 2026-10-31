/**
 * 功能：验证 Word 交互能力——VML 方框勾选（点击/往返）、内容控件（复选框/下拉/日期）解析与写回、
 * 表格整选拖柄；时间：2026-10-30；作者：hou0。
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createBlankDoc, type DocModel } from '@mini-office/document-model';
import {
  applyControlPatch,
  controlChoiceIndex,
  editTable,
  formatControlDate,
  parseDocx,
  renderParagraphHtml,
  serializeDocx,
  serializeDomRuns,
  tableCellPlacements,
} from '@mini-office/editor-word';

async function makeDocx(body: string, sectPr: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  );
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>${body}${sectPr}</w:body></w:document>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

const SECT = '<w:sectPr><w:pgSz w:w="11910" w:h="16840"/><w:pgMar w:top="1580" w:bottom="280" w:left="880" w:right="680"/></w:sectPr>';

describe('Word 可交互控件', () => {
  it('VML 方框可点击勾选，勾选态随文档保存并读回', async () => {
    const rect = '<w:r><w:pict><v:rect style="position:absolute;margin-left:100pt;margin-top:5pt;width:9.84pt;height:9.84pt;mso-position-horizontal-relative:page" stroked="true" strokecolor="#000000"/></w:pict></w:r>';
    const doc = await parseDocx(await makeDocx(`<w:p>${rect}<w:r><w:t>是否开展</w:t></w:r></w:p>`, SECT));
    const paragraph = doc.blocks[0];
    expect(paragraph.type).toBe('paragraph');
    const shape = paragraph.type === 'paragraph' ? paragraph.runs.find((run) => run.image?.vmlRect)?.image : undefined;
    expect(shape?.vmlRect?.checked).toBeUndefined();

    // 未勾选：渲染出可点击框，无可勾标记。
    const unchecked = renderParagraphHtml(paragraph as never);
    expect(unchecked).toContain('data-vml-checked="0"');
    expect(unchecked).toContain('cursor:pointer');
    expect(unchecked).not.toContain('data-vml-check="1"');

    // 勾选后：渲染 ✓，写回 v:textbox，重新打开仍是勾选状态。
    const checkedDoc: DocModel = {
      ...doc,
      blocks: doc.blocks.map((block) =>
        block.type === 'paragraph'
          ? { ...block, runs: block.runs.map((run) => (run.image?.vmlRect ? { ...run, image: { ...run.image, vmlRect: { ...run.image.vmlRect!, checked: true } } } : run)) }
          : block,
      ),
    };
    const checkedParagraph = checkedDoc.blocks[0];
    const checkedHtml = renderParagraphHtml(checkedParagraph as never);
    expect(checkedHtml).toContain('data-vml-checked="1"');
    expect(checkedHtml).toContain('data-vml-check="1"');

    const saved = await serializeDocx(checkedDoc, { name: 'checkbox.docx' });
    const reopened = await parseDocx(saved);
    const reopenedShape = reopened.blocks[0].type === 'paragraph'
      ? reopened.blocks[0].runs.find((run) => run.image?.vmlRect)?.image?.vmlRect
      : undefined;
    expect(reopenedShape?.checked).toBe(true);
  });

  it('内容控件：复选框/下拉/日期解析出控件语义并可写回', async () => {
    const body = [
      '<w:p><w:sdt><w:sdtPr><w:alias w:val="确认"/><w14:checkbox><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr>',
      '<w:sdtContent><w:r><w:t>☑</w:t></w:r></w:sdtContent></w:sdt></w:p>',
      '<w:p><w:sdt><w:sdtPr><w:tag w:val="level"/><w:dropDownList>',
      '<w:listItem w:displayText="合格" w:value="合格"/><w:listItem w:displayText="不合格" w:value="不合格"/>',
      '</w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>合格</w:t></w:r></w:sdtContent></w:sdt></w:p>',
      '<w:p><w:sdt><w:sdtPr><w:date><w:dateFormat w:val="yyyy年M月d日"/></w:date></w:sdtPr>',
      '<w:sdtContent><w:r><w:t>2026年10月30日</w:t></w:r></w:sdtContent></w:sdt></w:p>',
    ].join('');
    const doc = await parseDocx(await makeDocx(body, SECT));
    const controls = doc.blocks.flatMap((block) =>
      block.type === 'paragraph' ? block.runs.filter((run) => run.control).map((run) => run.control!) : [],
    );
    expect(controls.map((control) => control.kind)).toEqual(['checkbox', 'dropdown', 'date']);
    expect(controls[0]).toMatchObject({ kind: 'checkbox', checked: true, alias: '确认' });
    expect(controls[1]).toMatchObject({ kind: 'dropdown', tag: 'level', choices: ['合格', '不合格'] });
    expect(controls[2]).toMatchObject({ kind: 'date', dateFormat: 'yyyy年M月d日' });

    // 渲染：复选框显示勾选方块，下拉/日期显示带值的 chip。
    const checkboxHtml = renderParagraphHtml(doc.blocks[0] as never);
    expect(checkboxHtml).toContain('data-control="checkbox"');
    expect(checkboxHtml).toContain('data-control-checked="1"');
    expect(checkboxHtml).toContain('☑');
    const dropdownHtml = renderParagraphHtml(doc.blocks[1] as never);
    expect(dropdownHtml).toContain('data-control="dropdown"');
    expect(dropdownHtml).toContain('data-control-choices="[&quot;合格&quot;,&quot;不合格&quot;]"');
    expect(dropdownHtml).toContain('合格');

    // 写回：三类控件都还原成 w:sdt（含 w14:checkbox / dropDownList / date）。
    const saved = await serializeDocx(doc, { name: 'controls.docx' });
    const xml = await (async () => {
      const zip = await JSZip.loadAsync(saved);
      return (await zip.file('word/document.xml')!.async('text'));
    })();
    expect(xml).toContain('<w14:checkbox>');
    expect(xml).toContain('w14:val="1"');
    expect(xml).toContain('<w:dropDownList>');
    expect(xml).toContain('w:displayText="不合格"');
    expect(xml).toContain('<w:dateFormat w:val="yyyy年M月d日"/>');

    const reopened = await parseDocx(saved);
    const again = reopened.blocks.flatMap((block) =>
      block.type === 'paragraph' ? block.runs.filter((run) => run.control).map((run) => run.control!) : [],
    );
    expect(again.map((control) => control.kind)).toEqual(['checkbox', 'dropdown', 'date']);
    expect(again[1].choices).toEqual(['合格', '不合格']);
  });

  it('控件取值更新与日期格式化', () => {
    const doc = createBlankDoc();
    doc.blocks = [{
      type: 'paragraph',
      id: 'p1',
      style: {},
      runs: [{ text: '', style: {}, control: { kind: 'dropdown', choices: ['A', 'B'], placeholder: true } }],
    }];
    const toggled = applyControlPatch(doc.blocks, 'p1', 0, {}, 'B');
    const run = toggled[0].type === 'paragraph' ? toggled[0].runs[0] : undefined;
    expect(run?.text).toBe('B');
    expect(run?.control?.placeholder).toBe(false);
    expect(controlChoiceIndex(run?.control, 'B')).toBe(1);
    expect(controlChoiceIndex(run?.control, 'C')).toBe(-1);

    expect(formatControlDate('2026-10-30', 'yyyy年M月d日')).toBe('2026年10月30日');
    expect(formatControlDate('2026-03-05', 'yyyy-MM-dd')).toBe('2026-03-05');
    expect(formatControlDate('2026-03-05', 'M月d日')).toBe('3月5日');
  });

  it('编辑回读保留方框勾选与控件，不因改字丢失', () => {
    const root = document.createElement('div');
    root.innerHTML =
      '<span data-vml-rect="1" data-vml-checked="1" data-vml-left="10" data-vml-top="2" data-vml-width="13" data-vml-height="13" data-vml-xrel="page" data-vml-stroke="#000000" data-vml-stroke-width="1"></span>' +
      '<span data-control="dropdown" data-run-index="1" data-control-value="A" data-control-choices="A\u0001B"></span>' +
      '正文';
    const runs = serializeDomRuns(root, {}, new Map());
    const rectRun = runs.find((run) => run.image?.vmlRect);
    expect(rectRun?.image?.vmlRect?.checked).toBe(true);
    const controlRun = runs.find((run) => run.control);
    expect(controlRun?.control).toMatchObject({ kind: 'dropdown', choices: ['A', 'B'] });
    expect(controlRun?.text).toBe('A');
    expect(runs.some((run) => run.text === '正文')).toBe(true);
  });

  it('表格右侧/底部加号：整表选区下追加一列或一行', () => {
    const doc = createBlankDoc();
    const blank = () => ({ id: `c${Math.random()}`, blocks: [], colSpan: 1, rowSpan: 1 });
    doc.blocks = [{
      type: 'table',
      id: 't1',
      columnWidthsPx: [100, 100],
      rows: [
        { id: 'r1', cells: [blank(), blank()] },
        { id: 'r2', cells: [blank(), blank()] },
      ],
    }];
    const table = doc.blocks[0];
    if (table.type !== 'table') throw new Error('table expected');
    const placements = tableCellPlacements(table);
    const selection = { tableId: 't1', anchorId: placements[0].cell.id, focusId: placements[placements.length - 1].cell.id };

    const withRow = editTable(table, selection, 'row-below', 600);
    expect(withRow?.rows).toHaveLength(3);
    expect(withRow?.rows[2].cells).toHaveLength(2);

    const withColumn = editTable(table, selection, 'column-after', 600);
    expect(withColumn?.columnWidthsPx).toHaveLength(3);
    expect(withColumn?.rows[0].cells).toHaveLength(3);
  });
});
