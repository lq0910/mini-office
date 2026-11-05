import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import {
  analyzeCapabilities,
  detectKind,
  loadOoxmlPackage,
} from '@mini-office/ooxml';
import { parseDocxPackage } from '../packages/editor-word/src/parse';
import { serializeDocx } from '../packages/editor-word/src/serialize';
import type { DocBlock, DocTable } from '@mini-office/document-model';
import { parseWorkbookPackage } from '../packages/editor-sheet/src/parse';
import { serializeXlsx } from '../packages/editor-sheet/src/serialize';

function fixture(path: string): Uint8Array {
  return new Uint8Array(readFileSync(path));
}

function allText(blocks: DocBlock[]): string {
  let text = '';
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      text += block.runs.map((run) => run.text ?? '').join('') + '\n';
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) text += allText(cell.blocks);
      }
    }
  }
  return text;
}

describe('DOCX fixture compatibility', () => {
  it('parses a real Chinese contract with table, image and page-number fields', async () => {
    if (typeof URL.createObjectURL !== 'function') {
      (URL as unknown as { createObjectURL: (blob: Blob) => string }).createObjectURL = () => 'blob:test';
      (URL as unknown as { revokeObjectURL: (url: string) => void }).revokeObjectURL = () => {};
    }
    const bytes = fixture('tests/fixtures/docx/contract.docx');
    const pkg = await loadOoxmlPackage(bytes);
    try {
      const detection = await detectKind(pkg, 'contract.docx');
      expect(detection.kind).toBe('docx');
      const capability = await analyzeCapabilities(pkg, 'docx', { fileName: 'contract.docx' });
      expect(capability.level).toBe(3);
      const doc = await parseDocxPackage(pkg, { fileName: 'contract.docx' });
      expect(doc.page.sizeName).toBe('A4');
      expect(Math.round(doc.page.widthPx)).toBe(794);
      // 功能：合同样本带 w:docGrid/@linePitch=360，排版需按 24px 行网格对齐；时间：2026-11-02；作者：lq。
      expect(doc.page.docGrid?.type).toBe('lines');
      expect(Math.round(doc.page.docGrid?.linePitchPx ?? 0)).toBe(24);
      const text = allText(doc.blocks);
      expect(text).toContain('房屋租赁合同');
      expect(text).toContain('北京示例科技有限公司');
      expect(text).toContain('2029 年 10 月 31 日');
      expect(text).toContain('合计：人民币柒万肆仟肆佰元整');
      expect(doc.header?.blocks.length).toBeGreaterThan(0);
      expect(allText(doc.header?.blocks ?? [])).toContain('Mini Office 合同样例');
      const footerRuns = doc.footer?.blocks
        .filter((block): block is Extract<DocBlock, { type: 'paragraph' }> => block.type === 'paragraph')
        .flatMap((paragraph) => paragraph.runs) ?? [];
      expect(footerRuns.some((run) => run.field === 'PAGE')).toBe(true);
      expect(footerRuns.some((run) => run.field === 'NUMPAGES')).toBe(true);
      const table = doc.blocks.find((block): block is DocTable => block.type === 'table');
      expect(table).toBeTruthy();
      expect(table!.rows.length).toBe(4);
      expect(table!.rows.length > 0).toBe(true);
      const image = doc.blocks
        .filter((block): block is Extract<DocBlock, { type: 'paragraph' }> => block.type === 'paragraph')
        .flatMap((paragraph) => paragraph.runs)
        .find((run) => run.image);
      expect(image?.image?.widthPx).toBeGreaterThan(50);
      expect(doc.stats?.words).toBeGreaterThan(100);
    } finally {
      pkg.dispose();
    }
  });

  it('round-trips the contract while preserving headers, images and table content', async () => {
    if (typeof URL.createObjectURL !== 'function') {
      (URL as unknown as { createObjectURL: (blob: Blob) => string }).createObjectURL = () => 'blob:test';
      (URL as unknown as { revokeObjectURL: (url: string) => void }).revokeObjectURL = () => {};
    }
    const bytes = fixture('tests/fixtures/docx/contract.docx');
    const pkg = await loadOoxmlPackage(bytes);
    const doc = await parseDocxPackage(pkg, { fileName: 'contract.docx' });
    pkg.dispose();
    const saved = await serializeDocx(doc, { sourceBytes: bytes, name: 'contract.docx' });
    expect(saved.byteLength).toBeGreaterThan(1000);
    const savedPkg = await loadOoxmlPackage(saved);
    try {
      expect(savedPkg.has('word/header1.xml') || savedPkg.has('word/header2.xml')).toBe(true);
      const reparsed = await parseDocxPackage(savedPkg, { fileName: 'contract.docx' });
      const text = allText(reparsed.blocks);
      expect(text).toContain('房屋租赁合同');
      expect(text).toContain('合计：人民币柒万肆仟肆佰元整');
      expect(allText(reparsed.header?.blocks ?? [])).toContain('Mini Office 合同样例');
      const table = reparsed.blocks.find((block): block is DocTable => block.type === 'table');
      expect(table?.rows.length).toBe(4);
      const image = reparsed.blocks
        .filter((block): block is Extract<DocBlock, { type: 'paragraph' }> => block.type === 'paragraph')
        .flatMap((paragraph) => paragraph.runs)
        .find((run) => run.image);
      expect(image?.image?.part).toContain('media/');
    } finally {
      savedPkg.dispose();
    }
  });
});

describe('XLSX fixture compatibility', () => {
  it('parses shared strings, dates, formulas, styles, merges and sheet metadata', async () => {
    const bytes = fixture('tests/fixtures/xlsx/workbook.xlsx');
    const pkg = await loadOoxmlPackage(bytes);
    try {
      const detection = await detectKind(pkg, 'workbook.xlsx');
      expect(detection.kind).toBe('xlsx');
      const capability = await analyzeCapabilities(pkg, 'xlsx', { fileName: 'workbook.xlsx' });
      expect(capability.level).toBeGreaterThanOrEqual(2);
      const workbook = await parseWorkbookPackage(pkg, 'workbook.xlsx');
      expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(['合同台账', '汇总', '公式验证']);
      const ledger = workbook.sheets[0];
      expect(ledger.cells['0:0']?.value).toBe('合同台账');
      expect(ledger.cells['2:1']?.value).toBe('北京示例科技');
      expect(ledger.cells['2:2']?.value).toBe(12000);
      expect(ledger.merges).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 5 }]);
      expect(ledger.freeze).toEqual({ rows: 2, cols: 0 });
      expect(ledger.autoFilter?.r1).toBe(1);
      expect(ledger.colWidths[0]).toBeGreaterThan(80);
      expect(ledger.rowHeights[0]).toBeGreaterThan(30);
      const amountStyle = workbook.styles[ledger.cells['2:2']?.styleId ?? ''];
      expect(amountStyle?.numFmt).toBe('¥#,##0.00');
      const dateStyle = workbook.styles[ledger.cells['2:3']?.styleId ?? ''];
      expect(dateStyle?.numFmt).toContain('yyyy');
      expect(workbook.computed?.[ledger.id]?.['2:3']?.display).toBe('2026-10-18');
      expect(workbook.computed?.[ledger.id]?.['1:8']?.value).toBe(95800.5);
      const summary = workbook.sheets[1];
      expect(workbook.computed?.[summary.id]?.['1:1']?.value).toBe(12000);
      expect(workbook.computed?.[summary.id]?.['3:1']?.value).toBe(42000);
      const formulas = workbook.sheets[2];
      expect(workbook.computed?.[formulas.id]?.['0:1']?.value).toBe('通过');
      expect(workbook.computed?.[formulas.id]?.['1:1']?.value).toBe(20);
      expect(workbook.computed?.[formulas.id]?.['2:1']?.value).toBe(30);
      expect(workbook.computed?.[formulas.id]?.['3:1']?.value).toBe(0.33);
      expect(workbook.computed?.[formulas.id]?.['4:1']?.value).toBe('合同');
      expect(workbook.computed?.[formulas.id]?.['5:1']?.value).toBe('安全');
    } finally {
      pkg.dispose();
    }
  });

  it('round-trips the real workbook with in-place save semantics', async () => {
    const bytes = fixture('tests/fixtures/xlsx/workbook.xlsx');
    const pkg = await loadOoxmlPackage(bytes);
    const workbook = await parseWorkbookPackage(pkg, 'workbook.xlsx');
    pkg.dispose();
    const saved = await serializeXlsx(workbook, { sourceBytes: bytes, name: 'workbook.xlsx' });
    const savedPkg = await loadOoxmlPackage(saved);
    try {
      expect(savedPkg.has('xl/calcChain.xml')).toBe(false);
      const reparsed = await parseWorkbookPackage(savedPkg, 'workbook.xlsx');
      expect(reparsed.sheets.map((sheet) => sheet.name)).toEqual(['合同台账', '汇总', '公式验证']);
      const summary = reparsed.sheets[1];
      expect(reparsed.computed?.[summary.id]?.['1:1']?.value).toBe(12000);
      expect(reparsed.computed?.[summary.id]?.['3:1']?.value).toBe(42000);
      const ledger = reparsed.sheets[0];
      expect(ledger.merges).toHaveLength(1);
      expect(ledger.freeze).toEqual({ rows: 2, cols: 0 });
      expect(ledger.cells['2:2']?.value).toBe(12000);
    } finally {
      savedPkg.dispose();
    }
  });

  // 功能：回归验证 XLSX 编辑保存不丢失未知部件；时间：2026-10-28；作者：lq。
  it('preserves unknown XLSX parts during save', async () => {
    const bytes = fixture('tests/fixtures/xlsx/workbook.xlsx');
    const source = await (await import('jszip')).default.loadAsync(bytes);
    source.file('custom/unknown.bin', new Uint8Array([4, 2, 4, 2]));
    const withUnknown = await source.generateAsync({ type: 'uint8array' });
    const pkg = await loadOoxmlPackage(withUnknown);
    const workbook = await parseWorkbookPackage(pkg, 'workbook.xlsx');
    pkg.dispose();
    const saved = await serializeXlsx(workbook, { sourceBytes: withUnknown, name: 'edited.xlsx' });
    const output = await (await import('jszip')).default.loadAsync(saved);
    expect(await output.file('custom/unknown.bin')?.async('uint8array')).toEqual(new Uint8Array([4, 2, 4, 2]));
  });
});
