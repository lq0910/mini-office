/** 功能：阻止缺失、陈旧、错误来源及越界路径的 Office 基准被计为通过；时间：2026-11-03；作者：lq。 */
import { describe, expect, it } from 'vitest';
import { baselinePath, validateOfficeBaseline } from '../scripts/word-fidelity-baseline.mjs';

const fixture = { id: 'grid-line-rules', sha256: 'a'.repeat(64) };
const baseline = {
  schemaVersion: 1, application: 'Microsoft Word', version: '16.x test', operatingSystem: 'test OS',
  exportedAt: '2026-11-03', exportMethod: 'Export PDF; render 96 DPI', dpi: 96, fonts: ['Calibri'],
  fixtureId: fixture.id, sourceSha256: fixture.sha256, pdfFile: 'source.pdf', pdfSha256: 'b'.repeat(64),
  pages: [{ file: 'page-001.png', sha256: 'c'.repeat(64) }],
};

describe('Office 基准绑定与边界', () => {
  it('接受来源、字体、原件和页面都明确的基准', () => {
    expect(validateOfficeBaseline(baseline, fixture)).toEqual(baseline);
  });
  it.each([
    { application: 'Chromium' }, { version: '' }, { operatingSystem: '' }, { exportMethod: '' },
    { fonts: [] }, { fonts: [''] }, { exportedAt: 'not a date' }, { dpi: 144 },
    { sourceSha256: 'd'.repeat(64) }, { fixtureId: 'other' }, { pdfFile: '../source.pdf' }, { pdfSha256: '' }, { pages: [] },
    { pages: [{ file: '../page.png', sha256: 'c'.repeat(64) }] },
    { pages: [{ file: '/tmp/page.png', sha256: 'c'.repeat(64) }] },
    { pages: [baseline.pages[0], baseline.pages[0]] },
  ])('拒绝不完整或不匹配的基准 %j', (patch) => {
    expect(() => validateOfficeBaseline({ ...baseline, ...patch }, fixture)).toThrow();
  });
  it.each(['../outside.png', '/tmp/outside.png', 'a/../../outside.png', 'C:\\outside.png'])('拒绝基准路径 %s', (file) => {
    expect(() => baselinePath('/tmp/baselines', file)).toThrow();
  });
});
