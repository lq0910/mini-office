import { describe, expect, it } from 'vitest';
import { createBlankDoc, type DocParagraph } from '@mini-office/document-model';
import { parseDocx, serializeDocx } from '@mini-office/editor-word';

async function roundTrip(style: Record<string, unknown>) {
  const doc = createBlankDoc();
  doc.blocks[0] = { ...(doc.blocks[0] as DocParagraph), style };
  const bytes = await serializeDocx(doc, { name: 'tab.docx' });
  const reopened = await parseDocx(bytes, { fileName: 'tab.docx' });
  return reopened.blocks[0] as DocParagraph;
}

describe('tabStops parse/serialize round-trip', () => {
  it('round-trips a left tab stop', async () => {
    const para = await roundTrip({ tabStops: [{ posPx: 100, kind: 'left' }] });
    expect(para.style?.tabStops).toEqual([{ posPx: 100, kind: 'left' }]);
  });

  it('round-trips a right tab stop', async () => {
    const para = await roundTrip({ tabStops: [{ posPx: 200, kind: 'right' }] });
    expect(para.style?.tabStops).toEqual([{ posPx: 200, kind: 'right' }]);
  });

  it('round-trips multiple tab stops in order', async () => {
    const para = await roundTrip({
      tabStops: [
        { posPx: 100, kind: 'left' },
        { posPx: 300, kind: 'center' },
        { posPx: 500, kind: 'decimal' },
      ],
    });
    expect(para.style?.tabStops).toEqual([
      { posPx: 100, kind: 'left' },
      { posPx: 300, kind: 'center' },
      { posPx: 500, kind: 'decimal' },
    ]);
  });
});
