/** 功能：现有工具栏格式摘要与新核心原生单位之间的局部转换；时间：2026-11-10；作者：lq。 */
import type { ParagraphStyle, RunStyle } from '@mini-office/document-model';
import { applyWordParagraphFormat, applyWordTextFormat, readWordFormatting, type Editor, type WordParagraphFormat, type WordTextFormat } from '../../word-core/src/editing/index';
import type { WordCoreSession } from '../../word-core/src/index';

/** 功能：只转换格式键，不转换正文、表格或文档模型；时间：2026-11-10；作者：lq。 */
const textKeys: Partial<Record<keyof WordTextFormat, keyof RunStyle>> = {
  font: 'fontFamily', sizeHalfPoints: 'fontSizePt', bold: 'bold', italic: 'italic', underline: 'underline', strike: 'strike', color: 'color', highlight: 'highlight',
};
/** 功能：段落长度按 1px=15twips 转换，行距固定值与倍数分开处理；时间：2026-11-10；作者：lq。 */
const paragraphKeys = {
  indentLeftPx: 'indentLeft', indentRightPx: 'indentRight', firstLineIndentPx: 'indentFirstLine',
  spaceBeforePx: 'spaceBefore', spaceAfterPx: 'spaceAfter',
} as const;

/** 功能：当前选区摘要供原有工具栏与对话框使用，混合值单独反馈；时间：2026-11-10；作者：lq。 */
export function coreToolbarFormatting(core: WordCoreSession) {
  const format = readWordFormatting(core.commandEditor(), core.formattingSource());
  const text = format.text.value;
  const p = format.paragraph.value;
  const activeStyle: RunStyle = {
    fontFamily: text.font ?? text.fontAscii, fontSizePt: text.sizeHalfPoints === undefined ? undefined : text.sizeHalfPoints / 2,
    bold: text.bold, italic: text.italic, underline: text.underline, strike: text.strike, color: text.color, highlight: text.highlight,
  };
  const paragraph: ParagraphStyle = {
    alignment: p.align === 'distribute' ? 'justify' : p.align as ParagraphStyle['alignment'],
    pageBreakBefore: p.pageBreakBefore, keepNext: p.keepNext ?? undefined, keepLines: p.keepLines ?? undefined,
    widowControl: p.widowControl ?? undefined, snapToGrid: p.snapToGrid ?? undefined,
  };
  for (const [ui, native] of Object.entries(paragraphKeys)) {
    if (p[native as keyof typeof p] != null) (paragraph as Record<string, unknown>)[ui] = Number(p[native as keyof typeof p]) / 15;
  }
  if (p.lineRule === 'exact') paragraph.lineHeightExactPx = Number(p.lineRawTwips) / 15;
  else if (p.lineRule === 'atLeast') paragraph.lineHeightAtLeastPx = Number(p.lineRawTwips) / 15;
  else if (p.lineRule === 'auto') paragraph.lineHeight = p.lineSpacing ?? 1;
  const mixedText = format.text.mixed.map(key => textKeys[key]).filter((key): key is keyof RunStyle => !!key);
  if (format.text.mixed.includes('fontAscii') && !mixedText.includes('fontFamily')) mixedText.push('fontFamily');
  const mixedParagraph: Array<keyof ParagraphStyle> = [];
  for (const key of format.paragraph.mixed) {
    const mapped = Object.entries(paragraphKeys).find(([, native]) => native === key)?.[0];
    if (mapped) mixedParagraph.push(mapped as keyof ParagraphStyle);
    else if (key === 'align') mixedParagraph.push('alignment');
    else if (key.startsWith('line')) mixedParagraph.push('lineHeight', 'lineHeightExactPx', 'lineHeightAtLeastPx');
  }
  return { activeStyle, paragraph, mixedText, mixedParagraph };
}

/** 功能：显式选择字体同时更新中西文字体槽；清除颜色使用 auto，任意高亮色使用 OOXML 底纹；时间：2026-11-10；作者：lq。 */
export function coreFontFormat(editor: Editor, patch: Partial<RunStyle>): boolean {
  const attrs: Record<string, unknown> = {};
  if ('fontFamily' in patch) Object.assign(attrs, { font: patch.fontFamily ?? null, fontAscii: patch.fontFamily ?? null, eastAsiaFont: patch.fontFamily ?? null, eaSlotEmpty: null, themeRFonts: null });
  if ('fontSizePt' in patch && Number.isFinite(patch.fontSizePt) && patch.fontSizePt! > 0) attrs.sizeHalfPoints = Math.round(patch.fontSizePt! * 2);
  if ('color' in patch) Object.assign(attrs, { color: patch.color?.replace(/^#/, '') ?? 'auto', themeColor: null });
  if ('highlight' in patch) Object.assign(attrs, { highlight: null, shading: patch.highlight?.replace(/^#/, '') ?? null, shadingDisplay: null });
  return Object.keys(attrs).length > 0 && applyWordTextFormat(editor, attrs);
}

/** 功能：段落 UI 补丁转换为原生格式，清理冲突行距/字符缩进/自动间距；时间：2026-11-10；作者：lq。 */
export function coreParagraphFormat(editor: Editor, patch: Partial<ParagraphStyle>): boolean {
  const attrs: WordParagraphFormat = {};
  if ('alignment' in patch) attrs.align = patch.alignment ?? null;
  for (const [ui, native] of Object.entries(paragraphKeys)) {
    if (ui in patch) Object.assign(attrs, { [native]: patch[ui as keyof ParagraphStyle] == null ? null : Math.round(Number(patch[ui as keyof ParagraphStyle]) * 15) });
  }
  if ('firstLineIndentPx' in patch) attrs.indentFirstLineChars = null;
  if ('spaceBeforePx' in patch) attrs.spaceBeforeAuto = false;
  if ('spaceAfterPx' in patch) attrs.spaceAfterAuto = false;
  if (['lineHeight', 'lineHeightExactPx', 'lineHeightAtLeastPx'].some(key => key in patch)) {
    if (patch.lineHeightExactPx != null) Object.assign(attrs, { lineRule: 'exact', lineRawTwips: Math.round(patch.lineHeightExactPx * 15), lineSpacing: null });
    else if (patch.lineHeightAtLeastPx != null) Object.assign(attrs, { lineRule: 'atLeast', lineRawTwips: Math.round(patch.lineHeightAtLeastPx * 15), lineSpacing: null });
    else Object.assign(attrs, { lineRule: 'auto', lineRawTwips: Math.round((patch.lineHeight ?? 1) * 240), lineSpacing: patch.lineHeight ?? 1 });
  }
  for (const key of ['pageBreakBefore', 'keepNext', 'keepLines', 'widowControl', 'snapToGrid'] as const) if (key in patch) Object.assign(attrs, { [key]: patch[key] ?? null });
  return Object.keys(attrs).length > 0 && applyWordParagraphFormat(editor, attrs);
}
