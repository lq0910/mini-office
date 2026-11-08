/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */
import type { Editor } from '@tiptap/core';
import type { ParsedDocFull } from '../../../../packages/docx-engine/src/index';
import { setNoteNumFmts } from './note-format';
import { hangulSpaceWideningFor, setAdditiveParagraphSpacing, setHangulSpaceWidening, setLatinViaEastAsiaSlot, latinViaEastAsiaSlotFor } from './line-metrics';
import { docBodyEaFont } from './doc-style-css';
import { setStaticAsciiResolver } from './editor/protected-render';
export function applyDocLayoutSettings(editor: Editor, parsed: ParsedDocFull): void {
  setNoteNumFmts({ footnote: parsed.footnoteProps, endnote: parsed.endnoteProps })
  editor.storage.tabStops.defaultTabStopTwips = parsed.defaultTabStopTwips ?? null
  editor.storage.listNumbering.defaultTabTwips = parsed.defaultTabStopTwips ?? 720
  editor.storage.listNumbering.indentNotTabStop = parsed.indentNotNumberingTabStop === true
  // Word 2013+ justified lines pull words up by shrinking spaces; legacy
  // compatibility modes (and new blank docs) never do
  editor.storage.justifyShrink.enabled = (parsed.compatibilityMode ?? 0) >= 15
  // Word stretches a justified line that ends in a line break (w:br / w:cr)
  // unless the document opts out with doNotExpandShiftReturn
  editor.storage.justifyShrink.expandBreakLines = parsed.doNotExpandShiftReturn !== true
  setAdditiveParagraphSpacing(parsed.doNotUseHTMLParagraphAutoSpacing === true)
  editor.storage.cjkPunctShrink.enabled = parsed.compressPunctuation === true
  editor.storage.cjkPunctShrink.hangPunct = parsed.compressPunctuation !== true
  editor.storage.cjkPunctShrink.legacyLayout = (parsed.compatibilityMode ?? 0) < 15
  editor.storage.cjkPunctShrink.docEastAsiaLang = parsed.docDefaults?.eastAsiaLang ?? null
  setHangulSpaceWidening(hangulSpaceWideningFor(parsed))
  setLatinViaEastAsiaSlot(latinViaEastAsiaSlotFor(parsed), docBodyEaFont(parsed))
  setStaticAsciiResolver(() => editor.storage.listNumbering)
  // Chromium only hyphenates under an explicit lang (the app shell is zh-CN);
  // scoped to autoHyphenation docs so CJK font fallback is untouched elsewhere
  const lang = parsed.autoHyphenation ? parsed.docDefaults?.lang : undefined
  if (lang) editor.view.dom.setAttribute('lang', lang)
  else editor.view.dom.removeAttribute('lang')
}
