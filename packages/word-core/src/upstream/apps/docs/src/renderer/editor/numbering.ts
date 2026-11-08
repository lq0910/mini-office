/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */
// Implementation lives in the engine (parse-time TOC numbering reuses it);
// this module keeps the historical import path for the editor.
export {
  substituteBullet,
  markerFallbackFace,
  SEGOE_UI_SYMBOL_RE,
  computeListMarkerInfos,
  computeListMarkers,
  formatNumber,
  markerTabAdvance,
  type ListItemRef,
  type ListMarkerInfo,
} from '../../../../../packages/docx-engine/src/index'
