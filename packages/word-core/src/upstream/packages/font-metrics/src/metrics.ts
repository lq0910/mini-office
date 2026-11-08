/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */
export interface FaceVerticalMetrics {
  unitsPerEm: number
  hheaAscender: number
  hheaDescender: number
  hheaLineGap: number
  typoAscender: number | null
  typoDescender: number | null
  typoLineGap: number | null
  winAscent: number | null
  winDescent: number | null
  /** OS/2 fsSelection bit 7 (USE_TYPO_METRICS) */
  useTypoMetrics: boolean
  xAvgCharWidth: number | null
  /** OS/2 ulCodePageRange1; bits 17-21 flag the East Asian code pages */
  codePageRange1: number | null
}