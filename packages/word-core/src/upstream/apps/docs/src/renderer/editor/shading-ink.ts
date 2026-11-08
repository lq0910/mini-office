/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */
/**
 * Word paints automatic-colour text white on dark shading. Probe (2026-09-05):
 * the cut is the classic integer luma (77R + 151G + 28B) / 256 < 60 — 3B3B3B,
 * C60000 and 006500 flip to white, 3C3C3C and C80000 stay black. Explicit run
 * or style colours never flip, and highlight is ignored by the test. The
 * nearest fill decides, so a light fill inside a dark one resets the ink.
 */
export type FillInk = 'light' | 'dark'

export function fillInk(hex: string | null | undefined): FillInk | undefined {
  if (!hex) return undefined
  const s = hex.startsWith('#') ? hex.slice(1) : hex
  if (!/^[0-9a-fA-F]{6}$/.test(s)) return undefined
  const n = parseInt(s, 16)
  const luma = (77 * ((n >> 16) & 255) + 151 * ((n >> 8) & 255) + 28 * (n & 255)) / 256
  return luma < 60 ? 'light' : 'dark'
}

export function isDarkFill(hex: string | null | undefined): boolean {
  return fillInk(hex) === 'light'
}
