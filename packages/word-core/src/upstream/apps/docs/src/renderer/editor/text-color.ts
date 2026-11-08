/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */
import { dkColor } from './dark-page'

/** Word's automatic text colour resolves to the paper ink (the var already flips on the dark page) */
const AUTO_INK_CSS = 'var(--docs-paper-ink)'

/** CSS color value for a modeled run/style colour (hex without '#', or 'auto') */
export function textColorValue(color: string): string {
  return color === 'auto' ? AUTO_INK_CSS : `#${color}`
}

/** `color` declaration plus the dark-page twin for authored hex (auto needs none) */
export function textColorDecls(color: string): string[] {
  return color === 'auto' ? [`color:${AUTO_INK_CSS}`] : [`color:#${color}`, dkColor(color)]
}
