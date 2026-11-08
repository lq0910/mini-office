/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */
import { t } from './i18n/locale'

let current = ''
const listeners = new Set<() => void>()

/** Settings → General user name; the generic label until the shell has answered */
export function authorName(): string {
  return current || t('editorDefaultAuthor')
}

export function setAuthorName(name: string): void {
  const next = name.trim()
  if (next === current) return
  current = next
  for (const listener of listeners) listener()
}

export function onAuthorNameChanged(listener: () => void): () => void {
  listeners.add(listener)
  return () => {
    listeners.delete(listener)
  }
}
