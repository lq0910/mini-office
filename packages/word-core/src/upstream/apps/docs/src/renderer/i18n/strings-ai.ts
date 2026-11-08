/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */
import { defineStrings } from '../../../../../packages/i18n/src/index'
import { zh } from './ai/zh'
import { en } from './ai/en'
import { ja } from './ai/ja'
import { ko } from './ai/ko'
import { fr } from './ai/fr'
import { de } from './ai/de'
import { es } from './ai/es'
import { th } from './ai/th'
import { id } from './ai/id'
import { ru } from './ai/ru'
import { ar } from './ai/ar'
import { pt } from './ai/pt'
import { it } from './ai/it'
import { pl } from './ai/pl'
import { cs } from './ai/cs'
import { nl } from './ai/nl'
import { ms } from './ai/ms'
import { he } from './ai/he'
import { hi } from './ai/hi'
import { zhTW } from './ai/zh-TW'
import { vi } from './ai/vi'
import { tr } from './ai/tr'

/** User-visible strings for the ai/ panel and tool feedback (LLM prompts are not here) */
export const aiStrings = defineStrings({
  zh,
  en,
  ja,
  ko,
  fr,
  de,
  es,
  th,
  id,
  ru,
  ar,
  pt,
  it,
  pl,
  cs,
  nl,
  ms,
  he,
  hi,
  'zh-TW': zhTW,
  vi,
  tr,
})
