/** 功能：融合 GenOffice Word 内核；时间：2026-11-08；集成作者：lq；上游作者、固定版本及本地修改见 packages/word-core/UPSTREAM.md 和 SOURCE-MANIFEST.json。 */
import { defineStrings } from '../../../../../packages/i18n/src/index'
import { zh } from './app/zh'
import { en } from './app/en'
import { ja } from './app/ja'
import { ko } from './app/ko'
import { fr } from './app/fr'
import { de } from './app/de'
import { es } from './app/es'
import { th } from './app/th'
import { id } from './app/id'
import { ru } from './app/ru'
import { ar } from './app/ar'
import { pt } from './app/pt'
import { it } from './app/it'
import { pl } from './app/pl'
import { cs } from './app/cs'
import { nl } from './app/nl'
import { ms } from './app/ms'
import { he } from './app/he'
import { hi } from './app/hi'
import { zhTW } from './app/zh-TW'
import { vi } from './app/vi'
import { tr } from './app/tr'

/** Strings for App.tsx / SettingsModal / non-Ribbon components (keys use the app prefix to mark the area) */
export const appStrings = defineStrings({
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
