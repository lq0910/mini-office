/** 功能：集成 GenOffice DOC 兼容代码；时间：2026-11-03；集成作者：lq；上游作者及 Apache-2.0 许可见本目录 LICENSE。 */
/**
 * The Word 97-2003 font table (SttbfFfn, FIB pair 15): an STTB whose strings
 * are FFN records. Each FFN starts with cbFfnM1 (its size minus one), then
 * ffid, wWeight, chs, ixchSzAlt, a 10-byte PANOSE and a 24-byte FontSignature,
 * and the UTF-16 font name (xszFfn, null-terminated) at offset 40. Character
 * runs refer to fonts by their index in this table (ftc). A malformed table
 * keeps the names read so far and records a warning.
 */
import type { FibInfo } from './fib'

const FFN_NAME_OFFSET = 40
const MAX_FONTS = 0x7fff

/** 功能：保留 DOC 字体主名及 FFN 声明的备用名，供 OOXML 字体表回写；时间：2026-11-05；作者：lq。 */
export interface FontEntry {
  name: string
  alternateName?: string
}

function u16At(bytes: Uint8Array, offset: number): number {
  return bytes[offset]! | (bytes[offset + 1]! << 8)
}

/** Font names by ftc; empty when the document has no font table. */
export function parseFontTable(fib: FibInfo, table: Uint8Array, warnings: string[], entries?: FontEntry[]): string[] {
  const { fcSttbfFfn: fc, lcbSttbfFfn: lcb } = fib
  if (lcb === 0) return []
  const names: string[] = []
  const warn = (detail: string): void => {
    warnings.push(
      'the font table is invalid (' + detail + '); font names after that point are unknown',
    )
  }
  if (fc > table.length || lcb > table.length - fc || lcb < 4) {
    warn('its range lies outside the table stream')
    return names
  }
  const end = fc + lcb
  const count = u16At(table, fc)
  const cbExtra = u16At(table, fc + 2)
  if (count > MAX_FONTS) {
    warn('it declares ' + count + ' fonts')
    return names
  }
  let offset = fc + 4
  for (let index = 0; index < count; index++) {
    if (offset + 1 > end) {
      warn('font ' + index + ' is missing')
      break
    }
    const size = table[offset]! + 1
    if (offset + size > end) {
      warn('font ' + index + ' overruns the table')
      break
    }
    let name = ''
    for (let at = offset + FFN_NAME_OFFSET; at + 2 <= offset + size; at += 2) {
      const code = u16At(table, at)
      if (code === 0) break
      name += String.fromCharCode(code)
    }
    names.push(name)
    // 功能：ixchSzAlt 是相对 xszFfn 的 UTF-16 字符索引；必须限制在本条记录内，不能误读下一字体；时间：2026-11-05；作者：lq。
    const alternateIndex = size >= FFN_NAME_OFFSET ? table[offset + 5]! : 0
    let alternateName = ''
    if (alternateIndex > 0) {
      const start = offset + FFN_NAME_OFFSET + alternateIndex * 2
      if (start + 2 <= offset + size) {
        for (let at = start; at + 2 <= offset + size; at += 2) {
          const code = u16At(table, at)
          if (code === 0) break
          alternateName += String.fromCharCode(code)
        }
      }
    }
    entries?.push({ name, ...(alternateName && alternateName !== name ? { alternateName } : {}) })
    offset += size + cbExtra
  }
  return names
}

/** The resolved font names of a run (w:rFonts ascii/hAnsi, eastAsia, cs). */
export interface FontNames {
  ascii?: string
  eastAsia?: string
  cs?: string
}

/** Default ftc values (STSHI rgftcStandardChpStsh) used when a run names no font. */
export interface DefaultFontIndices {
  ascii?: number
  eastAsia?: number
  other?: number
  cs?: number
}

/** The font table plus the stylesheet defaults the resolver falls back to. */
export interface FontTable {
  names: string[]
  defaults: DefaultFontIndices
}

/** Map the ftc indices of a run to names; undefined slots stay unset. */
export function fontNamesFor(
  fonts: FontTable,
  ftc: { ascii?: number; hAnsi?: number; eastAsia?: number; cs?: number },
): FontNames | undefined {
  const nameOf = (index: number | undefined): string | undefined => {
    if (index === undefined) return undefined
    const name = fonts.names[index]
    return name === undefined || name === '' ? undefined : name
  }
  const ascii = nameOf(ftc.ascii ?? ftc.hAnsi ?? fonts.defaults.ascii)
  const eastAsia = nameOf(ftc.eastAsia ?? fonts.defaults.eastAsia)
  const cs = nameOf(ftc.cs ?? fonts.defaults.cs)
  if (ascii === undefined && eastAsia === undefined && cs === undefined) return undefined
  return {
    ...(ascii !== undefined ? { ascii } : {}),
    ...(eastAsia !== undefined ? { eastAsia } : {}),
    ...(cs !== undefined ? { cs } : {}),
  }
}
