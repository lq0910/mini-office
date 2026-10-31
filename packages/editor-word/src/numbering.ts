import type { DocBlock, DocModel, NumberingLevel } from '@mini-office/document-model';
import { mapParagraphs } from './doc-utils';

/** 功能：列表编号格式化——与解析端一致的十进制/字母/罗马/中文序号；时间：2026-10-31。 */
function toLetters(value: number, upper: boolean): string {
  let result = '';
  let current = value;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode((upper ? 65 : 97) + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }
  return result || 'A';
}

function toRoman(value: number, upper: boolean): string {
  const table: Array<[number, string]> = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let result = '';
  let current = value;
  for (const [arabic, roman] of table) {
    while (current >= arabic) {
      result += roman;
      current -= arabic;
    }
  }
  return upper ? result.toUpperCase() : result;
}

function toChineseNumber(value: number): string {
  const digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
  const units = ['', '十', '百', '千'];
  if (value <= 0) return String(value);
  if (value < 10) return digits[value];
  if (value === 10) return '十';
  if (value < 20) return '十' + digits[value % 10];
  const text = String(value);
  let result = '';
  const length = text.length;
  for (let index = 0; index < length; index += 1) {
    const digit = Number(text[index]);
    const unit = units[length - index - 1];
    if (digit === 0) {
      if (!result.endsWith('零') && index < length - 1) result += '零';
      continue;
    }
    result += digits[digit] + unit;
  }
  return result.replace(/零+$/, '');
}

export function formatListNumber(value: number, format: string): string {
  switch (format) {
    case 'lowerLetter':
      return toLetters(value, false);
    case 'upperLetter':
      return toLetters(value, true);
    case 'lowerRoman':
      return toRoman(value, false);
    case 'upperRoman':
      return toRoman(value, true);
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'japaneseCounting':
      return toChineseNumber(value);
    case 'decimalZero':
      return String(value).padStart(2, '0');
    case 'decimal':
    default:
      return String(value);
  }
}

/**
 * 功能：按编号定义与计数器计算段落标记——与解析端 listInfoFor 同一套规则
 * （计数器递增、更深层级重置、lvlText 的 %N 替换）；时间：2026-10-31。
 */
export function computeListMarker(levels: NumberingLevel[], counters: Map<number, number>, ilvl: number): string {
  const level = levels.find((item) => item.ilvl === ilvl) ?? levels[0];
  if (!level) return '';
  const current = (counters.get(ilvl) ?? level.start - 1) + 1;
  counters.set(ilvl, current);
  for (const key of [...counters.keys()]) {
    if (key > ilvl) counters.delete(key);
  }
  const values = new Map<number, number>();
  for (const item of levels) {
    values.set(item.ilvl, counters.get(item.ilvl) ?? item.start);
  }
  values.set(ilvl, current);
  if (level.numFmt === 'none') return '';
  if (level.numFmt === 'bullet') return level.lvlText || '•';
  let marker = level.lvlText.replace(/%([1-9])/g, (_match, index: string) => {
    const target = Number(index) - 1;
    const targetLevel = levels.find((item) => item.ilvl === target);
    return formatListNumber(values.get(target) ?? targetLevel?.start ?? 1, targetLevel?.numFmt ?? 'decimal');
  });
  if (!marker) marker = `${formatListNumber(current, level.numFmt)}.`;
  return marker;
}

/** 功能：无编号定义时的文本增量兜底——把标记里的首个数字串 +1；时间：2026-10-31。 */
function incrementMarkerText(marker: string, next: number): string {
  const match = /\d+/.exec(marker);
  if (!match) return `${next}.`;
  return marker.slice(0, match.index) + String(next) + marker.slice(match.index + match[0].length);
}

/**
 * 功能：全文档列表重编号——分段/粘贴/插入段落后调用，保证列表标记连续递增
 * （此前新段落继承源段落的静态标记，回车永远出同一个编号）；时间：2026-10-31。
 */
export function renumberListMarkers(doc: DocModel): DocModel {
  const definitions = doc.numberingLevels;
  const counters = new Map<string, Map<number, number>>();
  const fallbackCounters = new Map<string, number>();
  let anyList = false;
  const renumberBlocks = (blocks: DocBlock[]): DocBlock[] =>
    mapParagraphs(blocks, (paragraph) => {
      const list = paragraph.list;
      if (!list?.numId) return paragraph;
      anyList = true;
      const levels = definitions?.[list.numId];
      let marker: string;
      if (levels?.length) {
        let countersForNum = counters.get(list.numId);
        if (!countersForNum) {
          countersForNum = new Map();
          counters.set(list.numId, countersForNum);
        }
        marker = computeListMarker(levels, countersForNum, list.level);
      } else {
        const next = (fallbackCounters.get(list.numId) ?? 0) + 1;
        fallbackCounters.set(list.numId, next);
        marker = incrementMarkerText(list.marker, next);
      }
      if (marker === list.marker) return paragraph;
      return { ...paragraph, list: { ...list, marker } };
    });

  const walkBlocks = (blocks: DocBlock[]): DocBlock[] => {
    const next = renumberBlocks(blocks);
    return next.map((block) => {
      if (block.type !== 'table') return block;
      return {
        ...block,
        rows: block.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({ ...cell, blocks: walkBlocks(cell.blocks) })),
        })),
      };
    });
  };

  const blocks = walkBlocks(doc.blocks);
  return anyList ? { ...doc, blocks } : doc;
}
