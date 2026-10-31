import type { DocBlock, DocRun, RunControl } from '@mini-office/document-model';
import { mapParagraphs } from './doc-utils';

/**
 * 功能：按格式串把日期写成控件显示值——支持 Word 日期控件常见的 yyyy/yy/M/M月/d/d日 组合，
 * 未识别字符原样保留；时间：2026-10-30；作者：hou0。
 */
export function formatControlDate(isoDate: string, format?: string): string {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(isoDate.trim());
  if (!match) return isoDate;
  const [, yyyy, mm, dd] = match;
  const pattern = format && format.trim() ? format : 'yyyy年M月d日';
  const month = Number(mm);
  const day = Number(dd);
  return pattern
    .replace(/yyyy/g, yyyy)
    .replace(/yy/g, yyyy.slice(2))
    .replace(/MM/g, mm)
    .replace(/M(?![a-zA-Z])/g, String(month))
    .replace(/dd/g, dd)
    .replace(/d(?![a-zA-Z])/g, String(day));
}

/** 功能：下拉/组合框的当前选项在候选中的位置，用于弹层打勾；时间：2026-10-30；作者：hou0。 */
export function controlChoiceIndex(control: RunControl | undefined, value: string): number {
  if (!control?.choices?.length) return -1;
  return control.choices.findIndex((choice) => choice === value);
}

/**
 * 功能：就地修改段落里某个内容控件的状态或取值（纯函数，便于测试与撤销）；
 * patch 合并进 control，value 给了就同时替换显示文本并清掉占位标记；时间：2026-10-30；作者：hou0。
 */
export function applyControlPatch(
  blocks: DocBlock[],
  paragraphId: string,
  runIndex: number,
  patch: Partial<RunControl>,
  value?: string,
): DocBlock[] {
  return mapParagraphs(blocks, (paragraph) => {
    if (paragraph.id !== paragraphId) return paragraph;
    if (runIndex < 0 || runIndex >= paragraph.runs.length) return paragraph;
    const runs = paragraph.runs.map((run, index) => {
      if (index !== runIndex || !run.control) return run;
      const next: DocRun = { ...run, control: { ...run.control, ...patch } };
      if (value !== undefined) {
        next.text = value;
        next.control = { ...next.control!, placeholder: false };
      }
      return next;
    });
    return { ...paragraph, runs };
  });
}
