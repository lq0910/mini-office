/** 功能：页眉页脚浮动文本框以独立受控投影编辑，保留原框定位、其他区域及动态域；时间：2026-11-11；作者：lq。 */
import type { HeaderFooter, HfParagraph } from '../upstream/packages/docx-engine/src/index';
import { PAGE_MARK } from '../upstream/packages/docx-engine/src/index';

/** 功能：编辑身份绑定原框 ID；保存标记只记录发生内容变化的原有形状；时间：2026-11-11；作者：lq。 */
export type WordBoxFooter = HeaderFooter & { miniOfficeBoxEdits?: number[] };

/** 功能：将指定框内段落投影为普通区域 schema，隐藏其他框及正文锚段；时间：2026-11-11；作者：lq。 */
export function wordHfBoxProjection(value: HeaderFooter, boxId?: number): HeaderFooter {
  if (boxId === undefined) return value;
  const paras = value.paras?.filter(para => para.box?.id === boxId).map(para => {
    const { box: _box, boxAnchored: _anchored, ...editable } = para;
    // 功能：正文外的页码框不能依赖 .doc-page 的样式选择器，显式投影已解析的继承字符格式；时间：2026-11-11；作者：lq。
    return { ...editable, runs: para.runs.map(run => ({
      ...run,
      sizeHalfPoints: run.sizeHalfPoints ?? run.effective?.sizeHalfPoints,
      fontAscii: run.fontAscii ?? run.effective?.fonts.ascii,
      eastAsiaFont: run.eastAsiaFont ?? run.effective?.fonts.eastAsia,
      bold: run.bold ?? run.effective?.bold, italic: run.italic ?? run.effective?.italic,
    })) };
  });
  if (!paras?.length) throw new Error('原页眉页脚文本框已不存在');
  return { text: paras.map(p => p.runs.map(r => r.text).join('')).join(''), paras };
}

/** 功能：局部编辑结果回填原框，不展开为普通页脚，新增段落沿用原框几何；时间：2026-11-11；作者：lq。 */
export function mergeWordHfBox(base: HeaderFooter, boxId: number | undefined, value: HeaderFooter): HeaderFooter {
  if (boxId === undefined) return value;
  const templates = base.paras?.filter(p => p.box?.id === boxId) ?? [];
  if (!templates.length || !templates[0].box?.miniOfficeShapeId) throw new Error('无法定位页眉页脚文本框来源，已取消修改');
  const replacement: HfParagraph[] = (value.paras ?? []).map((para, index) => ({ ...para, boxAnchored: true, box: templates[Math.min(index, templates.length - 1)].box }));
  let inserted = false;
  const paras = (base.paras ?? []).flatMap(para => {
    if (para.box?.id !== boxId) return [para];
    if (inserted) return [];
    inserted = true; return replacement;
  });
  const allRuns = paras.flatMap(p => [...p.runs, ...p.cells?.flatMap(c => c.paras.flat()) ?? []]);
  const { keepContent: _keep, pictureOnly: _picture, ...rest } = base;
  return { ...rest, paras, text: allRuns.map(r => r.text).join(''), pageNumber: allRuns.some(r => r.text.includes(PAGE_MARK)),
    newPictures: { ...base.newPictures, ...value.newPictures },
    miniOfficeBoxEdits: [...new Set([...(base as WordBoxFooter).miniOfficeBoxEdits ?? [], boxId])],
  } as WordBoxFooter;
}
