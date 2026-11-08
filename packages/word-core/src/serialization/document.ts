/** 功能：从完整事务文档生成 DOCX，不依赖产品 UI 或磁盘写入；时间：2026-11-08；作者：lq。 */
import {
  findChartWorkbookPath, parseChartPartXml, patchChartPartXml,
  patchChartWorkbookXlsxBase64, readDocxPartBase64, saveDocx,
  type ParsedDocFull, type SaveOptions, type NewInkImage,
} from '../upstream/packages/docx-engine/src/index';
import { pmDocToSavePlan, type PmNode } from '../upstream/apps/docs/src/renderer/editor/convert';

/** 功能：序列化事务快照，保留上游图表/嵌入工作簿的联动回写及来源计划；时间：2026-11-08；作者：lq。 */
export async function serializeWordDocument(parsed: ParsedDocFull, document: PmNode, options: SaveOptions = {}): Promise<Uint8Array> {
  const plan = pmDocToSavePlan(document, parsed.blocks);
  const partXml = { ...options.partXml };
  const partBinary = { ...options.partBinary };
  for (const { partPath, patch } of plan.chartPatches) {
    const original = parsed.extras.chartParts[partPath];
    if (!original) throw new Error(`无法定位原图表部件，拒绝有损保存：${partPath}`);
    const xml = patchChartPartXml(original, patch);
    partXml[partPath] = xml;
    const workbookPath = await findChartWorkbookPath(parsed.internal.originalBytes, partPath);
    if (!workbookPath) continue;
    const base64 = await readDocxPartBase64(parsed.internal.originalBytes, workbookPath);
    const display = parseChartPartXml(xml, partPath);
    if (!base64 || !display) throw new Error('无法同步图表的嵌入工作簿，原文件未被改写');
    const updated = await patchChartWorkbookXlsxBase64(base64, display.categories, display.series.map((series, index) => ({ name: series.name ?? `Series${index + 1}`, values: series.values as (number | null)[] })));
    if (!updated) throw new Error('图表工作簿回写失败，原文件未被改写');
    partBinary[workbookPath] = updated;
  }
  // 功能：重新生成锚段时仍保留既有墨迹 PNG 及矢量载荷，不强制重新栅格化；时间：2026-11-08；作者：lq。
  let inks = options.inks;
  if (inks === undefined && plan.changedCount > 0 && parsed.inks.length) {
    inks = [];
    for (const ink of parsed.inks) {
      const blockIndex = plan.saveBlockIndexByDocx.get(ink.anchorIndex);
      if (blockIndex === undefined) continue;
      const prefix = 'data:image/png;base64,';
      if (!ink.dataUrl?.startsWith(prefix)) throw new Error('既有墨迹缺少可保留的原图，拒绝有损保存');
      const entry: NewInkImage = { blockIndex, base64: ink.dataUrl.slice(prefix.length), offsetXPx: ink.offsetXPx, offsetYPx: ink.offsetYPx, widthPx: ink.widthPx, heightPx: ink.heightPx };
      if (ink.payload) entry.payload = ink.payload;
      inks.push(entry);
    }
  }
  return saveDocx(parsed, plan.saveBlocks, {
    ...options,
    newPictures: Object.keys(plan.newPictures).length ? { ...options.newPictures, ...plan.newPictures } : options.newPictures,
    trailingSectPrChange: plan.trailingSectPrChange ?? options.trailingSectPrChange,
    inks,
    partXml: Object.keys(partXml).length ? partXml : undefined,
    partBinary: Object.keys(partBinary).length ? partBinary : undefined,
  });
}
