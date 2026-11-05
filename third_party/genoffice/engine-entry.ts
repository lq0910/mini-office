/** 功能：仅导出 DOC 转换使用的 DOCX 生成接口，避免加载整套上游引擎；时间：2026-11-03；作者：lq。 */
export { buildBlankDocx } from './docx-engine/blank';
export { generateParagraphXml, mergeRPrModel } from './docx-engine/generate';
export type { GenerateContext } from './docx-engine/generate';
export type { GeneratedBlock, ParaFormat, Run } from './docx-engine/types';
