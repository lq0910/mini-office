/** 功能：声明固定版本 DOC 转换模块的应用接口；时间：2026-11-03；作者：lq。 */
export function convertLegacyDocToDocx(bytes: Uint8Array): Promise<{ docx: Uint8Array; warnings: string[] }>;
