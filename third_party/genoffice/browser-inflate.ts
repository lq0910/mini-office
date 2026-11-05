/** 功能：为 WebView 中的 DOC 图片解压提供 Node inflateSync 的兼容接口；时间：2026-11-03；作者：lq。 */
export { unzlibSync as inflateSync } from 'fflate';
