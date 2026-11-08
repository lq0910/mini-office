/** 功能：完整导出配套的 Word 排版、测量、分页与字体规则；时间：2026-11-08；作者：lq。 */
export * from '../upstream/apps/docs/src/renderer/pagination';
export * from '../upstream/apps/docs/src/renderer/pagination-plan';
export { FollowUpChain } from '../upstream/apps/docs/src/renderer/pagination-loop';
export { BlockIndex } from '../upstream/apps/docs/src/renderer/pagination-index';
export { docStyleCss, docThemeCss, docBodyFont, docHasCjk } from '../upstream/apps/docs/src/renderer/doc-style-css';
export { adoptEmbeddedFonts } from '../upstream/apps/docs/src/renderer/embedded-fonts';
export { prefetchFaces, setResolverFontTable, setResolverEmbeddedFaces } from '../upstream/apps/docs/src/renderer/font-resolve';
export { WordLayoutController, measureWordLayout, type WordLayoutInput, type WordLayoutResult } from './controller';
