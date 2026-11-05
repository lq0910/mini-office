# GenOffice DOC 兼容代码

功能：固定上游 DOC 转换与 XML 保留工具的来源，支持离线重建；时间：2026-11-03；集成作者：lq。

- 来源：https://github.com/genspark-ai/genoffice
- 提交：`e1f50e90d246f3ecc6427d2a0af90f04f37810ec`
- 许可：Apache-2.0，保留本目录 LICENSE、NOTICE 及上游代码注释。
- 范围：`packages/doc2docx/src`，以及其 DOCX 写入所需模块、正文 XML 扫描和文本补丁模块。
- 适配：通过精简入口避免引入整个应用；Node zlib 替换为 fflate 的浏览器解压；JSZip/CFB/XML 解析复用应用依赖。
- 构建：`node scripts/build-legacy-doc.mjs`。生成独立的按需加载模块，不需要 Electron 或 LibreOffice。
- 以固定上游实现为起点；每个文件新增的首行注释标明本项目的集成人员，不改变原作者归属。

## 本地兼容修正（2026-11-05）

功能：登记相对固定上游的增量，便于重建和以后升级时合并；时间：2026-11-05；作者：lq。

- `doc2docx/fonts.ts`、`index.ts`、`docx-writer.ts`：保留 FFN 主名与显式备用名，写入带关系和类型声明的 `fontTable.xml`。
- `doc2docx/formatting.ts`、`docx-writer.ts`：保留孤行控制和网格开关的显式关闭值；`docx-engine/generate.ts` 将网格开关纳入三态段落标志生成。
- `doc2docx/lists.ts`：补 NFC 37 到 `chineseCounting` 的映射。
- `doc2docx/drawings.ts`、`docx-writer.ts`：保留形状的 `posh` 水平对齐和 `WrapText` 不折行属性，恢复自动适应页码框的实际定位。
- Mini Office 的 `packages/editor-word/src/word-font-metrics.ts` 摘取同一提交的 `packages/font-metrics/word-faces.json` 中五个常用中文字体族设计行高，并按 `src/resolve.ts` 的东亚字体规则换算。不包含字体文件，未引入整个字体索引或编辑内核。
