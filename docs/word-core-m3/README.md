# M3 工具栏融合验收记录

功能：记录现有工具栏、上下文操作、区域编辑及输出的新核心迁移结果和可复核证据；时间：2026-11-11；作者：lq。

**M3 已完成原计划的工具栏融合范围。** M0 清单 W01–W35 的编辑、状态及输出入口已接入同一个 `WordCoreSession`，保留 Mini Office 的 React 18、六标签工具栏、配色和现有控件。W02–W04 的打开/编辑闭环沿用 M2；桌面工作区与平台组合 W36–W38 按 M4/M5 验收。开发入口仍使用 `VITE_WORD_CORE_M2=true`；尚未切换正式默认内核、制作新安装包、提交、推送或发布。

## 原计划完成条件

- [x] `WordToolbar.tsx` 只接收格式摘要、历史状态及能力参数，新正文不转换成旧 `DocModel`。
- [x] 工具栏、菜单及快捷键执行原生事务；新核心正文不经旧 `execCommand` 或手工 DOM 回读修改。
- [x] 混合字符/段落格式、矩形表格选区、页眉页脚及文本框的当前活动编辑区正确回填状态。
- [x] 增补必要的表格、区域、页面和输出操作，使用本项目控件；没有引入上游 Ribbon 或第二套正文/撤销栈。

## W01–W35 功能与证据

下表的完成表示现有产品功能迁移完成，支持边界见后文，不表示任意 Word 结构均可重建。

| 编号 | 已完成的新核心行为 | 实现与验证入口 |
| --- | --- | --- |
| W01 | 六标签、功能区折叠、更多菜单及选区保持；工具栏弹层夺焦后命令仍作用于原选区 | `WordCoreEditor.tsx`、`WordToolbar.tsx`；formatting、objects 浏览器检查 |
| W02–W04 | 新建/打开、文件名、中文确认/取消、普通输入、软换行、跨段跨页选区及删除；保存/重开始终使用完整新核心 | `core-handler.ts`、`core-surface.ts`；M2 surface/workspace 24 项浏览器回归 |
| W05 | 复制/剪切、来源/目标/纯文本三种粘贴原位切换、图片剪贴板与文件解码；切换模式不重复内容，后续编辑令旧载荷失效 | `editing/paste.ts`、`picture-files.ts`、`WordCorePasteOptions.tsx`；paste-properties、header-footer、textbox |
| W06 | 正文、样式目录、页面、页眉页脚和文本框共享主历史；连续输入分组、独立工具操作单步撤销/重做；保存及标签切换不丢历史 | `editing/region-history.ts`、`session/word-session.ts`；各核心测试及区域/workspace 浏览器 |
| W07 | 当前编辑区查找、高亮、前后导航、大小写、替换/全部替换；跨字符格式命中，批量替换一步撤销，不修改其他编辑区 | `editing/search.ts`、`WordCoreFindBar.tsx`；search-lists、header-footer、textbox |
| W08–W11 | 系统字体选择、中文/西文字体槽、混合字体/字号、自定义字号、粗斜体/下划线/删除线、颜色/高亮及显式关闭继承格式 | `editing/formatting.ts`、`core-formatting.ts`；formatting、styles-painter 及区域浏览器 |
| W12 | 单次/连续格式刷、Esc 退出；字符、段落、边框/制表位和字符特效摘要跨正文/表格/区域应用，保留目标链接/语言和未修改来源 XML | `editing/format-painter.ts`；styles-painter 核心/浏览器、header-footer 跨区域鼠标拖选 |
| W13–W14 | 对齐、固定/最小/倍数行距、左右及首行/悬挂缩进、段前段后、分页属性；混合选区只提交本次改动字段 | `ParagraphDialog.tsx`、`editing/formatting.ts`；formatting、page-layout 及页眉格内真实段落对话框 |
| W15 | 正文/表格/页眉页脚/文本框项目符号和编号、取消、层级、Enter 续项/空项退出、Tab/Shift+Tab；新编号定义随捕获快照保存 | `editing/lists.ts`、`WordCoreSession.toggleList()`；search-lists、page-layout、textbox 核心及区域浏览器 |
| W16 | 样式应用/创建/改名/删除、样式集保存/导入/删除，支持列表/边框预设及冲突 ID 重分配；区域与正文共享定义和历史 | `editing/styles.ts`、`WordCoreStyles.tsx`；styles-painter、page-layout、textbox。删除导入样式时保留隐藏格式基底并迁移跨部件引用，未知效果继续继承 |
| W17 | 编辑标记开关只改变显示，不进入正文/保存历史 | `WordCoreEditor.tsx`、`core-surface.ts`；M2 surface 与输出清理检查 |
| W18 | 分节纸张、横纵向、四侧边距及正文页边距拖动；拖动预览不改版本，松手一次事务，Esc/取消不提交 | `editing/page-layout.ts`、`core-page-margins.ts`、`WordCorePagePanel.tsx`；page-layout 及 output |
| W19 | 分页符按钮/快捷键、软换行、按实际滚动/缩放更新当前页，保持文档版本不变 | 同源 `page-break.ts`、`WordCoreEditor.tsx`；formatting、page-layout、surface |
| W20 | 双击页眉/页脚空白或现有区域直接编辑，富文本/布局表格/图片/列表/域；首页不同、奇偶不同、链接/解除前节继承，动态页码/总数/起始值和距边设置 | `editing/header-footer.ts`、`WordCoreHeaderFooterDialog.tsx`、`core-header-footer.ts`；page-layout 核心、header-footer 浏览器。区域打开时即可保存，不依赖失焦 |
| W21 | 标尺、缩进拖动、制表位增删、缩放换算；正文/区域按实际净宽定位；1–3 级目录插入/原位更新/修饰键导航，页码采用真实分页和分节编号缓存 | `editing/ruler.ts`、`toc.ts`、`WordCoreRuler.tsx`；ruler-toc、page-layout、header-footer、textbox。重测不产生 dirty/撤销项 |
| W22–W23 | 图片文件插入、选中/删除、尺寸拖动/输入，嵌入/左/右/上下环绕；保留来源绘图属性；区域内图片保存媒体关系并能重开解码 | `editing/objects.ts`、`picture-files.ts`、`WordCoreObjects.tsx`；objects、paste-properties、header-footer、textbox |
| W24–W28 | 表格插入/整表移动、行列/矩形选择、增删行列/表、合并拆分、平均分布、适应页宽、重复表头、宽高/对齐/底纹/垂直对齐/边框和尺寸拖动 | 同源 `table-ops`、`table-properties`、`table-borders` 及 `editing/table-source.ts`；objects 核心/浏览器。保存保留表格/行/格未知 XML、段落来源和隐藏纵向续格 |
| W29 | 实际 XLSX/CSV 导入、字段插入、逐条预览；字段落到原生域，摘要读取当前正文/区域，删除/撤销不沿用旧来源字段 | `editing/mail-merge.ts`、`WordCoreMailMerge.tsx`；mail-merge、textbox 核心及 mail-merge 浏览器 |
| W30 | 多条件筛选/排序、记录选择、空值/文件命名选项及配置恢复；预览不改模板正文/历史，设置随模板保存重开 | `MailMergePanel.tsx`、`serialization/mail-merge-settings.ts`；mail-merge 核心/浏览器 |
| W31 | 新核心合并为可编辑 Word、单份/逐记录 DOCX ZIP、PDF/逐记录 PDF ZIP；支持简单/复杂域、嵌套 IF、NEXT/NEXTIF/SKIPIF，页眉脚注批注/书签/编号等按记录重映射 | `serialization/mail-merge.ts`、`merge-fields.ts`、`mail-merge-controller.ts`；mail-merge 核心/浏览器。失败/取消不改源模板，成品清理邮件连接 |
| W32–W33 | 缩放/适宽、当前页/总页数、字数/字符统计、布局/段落/属性面板及扫描实际部件的兼容报告；宏/签名/强保护在 handler、事务和保存层一致只读 | `editing/statistics.ts`、`core-capabilities.ts`、`WordCorePagePanel.tsx`；properties 核心、paste-properties、page-layout 浏览器 |
| W34 | 同源版式打印预览、翻页/缩放、预览拖动页边距、方向设置及平台打印入口；混合纸张和重复表头使用已提交的真实分页几何 | `core-output.ts`、`WordCoreOutput.tsx`、`platform-adapter/print.ts`；output 浏览器，实际打印媒体 PDF |
| W35 | PDF 页码范围校验、进度、取消和保存；输出固定捕获版本，取消丢弃部分结果并释放隐藏会话；邮件 PDF 复用同一路径 | `WordCoreOutput.tsx`、`export-service`；output 核心/浏览器及 mail-merge 实际下载 |

## 区域与保存的一致性

正文、页眉页脚及文本框仍是原生编辑器各自的 schema；命令从 `commandEditor()` 获取当前区域，能力参数限制不适用的结构操作。共享样式目录和页面状态通过原生 `DocAttrStep` 进入主历史，区域投影恢复选区与焦点，避免文档属性事务重建节点视图时失焦回写旧内容或抹掉重做。

页眉页脚布局表格格内承接完整段落格式，包括右缩进、底纹、双线边框、keepNext/keepLines/widowControl；保存时缩进和段距按改动的属性组精确合并，保留未修改的显式零值、字符单位及未知属性。文本框图片在捕获保存计划时重新登记媒体，不依赖已清空的全局暂存；重复保存和重开都验证了实际媒体字节和图片解码。

页面/标尺拖动只在松手提交。页码、总页数和 TOC 缓存绑定匹配文档的实际布局，旧测量不能污染新文档。输出使用只读快照和同源排版，源模板、保存确认与 dirty 状态独立。

## 最终验证

2026-11-11 最终 `pnpm verify`：**78 个测试文件通过、1 个跳过；756 项测试通过、3 项私人样本跳过**。类型检查、Web 构建、独立核心构建通过；独立构建 **425 模块 / 61 项输出**，固定来源核验 **335 项**。完整日志为 `output/playwright/word-core-m3.local/final-verify.log`。私人样本跳过不计为验收通过。

| 实际 Chromium 147 脚本 | 通过检查 | 产物目录（`output/playwright/` 下） |
| --- | ---: | --- |
| `verify-word-core-formatting.mjs` | 8 | `word-core-m3.local/` |
| `verify-word-core-search-lists.mjs` | 15 | `word-core-m3-search-lists.local/` |
| `verify-word-core-styles-painter.mjs` | 13 | `word-core-m3-styles-painter.local/` |
| `verify-word-core-page-layout.mjs` | 16 | `word-core-m3-page-layout.local/` |
| `verify-word-core-ruler-toc.mjs` | 14 | `word-core-m3-ruler-toc.local/` |
| `verify-word-core-objects.mjs` | 16 | `word-core-m3-objects.local/` |
| `verify-word-core-paste-properties.mjs` | 12 | `word-core-m3-paste-properties.local/` |
| `verify-word-core-mail-merge.mjs` | 8 | `word-core-m3-mail-merge.local/` |
| `verify-word-core-output.mjs` | 7 | `word-core-m3-output.local/` |
| `verify-word-core-header-footer.mjs` | 13 | `word-core-m3-header-footer.local/` |
| `verify-word-core-textbox.mjs` | 9 | `word-core-m3-textbox.local/` |
| `verify-word-core-surface.mjs`（M2 回归） | 16 | `word-core-m2.local/` |
| `verify-word-core-workspace.mjs`（M2 回归） | 8 | `word-core-m2-workspace.local/` |

本轮 M3 共 131 项真实浏览器检查，另有 M2 回归 24 项，各报告均为 passed 且无页面异常。最终序列化补齐后重跑了 header-footer、textbox、mail-merge 和 output；其余为本次 M3 工作中已完成的最终相应批次证据，并非声称每个脚本在最后一次代码修改后均重新执行。报告为各目录 `report.json`，截图、真实下载 DOCX/ZIP/PDF 均在同一忽略目录。

已实际打开检查正文/样式、窄窗口标尺、页眉区域编辑、文本框内文字/列表/图片、横向纸张及复杂表格第二页截图。浏览器打印入口以真实 print-media 页面和 Chromium PDF 验证；没有将其写成实际打印机出纸或 Tauri WebView 验收。

## 明确支持边界与后续阶段

- 目录承接原 Mini Office 既有的 1–3 级标题规则。完整独立目录可更新；多区域、嵌套、跨引用、损坏或夹带正文的导入目录拒绝覆盖并提示。没有新增旧产品不存在的任意高级目录编辑器。
- 单栏横排正文及可编辑区域开放标尺；混合制表位、列表/正文表格、镜像边距、多栏和竖排等复杂坐标情形受能力保护，不用错误坐标提交。原有布局和源 XML 仍按来源保留。
- 页眉页脚支持普通富文本和可编辑布局表格；浮动框、嵌套或锁定结构保留且禁止普通文本表单覆盖。区域 schema 不支持的整表结构/环绕等操作禁用。已有文本框可直接编辑文字、列表、样式和内联图片，本阶段不新增任意形状创建工具。
- 导入样式删除采用隐藏基底保留继承效果，不宣称彻底丢弃其他部件仍需的格式定义。未知字符/绘图/段落属性遵循局部改动保留规则，不能据合成语料推断所有私有业务文档已经完成 Word 视觉验收。
- 原生系统字体度量、外部位图剪贴板桥、拖入/双击打开、Tauri WebKit 输入和打印/文件平台组合，以及主题/窄窗口的完整生命周期仍在 M4/M5。浏览器中已完成的标签/恢复/连续保存是回归依据，不替代桌面验收。
- M5 才切换默认核心、进行安装包与实际 Word/私人样本验收和正式交付；本阶段没有发布行为。

以上是原功能和结构支持范围的保护，不再保留已完成的邮件合并、打印/PDF、区域编辑、样式/格式刷为“待迁移”。下一阶段为 M4 工作区及桌面平台融合。

## 复现

```bash
# 功能：启动开发核心并复现 M3 核心、浏览器及文件输出检查；时间：2026-11-11；作者：lq。
VITE_WORD_CORE_M2=true pnpm --filter @mini-office/desktop exec vite --port 1431 --strictPort
# 另一个终端执行；浏览器脚本支持 MINI_OFFICE_TEST_URL 和 MINI_OFFICE_CHROME。
pnpm verify
node scripts/verify-word-core-formatting.mjs
node scripts/verify-word-core-search-lists.mjs
node scripts/verify-word-core-styles-painter.mjs
node scripts/verify-word-core-page-layout.mjs
node scripts/verify-word-core-ruler-toc.mjs
node scripts/verify-word-core-objects.mjs
node scripts/verify-word-core-paste-properties.mjs
node scripts/verify-word-core-mail-merge.mjs
node scripts/verify-word-core-output.mjs
node scripts/verify-word-core-header-footer.mjs
node scripts/verify-word-core-textbox.mjs
node scripts/verify-word-core-surface.mjs
node scripts/verify-word-core-workspace.mjs
```

本环境 Playwright CLI 包此前无法从镜像下载，浏览器验收沿用仓库已安装的 `playwright-core` 和既有脚本。固定来源的修改仅由 `scripts/extract-word-core.mjs` 重放，并更新 `SOURCE-MANIFEST.json`；正常构建不读取忽略目录里的上游副本。
