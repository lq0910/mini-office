# 两份真实 DOC 的 Word 保真进度

功能：记录原件、真实 Word 基准、本轮修复和未达到验收的部分；时间：2026-11-05；作者：lq。

当前为 `main` 的本地工作区改动，未提交、推送或重新打包。私人文件、Word 导出件和截图均留在已忽略的 `output/word-user-fidelity.local/`，不加入测试仓库。

## 基准与结论

用户指定的 `/Users/liqiang/Downloads/采购文件.doc` 和 `/Users/liqiang/Downloads/测试分包合同.doc` 全程只读。通过副本在本机 Microsoft Word 16.114 中导出 PDF，再以 Poppler 96 DPI 渲染；Word 导出的 DOCX 仅用于分析属性，不作为产品导入捷径。

| 样本 | Word 原件 PDF | 修复前 Mini Office | 修复后 Mini Office |
| --- | ---: | ---: | ---: |
| 采购文件.doc | 51 页 | 50 页 | 50 页 |
| 测试分包合同.doc | 5 页 | 4 页 | 5 页 |

两份文件已全部捕获浏览器页面，共 55 页；Word 基准共 56 页。`output/word-user-fidelity.local/index.html` 提供原尺寸逐页对照，`comparison.json` 记录源文件/PDF 哈希、页数和逐页像素误差。采购文件最后一页明确列为缺页，不以交集页数冒充验收通过。

合同的前两页实际可见页尾分界与 Word 对上；检查了后续代表页，仍有替代字形、下划线空格宽度、签章区横向位置和文字基线差异。因此合同也没有达到像素一致验收。

首页原尺寸像素比较：合同的不同像素比例由 9.517% 降为 7.215%，平均通道误差由 11.079 降为 8.392（范围 0–255）；采购文件首页由 4.395% 变为 4.405%，没有改善。数值包含字体替代和栅格边缘，不能换算成还原率。

## 已修复

1. 段落逐 run 使用实际字体、字号和备用字体度量，停止把所有段落的字号套到默认 Calibri 上；显式小字号也不再被文档默认字号抬高。五个常用 Office 中文字体族使用固定上游的 Word 设计行高，未知字体继续实测。
2. 保留 DOC FFN 的显式备用名并生成完整 OOXML 字体表；解析 DOCX 的 `w:altName`。备用规则绑定原字体，工具栏切换到其他字体后不继续沿用旧规则。
3. 对安装清单确认缺失的旧 GB2312 字体，使用本机 Word 已验证的微软雅黑显示备用链；原字体名保留，确实安装原字体或安装情况未知时不推断替代。此行为以本次 macOS Word 环境为验证目标，未宣称覆盖所有 Office 平台。
4. 读取并回写 `sprmPFWidowControl (0x2431)`。合同旧模板在样式中关闭了孤行控制，原转换漏掉 false，造成页尾整段提前移走。修复后可以在第一页末尾显示“合计金额”段落的首行，再在第二页继续。
5. 读取并回写 `sprmPFUsePgsuSettings (0x2447)`，恢复标题等段落的 `snapToGrid=false`；生成器也加入对应三态标志处理。另补 NFC 37 的中文编号映射，恢复“第一章”。操作码依据 [Microsoft DOC 二进制规范](https://download.microsoft.com/download/0/b/e/0be8bdd7-e5e8-422a-abfd-4342ed7ad886/word97-2007binaryfileformat%28doc%29specification.pdf)。
6. 字符缩进改用段落标记的有效字号，修复采购模板 16pt 的三字缩进按默认 10pt 缩小的问题。
7. 新增私人文件验收脚本 `scripts/verify-word-user-fidelity.mjs`。报告同时保留原 DOM 文字和按裁剪窗口提取的可见文字，避免把跨页复制的隐藏段落误报为重复正文。

校正：合同原始 DOC 的 FFN 没有给“仿宋_GB2312”声明“微软雅黑”备用名；该备用名是 Word 导出 DOCX 时补入的。本轮保留源 FFN 备用名与缺失字体的兼容替代属于两条独立逻辑，不把导出件的信息冒充原始 DOC 声明。

## 仍需解决

- 采购文件为 50 页，原 Word 为 51 页；第三页开始就有页尾分界差异。不能用强制补页或整体偏移修正。
- 采购页脚文本框和水平锚定本轮已恢复，细节见下节；复杂文本框的全部绘制效果和正文内连续编辑仍不在已完成范围。
- 上一阶段 Word 打开转换包时采购为 51 页；补齐网格关闭值后，最新转换包在 Word 中为 50 页，与浏览器一致，但原 DOC 为 51 页。差异已包含 DOC 转换阶段，不能只调整浏览器分页，更不能仅凭转换件的页数证明转换无损。
- 字体替代仍改变空格和全角下划线字形，采购表格宽度、列内换行、标题位置和签章区还需逐页校准。没有复制或分发本机 Word 的商业字体文件。
- 采购转换仍提示 TOC/PAGEREF 缓存文字导入及一个无尺寸浮动对象省略。完整文本框、分节页眉页脚和连续跨页编辑不在本轮已完成范围内。

## 验证

- `MINI_OFFICE_FIDELITY_SAMPLES=/Users/liqiang/Downloads pnpm test`：603 项通过，包括三份私人样本，未跳过；验证无操作保存返回转换后的原始 OOXML 字节，普通文字修改后其他 ZIP 部件保持一致并能重新解析。这不等于 DOC 二进制无损往返，也不等于修改后文件的 Word 像素验收。
- `pnpm typecheck`、`pnpm build`、脚本语法和 `git diff --check` 通过。原有 React act 与 Vite 动态导入提示仍存在。
- 公开 8 份语料、9 页重新捕获，执行失败 0；仍有像素差异，严格命令按设计返回 1。
- 两份 Downloads 原件前后 SHA-256 一致。未验证 Tauri WebKit/WebView2 实机或修改后文件在 Word 中的逐页效果。

复验入口：先启动本机开发服务，再运行 `MINI_OFFICE_TEST_URL=http://127.0.0.1:4196/ node scripts/verify-word-user-fidelity.mjs /Users/liqiang/Downloads 采购文件.doc 测试分包合同.doc`。可用 `MINI_OFFICE_USER_FIDELITY_OUTPUT` 指定其他 `.local` 目录，避免覆盖本轮对照。

## 页码文本框与表格单位修复（2026-11-05，后续轮次）

功能：恢复采购正文页码，修复 DOC 文本框锚定与不折行属性，并校正表格的 twips/pt/px 换算；时间：2026-11-05；作者：lq。

1. DOC OfficeArt 读取 `posh (0x038F)`，居中等对齐输出为 `wp:align`，不再用零偏移矩形冒充左对齐。读取 `WrapText (0x0085)`，源 `msowrapNone=2` 输出为 `wps:bodyPr wrap="none"`。前者单独修复仍不足：Word 会按旧的 144pt 占位框居中，页码字符仍偏左；补回不折行属性后，自动适应文本框才按字形宽度收缩。规范依据：[posh](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw/4fa5d949-3bc9-49b0-aa3c-ad6286ef01bc)、[WrapText](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw/749ffcd2-b8f2-4311-9f9f-bec3231bdc73)。
2. 页眉页脚模型单独保留只读 DrawingML 文本框及源 XML，解析框内段落和 PAGE 域，按页面上下文渲染。自动适应且不折行的框按内容收缩；可折行框保留声明宽度。普通页眉页脚文字保存时仍写回原文本框，复杂内容在原有编辑对话框中受保护，避免普通文字表单无提示覆盖。
3. 实际浏览器捕获 55 页：采购 50、合同 5。采购前两页没有页码，正文第 3–50 页共 48 个框全部显示对应页码；框中心与页面中心最大偏差为 0 CSS px。截图像素检查确认页码位置存在实际黑色字形。
4. 将修复后的采购转换包在 Word 中打开并导出 PDF，第三页页码坐标为 x=393.92px、y=1044.23px，与原 DOC 导出的同一页码坐标一致。只补居中而未补不折行时 x=300.80px，此错误中间结果未作为最终验收。
5. 修复表格默认左右内边距：108 twips 是 5.4pt，即 7.2 CSS px，旧显示代码误用了 5.4px。左对齐表格按首格文字起点对齐版心，边框需扣除首格左内边距。采购首表左边框实测 x=75.9375px，基准 PNG 约 76px；数量列“1项”恢复为两行，左右内边距均为 7.2px。没有改变源列宽和保存模型的尺寸。

完整前端测试 605 项通过，包含全部私人样本；类型检查、生产构建及公开九页浏览器捕获通过执行检查，严格像素验收仍未通过。新增两项合成样本测试覆盖文本框域解析、无操作字节保持，以及修改普通页脚文字后的文本框保存回读。

最新私人产物位于 `output/word-user-fidelity.local/after-layout-final/`，入口 `index.html` 和 `comparison.json` 已刷新。采购仍为 50 页，合同仍为 5 页；未声明采购保真完成。两份原件 SHA-256 再次核对一致，未提交、推送或重新打包，未验证 Tauri WebView 实机。
