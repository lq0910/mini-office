# Word 版式保真方案（Layout Fidelity）

- 状态：历史方案；当前执行路线以 2026-11-03 更新为准
- 日期：2026-11-02
- 关联：ADR 0001（编辑引擎选型）、ADR 0004（PDF 策略）、`docs/兼容性增强轻量化路线图.md`

> 2026-11-03 更新，作者：lq：以下诊断为历史快照，部分结论已被后续评审纠正，不能作为当前代码事实。最新实施范围见 [实施记录](word-fidelity-implementation-20261103.md)，下一阶段以 [开发路线图](兼容性增强轻量化路线图.md) 的 GenOffice 样本驱动计划为准。

## 0. 结论

1. **`.doc` 打开后"格式全没了、页码也错"不是解析精度问题，而是组件选型错误。** 当前接入的 `office_oxide` 定位是"文本抽取库"，其 `convert_doc.rs` 不产出 `sectPr`、纸张尺寸、页边距、页眉页脚和域，转换结果必然只剩段落和表格骨架。
2. **分页不准的根因不在"测量"，而在"规则"。** 现有实现只做"实测高度 → 装箱"，缺少 Word 的 `keepNext`、`widowControl`、`keepLines`、`cantSplit`、`docGrid`、分节符类型等分页规则；同时字体被替换导致行宽与行高与 Word 不同。
3. **"完美还原"只能由同一排版引擎保证。** 同一份中文 DOCX 在 Word、WPS、LibreOffice 下版面本就不一致，独立实现只能"逼近"。因此方案目标改为**可验收的高保真**（页数一致率 ≥ 95%、分页点一致率 ≥ 90%），并用**可选外部引擎**兜住旧格式与复杂版式。默认安装包不增加任何大型运行时。

## 1. 现状诊断（附代码证据）

| 现象 | 根因 | 证据 |
|---|---|---|
| `.doc` 打开后版式、页码、页脚全部丢失 | `office_oxide` 是文本抽取库；`convert_doc.rs` 中不存在 `sectPr` / `pgSz` / `pgMar` / `headerReference` / `footerReference` / `instrText`，转换产物是重建的骨架 DOCX | `crates/office-compat/src/lib.rs:47`、`office_oxide-0.1.15/src/convert_doc.rs` |
| 打开 `.doc` 会在同目录生成 `xxx_兼容转换.docx` | 转换命令默认写入源文件同目录 | `apps/desktop/src-tauri/src/lib.rs:217-226` |
| 页脚"第 X 页 共 Y 页"出现在正文流里、位置错误 | 域被展平成普通文本，且 `PAGE` / `NUMPAGES` 在分页前一次性求值 | `packages/editor-word/src/parse.ts` 域解析 + `pagination.ts` 单遍装箱 |
| 分页点与 Word 不同、页数偏差 | 分页只按实测高度装箱，未实现 Word 分页规则 | `packages/editor-word/src/pagination.ts`；`keepNext` 仅被解析（`parse.ts:286`）未参与分页；`docGrid` 仅在序列化时透传（`serialize.ts:597`） |
| 表格跨页、行高、列宽偏差 | 未实现 `tblLayout=autofit` 的 Word 列宽分配算法与 `tblCellMar` | `parse.ts:1218-1330` 仅按 `tblGrid` / `tcW` 取值 |
| 只滚动未编辑也提示"已修改" | `updateData` 默认 `dirty = true`，任何非用户触发的数据刷新都会置脏 | `packages/document-model/src/store.ts:725` |

补充事实：中文 DOCX 的版面还原度，业界实测顺序是 **Word（基准）> WPS > LibreOffice ≈ ONLYOFFICE**。因此 LibreOffice 适合做"转换器/完整性兜底"，不能当作"保真度上限"来对标。

## 2. 边界定义：什么叫"还原"

用可验收的等级替代"完美"：

| 等级 | 含义 | 适用范围 |
|---|---|---|
| F0 | 文本、段落、表格结构正确 | 所有文档，必须 100% |
| F1 | 字体、字号、行距、缩进、对齐、边框底纹一致 | 所有文档，目标 ≥ 95% |
| F2 | 分页点、页数、页眉页脚、页码一致 | 合同/公文/通知/报告，目标 ≥ 95% |
| F3 | 像素级一致（浮动对象、分栏、文本框环绕、图文混排） | 复杂排版，**不承诺**，保留原件并提示 |

验收指标（对语料集统计，不靠主观判断）：

- 页数一致率
- 分页点一致率（同一段落在同一页的比例）
- 行盒位置中位偏差（px @96dpi，目标 ≤ 2px）
- 表格列宽相对偏差（目标 ≤ 1%）
- 页眉页脚基线偏差（目标 ≤ 2px）
- 页码正确率（`PAGE` / `NUMPAGES` / `SECTIONPAGES`）

## 3. 方案总览：双通道 + 共享内核

```
                    ┌──────────────────────────────┐
   DOCX/XLSX/PPTX → │  OPC 包 + 未知部件保留（现有）│
                    └───────────────┬──────────────┘
                                    │
                    ┌───────────────▼──────────────┐
                    │ 文档模型 IR + 度量层（新增） │  ← 字体度量、单位、样式链
                    └───────┬──────────────┬───────┘
                            │              │
              ┌─────────────▼───┐   ┌──────▼─────────────────┐
              │ 编辑通道（默认）│   │ 保真通道（按需）       │
              │ DOM 布局 + 分页 │   │ 外部引擎 → PDF/预览     │
              │ 可编辑、可回写  │   │ 只读、用于对照/打印     │
              └─────────────────┘   └────────────────────────┘
```

- **编辑通道**仍是自有内核（ADR 0001 决策不变），本方案只补"度量 + 分页规则"两块短板。
- **保真通道**不替换编辑器：用于 `.doc/.xls/.ppt` 结构转换、复杂版式的只读对照、以及"打印/导出必须与 Office 一致"的场景。默认不随包分发。

## 4. 关键技术设计

### 4.1 度量层：先把"尺子"校准

这是投入产出比最高的一项，也是当前偏差的最大来源。

1. **字体度量来源统一**：用 Rust `ttf-parser`（已依赖）读取 `head/OS2/hhea`，按 Word 公式计算行高：
   `行高 = ascent + descent + lineGap`（`useTypoMetrics` 时取 `sTypoAscender/Descender/LineGap`）。
2. **度量兼容字体**：西文优先内置 Carlito（Calibri）、Caladea（Cambria）、Liberation（Arial/Times/Courier），均为 SIL OFL，仅打包 Latin 子集（约 1.5–2 MB）。仓库已有 `@fontsource/caladea`、`@fontsource/carlito`，可直接复用。
3. **中文字体映射表**：宋体 → `Songti SC` / `SimSun`，黑体 → `Heiti SC` / `SimHei`，微软雅黑 → `Microsoft YaHei` / `PingFang SC`；映射不到时按"字面宽度最接近"降级，并在兼容性报告中列出缺失字体。
4. **禁止依赖 `line-height: normal`**：所有段落行高由度量层算出具体像素值，避免 WebView 差异（WKWebView / WebView2 / WebKitGTK 默认行高不同）。

### 4.2 段落盒模型

- 段前/段后折叠：页首不叠加 `spaceBefore`，页尾不叠加 `spaceAfter`，相邻同段样式按 `contextualSpacing` 抑制。
- 行距三规则精确换算：`auto`（240 分之一行）、`exact`、`atLeast`。
- `docGrid`（中文公文核心）：按 `linePitch` 计算每页行数，行高取 `contentHeight / lines`，实现"22 行 × 28 字"的公文档面；`snapToGrid` 控制是否吸附。
- 中英文混排：启用 `line-break: strict`，标点挤压按 Word 的"行末标点压缩"近似处理。

### 4.3 分页内核（从"单遍装箱"改为"两遍布局"）

第一遍：按块高度做首次适配；第二遍：回溯修正分页点并求解域。

必须补齐的规则：

1. `keepNext` 链：标题与后续段落绑定，整链放不下则整体下移（不是逐块判断）。
2. `widowControl`：段首/段尾孤行控制（默认开启）。
3. `keepLines`：段内不拆页。
4. 表格 `cantSplit`：行不可跨页；表头 `tblHeader` 重复。
5. 分节符类型：`nextPage` / `continuous` / `evenPage` / `oddPage`，含节内页码重启（`w:pgNumType`）。
6. 页眉页脚区：`w:header` / `w:footer` 距边界值参与正文可用高度计算，首页不同、奇偶页不同。
7. 域求值：`PAGE` / `NUMPAGES` / `SECTIONPAGES` / `TOC` 在分页完成后统一求值，避免"页码错"与"目录页码不更新"。

### 4.4 表格

- 实现 `tblLayout=autofit` 的列宽分配（内容优先、按 `tblW` 百分比收敛），`fixed` 时严格用 `tblGrid`。
- `tblCellMar` 单元格内边距、`vAlign` 垂直对齐。
- 边框优先级：单元格 > 表格 > 表格样式，缺省继承 `styles.xml` 的 `tblBorders`。
- 行高 `trHeight` 的 `atLeast` / `exact` 语义。

### 4.5 浮动对象与文本框

先做**只读保真**，不做编辑：`wp:anchor` 的相对定位（段落/页边距/页面/行）、`wrapSquare/Tight/Through/TopAndBottom`、`z-order`、`txbxContent` 文本渲染。渲染不了的复杂对象保留原部件，并在兼容性报告中标红。

### 4.6 保存保真：段落级脏区回写

现有策略是"复制未修改部件 + 重写主文档"，主文档仍是整篇重建。升级为 **byte-level patch**（GenOffice 路线）：

- 记录每个段落/表格的原始 XML 片段与哈希；
- 保存时只替换脏段落的片段，其余字节原样保留；
- 未被编辑器识别的节点（`mc:AlternateContent`、`w:sdt`、批注锚点）不参与重建。

收益：往返保存后 `document.xml` 的 diff 收敛到被编辑的段落，避免"打开再保存就变形"。

### 4.7 旧格式通道（`.doc/.xls/.ppt`）

三级降级，**永不写 sidecar 文件**：

| 优先级 | 方案 | 体积 | 保真度 | 条件 |
|---|---|---|---|---|
| 1 | 调用本机已安装的 Word / WPS 转换 | 0 | 最高 | 用户机器已装且授权允许自动化 |
| 2 | 内置可选 LibreOffice 组件（按需下载、独立进程） | 安装包 0，首次下载约 110–160 MB | 高（优于任何纯 Rust 解析） | 用户显式开启"高保真打开" |
| 3 | `office_oxide` 文本兜底 | 已有 | 低（仅文本/表格骨架） | 前两级不可用 |

产品语义：

- `.doc` 打开为**只读兼容会话**，标题栏标注"兼容模式（DOC）"；
- 保存默认"另存为 .docx"，**不覆盖原 `.doc`**；
- 转换在内存或应用缓存目录完成，源目录不新增文件；
- 转换失败给出明确原因，不静默降级为"看起来能打开但版式错"。

## 5. 体积与性能预算

| 项目 | 安装包增量 | 运行时内存 |
|---|---|---|
| 度量层（复用 `ttf-parser`） | 0 | 可忽略 |
| 度量兼容西文字体子集 | +1.5–2.5 MB | 按需加载 |
| 分页/域两遍求值 | 0 | 单文档 < 5 MB |
| PDF 对照视图（pdf.js） | +0.4 MB（gzip） | 按页渲染 |
| LibreOffice 可选组件 | 0（按需下载，不计入安装包） | 独立进程，退出即释放 |
| **默认安装包合计** | **≤ 3 MB** | — |

性能目标：10 页 DOCX 首屏 ≤ 1.2 s（含两遍布局），分页重算 ≤ 80 ms/次（仅脏块重算）。

## 6. 实施阶段

### P0 止血（已完成，2026-11-02）

1. `store.ts`：`updateData` / `updateDocMeta` 的 `dirty` 改为**显式声明**，默认 `false`；Word、Excel 的 `commit` / `undo` / `redo` 显式传 `{ dirty: true }`。滚动、重排、统计刷新不再置脏。
2. `.doc` 停止写 sidecar：`convert_legacy_document` 改为**内存转换**并返回 `{ name, sourceName, bytes }`，源目录零写入。
3. 兼容会话：标签显示原始 `.doc` 文件名，不绑定源路径，`canSaveInPlace=false`、`saveAsCopyRequired=true`，保存只能另存为 OOXML；打开时一次性提示"兼容模式"。
4. 回归测试：新增"内部刷新不置脏"用例，保存生命周期 9 项、全量 510 项通过。

**尚未解决（属 P1/P3）**：`.doc` 的版式保真本身——`office_oxide` 仍是文本通道，页边距、页眉页脚、域无法还原，需走 4.7 的保真通道。

### P1 度量与分页内核

**P1.1 已完成（2026-11-02）**

1. **度量层**：新增 `packages/editor-word/src/metrics.ts`。
   - `singleLineHeightPx()` 用 Canvas `fontBoundingBoxAscent/Descent` 取字体真实单行高度，按「字体栈 + 字号」缓存；无 Canvas 环境按中英文系数兜底。
   - `lineHeightPx()` 按 Word 语义换算：`exact` 直取、`atLeast` 取大、`auto` 为「单行高度 × 倍数」。**修正了此前把「1.5 倍行距」当成「字号 × 1.5」的系统性偏小**。
   - `paragraphFontSizePt()` 取段内最大字号，与 Word「按行内最大字号定行高」一致。
   - 字体加载完成后清空缓存并触发重新分页。
2. **文档网格**：`PageSetup.docGrid` 新增，解析 `w:docGrid/@linePitch`（`type` 缺省或 `default` 归一为 `lines`）；启用网格时单倍行高取网格行距，公文类文档每页行数得以对齐；`sectPr` 缺失时保存会补写。
3. **段落盒模型**：段前距在页首丢弃（`paragraphWrapperStyle(..., { dropSpaceBefore })`），分页高度同步扣减，消除每个新页顶部多出的空白。
4. **分页规则**：`keepNext` 链整体放不下时先换页，标题不再孤零零落在页尾。

验证：新增 `tests/word-layout-metrics.test.ts`（9 项）；合同样本断言网格 `linePitch=360 → 24px`；全量 519 项通过。

**P1.2 已完成（2026-11-03）**

1. **段落行级切分**：`pagination.ts` 新增 `ParagraphLayout` / `ParagraphSplit`。段落不再原子移动——放不下时按行盒切成「首片（可编辑）+ 续页片（只读）」，页底不再留半页空白。实现方式是**整段渲染 + 窗口裁剪**（固定高度 + 负 `margin-top`），而不是真正切开文本，因为切开会让 `onInput` 只回读到半段内容。
2. **行盒测量**：`metrics.measureParagraphLines()` 以「实测高度 ÷ 计算行高」推定行数；两者不能整除（行内有更高内联内容，如图片）时返回 `undefined`，分页自动回退到整段移动——宁可少切一刀，不把行切错。
3. **孤行/寡行控制**：`widowControl`（默认开启）保证页尾与续页开头都不出现单行；整段能放进一整页时优先整体移页。
4. **`keepLines`**：整段不拆页，用于条款、签名栏；标题预设同步补上 `keepLines`，与 Word 内置标题样式一致。
5. **字符缩进**：解析 `w:firstLineChars` / `w:hangingChars` / `w:leftChars`（单位 1/100 字符），并按段落字号换算。此前只读 twips 字段，WPS 常见的「首行缩进 2 字符」会整段丢失缩进。

验证：`tests/word-layout-metrics.test.ts` 新增 7 项（切分、段前距归属、孤行、寡行、keepLines、关闭孤行、无行盒回退），`word-docx-line-break` 新增 2 项（字符缩进、分页属性解析），`ooxml-element-order` 覆盖 `keepLines`/`widowControl` 的元素顺序；全量 528 项通过。

**已知限制**：跨页段落的首片虽可编辑，但 DOM 中仍含整段内容，光标若移入被裁剪区域会不可见。这与 JitWord 复盘中「宁可先不可编辑，也绝对不能丢内容」的取舍一致，后续可做「聚焦时展开裁剪窗口」。

**P1.3 待做**

- **字符网格**：`docGrid type="linesAndChars"` 的每行字数约束（当前只用了 `linePitch` 的行数约束）。
- **分栏**：`w:cols` 目前只解析不排版。

### P2 表格与域（进行中）

**P2.1 已完成（2026-11-03）**

1. **`tblW` 百分比宽度**：`w:tblW w:type="pct"` 此前只解析、不参与渲染，表格一律按 `tblGrid` 的绝对宽度显示。Word 新建表格默认就是 `pct 5000`（满宽），因此这类表格会明显偏窄。现在改为按百分比缩放列宽轨道，表格宽度取 `widthPct%`。
2. **单元格默认边距**：由 `3px 6px` 改为 `0 5.4px`，对齐 Word 默认的 108 twips（≈5.4px）左右边距、上下 0 边距。此前每行多出 6px。
3. **行高下限**：`minmax(28px, auto)` 改为 `minmax(row.heightPx ?? 20px, auto)`。Word 的行高由内容决定，固定 28px 让每行比 Word 高约 8px，表格一长就会顶到下一页。

三项合计，一份 10 行的表格比修正前矮约 140px，正好是半页——这是请假单这类表单文档页数偏差的直接来源。

**P2.2 待做**

- `w:tblCellMar` 自定义单元格边距、`w:tblInd` 表格左缩进。
- `tblLayout type="autofit"` 按内容自适应列宽。
- 目录/页码域联动（`PAGE`/`NUMPAGES` 已在分页后求值，`TOC` 条目仍只支持标题）。

### P3 保真通道（约 1.5 周）

- 4.7 三级旧格式通道 + PDF 对照视图（可选组件管理、下载校验、失败回退）。
- 验收：`.doc` 4 页合同在保真通道下页数与 Word 一致；无组件时明确提示而非错误渲染。

### P4 保存保真与回归体系（持续）

- 4.6 段落级脏区回写；建立 50 份语料 + 视觉回归 CI。

## 7. 度量体系：把"像不像"变成数字

1. **语料集**：50 份真实文档（合同、公文、通知、报告、请假单、报价表），覆盖单节/多节、合并单元格、页眉页脚、图片、文本框。
2. **基准图**：用 Microsoft Word（首选）与 WPS 各导出一次 PDF，离线生成后入库，CI 不需要安装 Office。
3. **比对方式**：
   - 页数、分页点（按段落 id 对齐）；
   - 页面栅格化后逐页 SSIM；
   - 文本行盒位置（`getBoundingClientRect` 归一化到页面坐标）；
   - 表格列边界 x 坐标序列。
4. **门禁**：页数一致率、分页点一致率、SSIM 三项作为 PR 检查项；任一低于阈值即阻断合并。
5. **兼容性报告**：在现有 L0–L3 能力等级旁增加 **F 级保真评分**，打开时即告知用户"本文档预计还原度 92%，含 2 个不支持对象"。

## 8. 本轮立即修复清单（可独立提交）

| # | 问题 | 位置 | 修法 |
|---|---|---|---|
| 1 | 只滚动也提示已修改 | `packages/document-model/src/store.ts:725` | `dirty` 默认 `false`，编辑动作显式置 `true` |
| 2 | 打开 `.doc` 生成 `_兼容转换.docx` | `apps/desktop/src-tauri/src/lib.rs:217` | 内存转换 + 只读会话，源目录零写入 |
| 3 | `.doc` 版式丢失 | `crates/office-compat` | 停止把 `office_oxide` 当版式转换器，改走 4.7 通道 |
| 4 | 页码/页数错误 | `pagination.ts` + `html.ts` | 已核实：`PAGE`/`NUMPAGES` 在渲染阶段按 `pages.length` 求值，逻辑正确；真实原因是旧格式转换把域展平成文本，属 P3 |
| 5 | Excel 合并单元格文字不显示、无两端对齐 | `packages/editor-sheet` | 待复现：模型层 `HorizontalAlign` 只有 general/left/center/right，两端对齐需先扩枚举再补工具栏，单独排期 |

## 9. 对既有决策的影响

- **ADR 0001 不变**：编辑内核仍自建，不引入 ONLYOFFICE / Univer 作为编辑器。
- **ADR 0004 延续**：其"后续"已写明"将 LibreOffice 作为可选高保真对照/转换插件，不在默认安装包内"，本方案把它落为 P3。
- **新增建议**：补一条 ADR 0006「保真通道与外部引擎使用边界」，明确三点——独立进程、只读产物、不参与编辑与保存链路。

## 10. 明确不做

- 不追求 F3 像素级（分栏、SmartArt、图文环绕编辑）。
- 不内置 LibreOffice 到安装包。
- 不用 AGPL 组件（ONLYOFFICE sdkjs、SuperDoc）进入闭源分发链路。
- 不为旧格式实现原格式写回。
