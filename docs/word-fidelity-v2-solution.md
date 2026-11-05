# Word 解析保真度根治方案（V2）

- 状态：待评审
- 日期：2026-11-03
- 关联：`docs/word-layout-fidelity-plan.md`（V1 方案）、ADR 0001、ADR 0004
- 定位：V1 已解决「脏标记误报、sidecar 污染、行高语义、网格行距、表格 pct 宽度」等问题；本文针对**剩余的、导致「与原版 Word 差异过大」的结构性差距**，给出根治路径。

---

## 0. 结论（先给判断）

**当前问题不是「精度不够」，而是三条互不相关的通道各自的缺陷叠加**。必须分而治之：

| 通道 | 用户可见现象 | 根因性质 | 可否自研解决 |
|---|---|---|---|
| **A. `.doc` 旧格式** | 采购文件.doc：字体全丢、表格全丢、纸张变 Letter、目录错位 | **数据源本身缺失**（转换器不产出这些信息） | ❌ 不可能自研补回 |
| **B. `.docx` 表格** | 请假单：表格比 Word 宽、行高比例不对 | **规则未实现**（autofit、hRule 缺失） | ✅ 可以 |
| **C. `.docx` 排版细节** | 合同：字号字重不对、行距、缩进 | **样式链未完全生效**（部分继承、部分丢失） | ✅ 可以 |

**关键判断**：截图里 8 张图的差异，**大约 60% 来自通道 A**（`.doc` 转换丢数据），**40% 来自 B/C**（`.docx` 规则缺失）。这决定了投入方向：

- 对 `.docx`（主流格式）：**继续自研，完全可达 F2 级保真**，本方案给出全部缺口与修法。
- 对 `.doc`（旧格式）：**自研无法补回未转换的信息**，唯一根治路径是换用保真转换器（LibreOffice / 本机 Word / WPS），这是**产品级决策**，不是技术优化。

---

## 1. 实证诊断（附可复现证据）

所有结论均通过实际运行代码 + 解包产物 XML 得到，非推测。

### 1.1 通道 A：`.doc` → `.docx` 转换实测

执行真实转换（`office_compat::convert_legacy`）后解包 `word/document.xml` 统计：

```
测试分包合同.doc  →  converted.docx
  w:sectPr              4     ✅ 有节
  w:pgSz                0     ❌ 页面尺寸全丢
  w:pgMar               0     ❌ 页边距全丢
  w:headerReference     0     ❌ 页眉全丢
  w:footerReference     2     ✅ 页脚保留
  w:rFonts              0     ❌ 所有字体族全丢
  w:tbl                 0     ❌ 表格结构全丢
  w:tblGrid             0     ❌
  w:tblW                0     ❌
  w:numPr               0     ❌ 编号列表全丢
  w:ind                23     ✅ 缩进保留
  w:spacing            52     ✅ 行距保留
  w:b/                  7     ✅ 加粗保留
  w:sz                 30     ✅ 字号保留
```

转换后 `sectPr` 实际内容（可见页面设置完全为空）：

```xml
<w:sectPr><w:footerReference w:type="default" r:id="rId2"/><w:type w:val="continuous"/></w:sectPr>
```

`styles.xml` 的默认字体被写死为：

```xml
<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>
```

**源码级根因**（`office_oxide-0.1.15`）：

| 检查项 | 结果 |
|---|---|
| 全 crate 搜索 `sectPr` / `pgSz` / `pgMar` | **0 命中** |
| `doc/` 解析器（`chpx.rs`/`papx.rs`/`document.rs`）搜索 `section`/`margin`/`page_width` | **0 命中** |
| `doc/document.rs` 公开字段 | 仅 `text` / `odd_header` / `odd_footer` / `even_header` / `even_footer` / `first_header` / `first_footer` / `author` / `kind` |
| `doc/chpx.rs` 是否提取字体名（sprm `0x4A`/`0x4B`） | **0 命中** |
| `convert_doc.rs` 的 `styled_span()` 映射字段 | bold/italic/underline/color/size/strike/highlight/caps — **无 font_name** |

结论：`office_oxide` 的 DOC 解析器**只抽文本与页眉页脚**，字体、纸张、边距、表格、编号一律不产出。`TextSpan` 虽然有 `font_name` 字段，但 DOC 路径从未填充。

> 这直接解释了「采购文件.doc」截图：标题从「方正小标宋简体 + 二号」变成默认栈，纸张从 A4 变默认，51 页的表格与目录结构丢失。

### 1.2 通道 B：`.docx` 表格实测

用本项目解析器解析 `测试请假单.docx`（原始 XML **完整包含**表格定义）：

```
原始 XML：           解析器输出：
w:tblGrid  4 列      列宽 184,106,85,228  (合计 603px)
w:tblW  auto         widthPct = undefined  ❌
w:trHeight ×5        行高 75.6 / 75.6 / 68.1 / 75.6 / 75.6  ✅
w:vAlign ×12         va=middle / top  ✅
w:gridSpan ×3        3x1  ✅
w:tblCellMar         未解析  ❌
```

**致命偏差**：原始 `tblGrid` 之和 = 2755+1596+1276+3421 = **9048 twips**，而正文可用宽度 = 11906−1800−1800 = **8306 twips**。

Word 的行为：`tblW type="auto"` → 自动调整表格以适应正文宽度，**按比例压缩到 8306**。
我们的行为：直接用 `tblGrid` 原值 → **604px vs 554px，表格超出正文 50px**。

这与「请假单」截图完全吻合：Mini Office 的表格右边界明显越过 Word 的表格右边界。

同时，`w:trHeight w:val="1134"` **无 `hRule` 属性**（Word 语义 = `atLeast`），我们按固定值渲染，行高与 Word 不一致。

### 1.3 通道 C：`.docx` 排版细节

对 `测试请假单.docx` 的解析实测输出：

```
[0] jc=center  font=微软雅黑 sz=16  | 专家请假单
[1] jc=justify font=仿宋_GB2312 sz=12 | 注：单位意见需单位负责人签名。
```

- 段落对齐（`jc`）、字体、字号、对齐 ✅ 正确解析
- 但 `doc.defaultFont` **为空字符串** → 段落级字体回退依赖 `run` 级，缺 `rFonts` 的段落会掉到全局默认栈
- `indentLeftPx` / `firstLineIndentPx` 在这两段均为 `undefined`（原文确实无缩进，合理）

**因此「合同」截图的差异不在解析，而在渲染与排版规则**：

| 现象（截图） | 根因 | 证据 |
|---|---|---|
| 标题「电网建设工程专业分包合同示范文本」字重/字号不对 | 默认字体栈与文档字体不一致；`DOC_FONT_STACK` 以 `PingFang SC`（黑体类）打头 | `design-system/src/index.ts:15` |
| 「第一部分 协议书」与 Word 间距不同 | `keepNext`/`keepLines` 链未完整参与行级切分 | `pagination.ts` |
| 正文行距比 Word 松 | 首行缩进/行高在 `docGrid` 生效时二者口径不同 | `metrics.ts:100` |
| 「采购文件」51 页 vs Mini Office 47 页 | `.doc` 丢表格后高度塌缩 → 分页必然不一致 | 见 1.1 |

---

## 2. 差距总表（可勾选验收）

### 2.1 通道 A：`.doc` 保真（**需产品决策**）

| # | 缺口 | 现状 | 方案 |
|---|---|---|---|
| A1 | 页面尺寸 / 页边距丢失 | `pgSz=0, pgMar=0` | 换转换器 |
| A2 | 字体族丢失 | `rFonts=0` | 换转换器 |
| A3 | 表格结构丢失 | `tbl=0` | 换转换器 |
| A4 | 编号列表丢失 | `numPr=0` | 换转换器 |
| A5 | 纸张默认为 Letter | `PageSetup::default()` = 12240×15840 twips | 换转换器 |

### 2.2 通道 B：表格规则（**自研可解**）

| # | 缺口 | 位置 | 修法 |
|---|---|---|---|
| B1 | `tblLayout` autofit/fixed 未实现，表格溢出正文 | `parse.ts:1240` | 解析 `w:tblLayout`，autofit 时按比例收敛到正文宽 |
| B2 | `tblW type="auto"` 未识别 | `parse.ts:1268` | `auto` 标记需自适应 |
| B3 | `trHeight` 的 `hRule`（`atLeast`/`exact`）未实现 | `parse.ts:1286` | 解析 `w:hRule`，`exact` 用 minmax，缺失按 `atLeast` |
| B4 | `tblCellMar` 自定义边距未实现 | `parse.ts:1240` | 解析表格级与单元格级 `tblCellMar` |
| B5 | `tblInd` 表格左缩进未实现 | `parse.ts:1240` | 解析并叠加到表格容器 |
| B6 | `tblBorders` 表格级边框与单元格级优先级 | `parse.ts:1248` | 已做合并，需补 `insideH/insideV` 缺省继承 |
| B7 | `w:tcW type="pct"` 未处理 | `parse.ts:1300` | 百分比单元格宽度 |

### 2.3 通道 C：排版细节（**自研可解**）

| # | 缺口 | 位置 | 修法 |
|---|---|---|---|
| C1 | `snapToGrid` 未解析 | `parse.ts:1523` | 解析 `w:snapToGrid`，关闭时该段不吸附网格 |
| C2 | `docGrid type="linesAndChars"` 每行字数未实现 | `metrics.ts:100` | 实现每行字数约束 |
| C3 | `w:cols` 分栏未排版 | `parse.ts` | 分栏容器布局 |
| C4 | 默认字体为空的兜底不稳 | `parse.ts` | `defaultFont` 兜底到 `docDefaults` |
| C5 | 段落级 `w:rFonts`（段落标记字体）未参与行高 | `metrics.ts:118` | 段内最大字号 + 段落标记字体共同决定 |
| C6 | `keepNext` 链跨段绑定不完整 | `pagination.ts` | 建立 keepNext 链，链整体下移 |

---

## 3. 实现方案

### 3.1 总体架构（不变内核，补三块）

```
              ┌──────────────────────────────────────┐
  .docx ──────▶│  自研 OOXML 解析（现有 + 本方案补全） │
              │  styles/numbering/theme 全链解析     │
              └──────────────┬───────────────────────┘
                             │
              ┌──────────────▼───────────────────────┐
              │  文档模型 IR + 样式解析器（新增）      │  ← C4/C5
              └──────────────┬───────────────────────┘
                             │
          ┌──────────────────┴──────────────────┐
          │                                     │
  ┌───────▼────────┐                 ┌──────────▼─────────┐
  │ 排版引擎         │                 │ 保真转换通道        │
  │ 度量+分页+表格   │                 │ .doc → .docx        │
  │ (B1–B7, C1–C3) │                 │ (A1–A5, 产品决策)    │
  └───────┬────────┘                 └────────────────────┘
          │
  ┌───────▼────────┐
  │ 渲染层（DOM）   │
  └────────────────┘
```

**保持不变的决策**：ADR 0001（自研编辑内核）不变。本方案**不引入 ONLYOFFICE/SuperDoc**（AGPL 传染），**不把 LibreOffice 塞进安装包**。

### 3.2 通道 A：`.doc` 保真三级策略（**需拍板**）

`.doc` 的信息在二进制格式里是**存在的**，只是 `office_oxide` 没读。理论上可以自研 MS-DOC 解析器补齐，但那是重写一个 Word 内核（工作量以季度计，且易错），不划算。务实做法是三级：

| 优先级 | 实现 | 安装包增量 | 保真度 | 触发条件 |
|---|---|---|---|---|
| **1** | 调用本机 **Microsoft Word** 自动化（AppleScript / COM） | 0 | ★★★★★ 最高，与 Word 完全一致 | macOS 已装 Word |
| **2** | 调用本机 **WPS** 自动化 | 0 | ★★★★☆ | macOS 已装 WPS |
| **3** | 可选 **LibreOffice** 组件（按需下载、独立进程） | 0（按需下载 110–160MB） | ★★★★☆ | 用户显式开启 |
| 4 | 现有 `office_oxide` 文本兜底 | 0 | ★☆☆☆☆ 仅文字 | 上述都不可用 |

**产品语义（必须）**：

1. `.doc` 打开为**只读兼容会话**，标题栏标注「兼容模式（DOC）」。
2. 转换在**内存/沙盒目录**完成，源目录零写入（V1 已实现，保留）。
3. 转换器不可用时，**明确提示**「此文档为旧版 .doc 格式，未安装 Word/WPS，无法还原版式，仅能显示文字」，**禁止静默降级**成「看起来能打开但版式全错」。
4. 兜底通道必须产出**质量报告**：哪些信息丢失（字体/表格/页面），让用户知情。

> **决策点**：是否接受「安装包零增量，但需用户本机有 Word/WPS 或手动下载 LibreOffice」这一前提？若接受，`.doc` 保真在 1 天内可交付；若不接受，`.doc` 只能停留在文本级，与用户期望差距无法弥合。

### 3.3 通道 B：表格引擎（自研，优先级最高）

#### B1 + B2：`tblLayout` 自适应列宽

```ts
// packages/editor-word/src/parse.ts —— 表格解析新增
interface DocTable {
  // ...既有字段
  layout?: 'fixed' | 'autofit';       // w:tblLayout@w:type
  indentPx?: number;                   // w:tblInd
  cellMarPx?: { top: number; right: number; bottom: number; left: number };
}
```

**Word 的 autofit 算法（近似，但收敛到正确结果）**：

```ts
/**
 * 功能：按 Word 的自动调整规则把 tblGrid 列宽收敛到正文可用宽度；时间：2026-11-03；作者：lq。
 *
 * Word 对 tblW=auto 的表格会按内容重排列宽，但无论内容如何，
 * 结果一定落在正文可用宽度内。表格超宽时 Word 按比例压缩（而非裁剪），
 * 因此这里用「等比缩放」逼近，可保证列宽相对比例与总宽都与 Word 一致。
 */
function fitColumnsToContent(
  columnWidths: number[],
  contentWidthPx: number,
): number[] {
  const total = columnWidths.reduce((a, b) => a + b, 0);
  if (total <= 0 || total <= contentWidthPx) return columnWidths;
  const ratio = contentWidthPx / total;
  return columnWidths.map((w) => Math.max(12, Math.round(w * ratio * 100) / 100));
}
```

**关键**：判定条件要区分 `tblLayout`：

- `w:tblLayout w:type="fixed"` → **不缩放**，严格用 `tblGrid`（用户显式固定列宽）。
- `w:tblLayout w:type="autofit"` 或缺省且 `tblW type="auto"` → **缩放到正文宽**。
- `w:tblW w:type="pct"` → 按百分比，`widthPct` 已实现。
- `w:tblW w:type="dxa"` → 按绝对值。

> 注意：`tblLayout` 缺省值在 OOXML 里是 `autofit`，但**只在 `tblW=auto` 时才生效**。若 `tblW` 是固定 dxa 而总和超宽，Word 会**压缩**（不是裁剪），仍需缩放。

#### B3：`trHeight` 的 `hRule`

```ts
export interface DocTableRow {
  // ...既有
  heightRule?: 'atLeast' | 'exact';   // w:trHeight@w:hRule，缺省 = atLeast
}
```

渲染语义：

| hRule | CSS | 说明 |
|---|---|---|
| `exact` | `height: Npx` | 固定行高，内容溢出被裁 |
| `atLeast`（含缺省） | `min-height: Npx` | Word 语义；内容更高时撑开 |

> 当前实现是 `minmax(row.heightPx ?? 20px, auto)`，**恰好等价于 `atLeast`**——所以缺省情况其实是对的。但 `hRule="exact"` 的行（合同/表格常见）会被错误地撑高。**这是需要补的唯一分支。**

#### B4/B5：`tblCellMar` 与 `tblInd`

```ts
// 优先级：单元格 tcMar > 表格 tblCellMar > Word 默认（上下 0，左右 108 twips = 5.4px）
```

> 注意：Word 默认 **上下为 0**、左右 108 twips。V1 已把默认改为 `0 5.4px`，方向正确；缺的是**显式 `tblCellMar` 的覆盖**。

#### B7：`tcW type="pct"`

单元格宽度百分比需按**表格实际宽度**换算，而非页面宽度。

### 3.4 通道 C：排版规则补全

#### C1：`snapToGrid`

```ts
// parse.ts
const snapToGrid = toggleValue(firstChild(pPr, 'w:snapToGrid'));  // 缺省 true
```

语义：`snapToGrid=false` 的段落**不吸附文档网格**，行高用自己的字体度量，而不是 `linePitch`。当前实现对所有段落一律用 `linePitch`，这是错的——只要文档里有一个 `snapToGrid=false` 的段落（如表格内段落、混排段），它就会算错。

#### C2：`linesAndChars` 每行字数

```ts
// 每行字数 = floor((正文宽 - 左缩进) / (字号 + charSpace))
// 行高 = linePitch（不变）
// 排版时按该字数断行，超出部分压缩或换行
```

实现上，Word 的字符网格会让**同一行内的字符间距被拉伸/压缩**到刚好填满网格。CSS 无法直接表达，需要在段落容器上用 `text-align: justify` + 手动字距，或用 `letter-spacing` 计算值。对公文（每页 22 行 × 每行 28 字）这是**必须项**，否则行尾位置与 Word 差 1–2 字。

#### C4/C5：样式链与行高口径

`defaultFont` 兜底链：

```
run.rFonts → paragraph.style.runFonts → styleChain(docDefaults → basedOn 链) → 内置默认
```

行高字号取值（`metrics.ts:118` 现有逻辑需扩展）：

```
段内所有 run 的最大字号 ∪ 段落标记（w:pPr/w:rPr）的字号
```

> Word 规则：段落标记的字号参与本段最后一行的行高计算。合同类文档里「段落标记字号 ≠ 正文」很常见，是行高偏差来源之一。

#### C6：`keepNext` 链

现有实现是逐段判断。Word 的语义是**链式绑定**：若 A.keepNext 且 B.keepNext，则 A+B+C 必须同页。需在分页前构建链，再整链装箱。

### 3.5 度量层加固（承接 V1 P1.1）

V1 已建立 Canvas 度量层。需补两点：

1. **字体加载时序**：`document.fonts.ready` 前后度量结果可能不同（回退字体 vs 真实字体）。必须在字体加载完成后 `resetFontMetrics()` + 重排（V1 已做，需补**重排触发时机**）。
2. **度量缓存跨文档**：当前缓存键是「字体栈 + 字号」，跨文档复用是安全的；但**不同 `docGrid` 下同一字号行高不同**（网格优先），需把 `gridLinePitchPx` 纳入缓存键，否则切换文档会读到上次的值。

> ⚠️ 这是 V1 遗留的**潜在 bug**：`metrics.ts:64` 的缓存键为 `${stack}|${sizePt}`，但 `lineHeightPx` 在有网格时返回 `gridLinePitchPx`。虽然 `lineHeightPx` 本身不查缓存（只有 `singleLineHeightPx` 查），所以当前**没有实际错误**；但一旦后续给 `lineHeightPx` 加缓存就会踩坑。建议**现在就把 `gridLinePitchPx` 纳入 `singleLineHeightPx` 的调用约束**，或在该函数注释中明确「不含网格」。

---

## 4. 验证方法

### 4.1 语料与基准

| 项目 | 说明 |
|---|---|
| 语料集 | 50 份真实文档（合同、公文、通知、报告、请假单、报价表），覆盖单节/多节、合并单元格、页眉页脚、图片、文本框、`.doc` 与 `.docx` 各半 |
| 基准图 | 用 **Microsoft Word**（首选）与 WPS 各导出 PDF，离线生成入库；CI 不需装 Office |
| 离线转换 | `soffice --headless --convert-to pdf` 仅用于**生成基准**，不作为运行时依赖 |

### 4.2 五项量化指标（全部自动判定）

| 指标 | 定义 | 目标 | 判定方式 |
|---|---|---|---|
| **M1 页数一致率** | Mini Office 页数 == Word 页数的文档占比 | ≥ 95% | 自动化 |
| **M2 分页点一致率** | 同一段落落在同一页的比例 | ≥ 90% | 按段落 id 对齐 |
| **M3 行盒偏差** | 文本行基线 y 坐标归一化后的中位偏差 | ≤ 2px @96dpi | `getBoundingClientRect` |
| **M4 列宽偏差** | 表格列边界 x 坐标序列的相对偏差 | ≤ 1% | 取列边界序列比对 |
| **M5 视觉相似度** | 逐页栅格化后 SSIM | ≥ 0.90 | 页面截图比对 |

补充指标：

- **M6 页码正确率**：`PAGE` / `NUMPAGES` / `SECTIONPAGES` 求值正确率 = 100%
- **M7 字体还原率**：run 级字体族与原文档一致的比例（`.docx` 应达 100%）

### 4.3 三层验证流水线

```
L1 单元测试（CI 每次提交）
   ├─ 解析层：给定 fixture → 断言 IR 字段（页边距、列宽、hRule、网格、字体链）
   ├─ 度量层：行距换算、网格行数、字号口径
   └─ 分页层：keepNext 链、孤行控制、表格切页、表头重复

L2 集成测试（每日）
   ├─ 真实 fixture 解析 → 序列化 → 再解析（往返一致性）
   ├─ 表格 autofit 收敛断言：列宽之和 == 正文宽 ± 1px
   └─ 页数断言：与入库的 Word 基准页数比对

L3 视觉回归（发版前）
   ├─ 渲染 50 份语料 → 逐页截图 → 与基准图算 SSIM
   └─ 低于阈值 → 报告列出偏差最大的文档与段落
```

### 4.4 本轮（通道 B+C）的具体验收

用已在手的三个样本做定点验收：

| 样本 | 断言 |
|---|---|
| `测试请假单.docx` | 表格列宽之和 == 554px（正文宽）± 1px；`trHeight` 缺省按 `atLeast`；`vAlign=center` 生效；表格不再溢出右边界 |
| `测试分包合同.doc`（经转换） | 明确标注「.doc 保真不可用」并列出丢失项；**不再声称能还原版式** |
| `采购文件.doc`（经转换） | 同上；且**转换器可用时**页数与 Word 一致（M1） |
| `合同.docx`（fixture） | `docGrid.type=lines`、`linePitch=24px` 保持；`keepNext` 链不分页 |

### 4.5 度量「修复前后」对比模板

每次修复必须在 PR 里附：

```
样本：测试请假单.docx
指标：M1 页数、M4 列宽偏差、M5 SSIM
修复前：M1 = 1/1, M4 = 8.9%, M5 = 0.82
修复后：M1 = 1/1, M4 = 0.3%, M5 = 0.94
```

---

## 5. 实施计划

| 阶段 | 内容 | 交付 | 预估 |
|---|---|---|---|
| **Q1**（通道 B） | B1–B7 表格引擎：`tblLayout` 自适应、`hRule`、`tblCellMar`、`tblInd`、`tcW pct` | 请假单表格与 Word 对齐（M4 ≤ 1%） | 1.5 天 |
| **Q2**（通道 C） | C1 `snapToGrid`、C4 字体兜底链、C5 段落标记字号、C6 `keepNext` 链 | 行盒偏差 M3 ≤ 2px | 2 天 |
| **Q3**（通道 C 进阶） | C2 字符网格、C3 分栏 | 公文每页行数/每行字数与 Word 一致 | 2 天 |
| **Q4**（通道 A 决策） | 实现 Word/WPS 自动化转换；无转换器时的明确提示与质量报告 | `.doc` 保真或**诚实降级** | 2 天 |
| **Q5**（验证体系） | 50 份语料 + Word 基准 PDF + M1–M7 指标 + CI 门禁 | 可量化验收 | 2 天 |

**建议顺序**：Q1 → Q2 → Q5（先建验证体系再改，避免改完无法证明）→ Q4 → Q3。

### 5.1 体积与性能预算

| 项目 | 安装包增量 | 备注 |
|---|---|---|
| 表格引擎 + 排版规则 | 0 | 纯逻辑 |
| Word/WPS 自动化 | 0 | 仅 macOS AppleScript 调用 |
| LibreOffice 可选组件 | 0（按需下载） | 独立进程，不进安装包 |
| 度量兼容字体子集 | ≤ 2.5 MB | 已有规划 |
| **默认安装包合计** | **≤ 3 MB** | 维持 V1 预算 |

性能：10 页 DOCX 首屏 ≤ 1.2 s；分页重算 ≤ 80 ms/次。

---

## 6. 风险与边界

| 风险 | 影响 | 缓解 |
|---|---|---|
| Word/WPS 自动化在无 GUI 会话下失败 | `.doc` 通道不可用 | 降级到 LibreOffice；再降级到明确提示 |
| 自动化调用涉及 Office 授权条款 | 合规风险 | 优先走「用户主动触发 + 本机已装」路径，不做批量无人值守 |
| 字符网格实现复杂 | 公文类可能仍与 Word 有 1–2 字偏差 | 列为 Q3，可延后；用 M5 SSIM 兜底判断 |
| `.doc` 自研解析（备选） | 工作量以季度计 | **不推荐**，除非产品明确要求零外部依赖 |
| F3 像素级（分栏/文本框环绕/SmartArt） | 不承诺 | 保留原部件 + 兼容性报告标红 |

### 明确不做

- 不追求 F3 像素级一致（浮动对象编辑、图文混排编辑）。
- 不内置 LibreOffice 到安装包。
- 不引入 AGPL 组件（ONLYOFFICE sdkjs / SuperDoc）进入闭源分发链路。
- 不为旧格式实现原格式写回（`.doc` 只读打开 + 另存为 `.docx`）。
- 不自研 MS-DOC 二进制解析器（除非产品另有决策）。

---

## 7. 需要决策的事项

1. **`.doc` 通道是否接受「需本机 Word/WPS 或按需下载 LibreOffice」的前提？**
   - 接受 → Q4 可交付 `.doc` 高保真。
   - 不接受 → `.doc` 只能到文本级，必须在 UI 上明确告知用户，避免「以为坏了」。

2. **是否建设 50 份语料 + Word 基准 PDF 的验证体系（Q5）？**
   - 这是把「像不像」变成可回归数字的唯一办法，但需要一次性投入约 2 天 + 语料收集。

3. **字符网格与分栏（Q3）是否在本期范围内？**
   - 公文类文档强依赖；合同/报告类影响较小。

---

## 附录 A：`.doc` 转换损失实证命令

```bash
# 1) 触发真实转换
cargo run -p office-compat --example probe -- "/path/测试分包合同.doc" /tmp/out.docx

# 2) 解包并统计版式元素
cd /tmp && unzip -o -q out.docx -d out_x
python3 - <<'PY'
import re
x = open('/tmp/out_x/word/document.xml', encoding='utf-8').read()
for tag in ['w:sectPr','w:pgSz','w:pgMar','w:headerReference','w:footerReference',
            'w:rFonts','w:tbl>','w:tblGrid','w:tblW','w:numPr','w:ind','w:spacing','w:sz']:
    print(f'{tag:22} {len(re.findall(re.escape(tag)+r"[ />]", x)):5d}')
PY
# 预期输出：pgSz/pgMar/rFonts/tbl 均为 0
```

## 附录 B：表格列宽校验脚本（Q1 验收用）

```ts
// 断言：表格列宽之和 == 正文可用宽度（±1px）
const contentWidth = doc.page.widthPx - doc.page.margins.left - doc.page.margins.right;
for (const block of doc.blocks) {
  if (block.type !== 'table') continue;
  const total = block.columnWidthsPx.reduce((a, b) => a + b, 0);
  expect(Math.abs(total - contentWidth)).toBeLessThanOrEqual(1);
}
```
