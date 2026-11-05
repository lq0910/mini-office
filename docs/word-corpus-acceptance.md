# Word 公开语料与逐页对照

功能：提供可执行的公开语料验收入口；时间：2026-11-03；作者：lq。

## 本轮范围

本轮推进路线图 P0：固定 8 份公开语料、4 个可离线重建的生成器、解析/局部修改保存回归，以及逐页浏览器截图与外部 Office PNG 对照入口。未改造编辑内核，未声称完成 P1 跨页编辑。

清单位于 third_party/genoffice/fixtures/manifest.json，来源与授权证据见同目录 README。私人文件没有加入这套语料。所有新脚本有功能、时间、作者 lq 注释。

## 命令

~~~bash
# 功能：离线校验固定语料并重建合成样本；时间：2026-11-03；作者：lq。
python3 scripts/verify-genoffice-fixtures.py

# 功能：结构、无操作保存、普通文字修改往返和基准约束测试；时间：2026-11-03；作者：lq。
pnpm exec vitest run tests/word-genoffice-corpus.test.ts tests/word-fidelity-baseline.test.ts

# 功能：启动独立本地 Vite，逐页捕获浏览器版面；时间：2026-11-03；作者：lq。
node scripts/verify-word-corpus-browser.mjs

# 功能：提供外部 Office 基准后执行严格验收；时间：2026-11-03；作者：lq。
MINI_OFFICE_OFFICE_BASELINES=/absolute/path/to/office-baselines node scripts/verify-word-corpus-browser.mjs --require-office
~~~

浏览器脚本默认在 127.0.0.1:4196 启动 Vite，完成后结束本次服务。可通过 MINI_OFFICE_TEST_URL 使用已启动的开发服务。该入口需要 Vite 的源码模块路径，不适用于静态生产预览。通过 MINI_OFFICE_CHROME 指定本机 Chromium；通过 MINI_OFFICE_CORPUS_IDS=grid-line-rules,table 缩小验收范围。

脚本优先使用 MINI_OFFICE_OFFICE_BASELINES 指定的目录；未设置时自动检查本机 `output/office-baselines.local/`。该目录不存在时仍记录待基准状态，不将浏览器回归当作 Office 验收。

输出位于 output/playwright/word-corpus/：index.html、report.json 和每页 PNG。该目录已忽略，避免将外部基准副本混入提交。每次报告仅列出当前选定样本，旧 PNG 不参与结果。

## Office 基准格式

在指定 Word 或 WPS 版本中打开清单登记的原件，记录实际安装字体，导出 PDF，再以 96 DPI、无裁边、无重采样导出逐页 PNG。DOC 基准必须来自原始 DOC，不能来自转换后 DOCX。

外部目录按样本 ID 分目录，例如：

~~~text
office-baselines/
  grid-line-rules/
    manifest.json
    source.pdf
    page-001.png
~~~

manifest.json 的填写格式如下。尖括号内容必须替换成真实环境和完整 SHA-256；示例本身不是可通过的基准。

~~~json
{
  "schemaVersion": 1,
  "application": "Microsoft Word",
  "version": "<实际版本及构建号>",
  "operatingSystem": "<实际系统及版本>",
  "exportedAt": "2026-11-03T12:00:00Z",
  "exportMethod": "<实际 PDF 导出方式及 PNG 转换工具版本>",
  "dpi": 96,
  "fonts": ["<实际字体名及版本，注明替代字体>"],
  "fixtureId": "grid-line-rules",
  "sourceSha256": "<fixtures/manifest.json 中的源文档 SHA-256>",
  "pdfFile": "source.pdf",
  "pdfSha256": "<导出 PDF 的 SHA-256>",
  "pages": [
    { "file": "page-001.png", "sha256": "<该 PNG 的 SHA-256>" }
  ]
}
~~~

application 只接受 Microsoft Word 或 WPS Writer，分别保留环境身份，不能把它们混为同一个验收目标。来源字段由制作者记录；程序能够校验文档绑定、PDF/PNG 完整性与字段格式，不能独立证明图像确实由该 Office 版本生成。

## 报告与退出码

- pending-office-baseline：尚无对应基准；已截图也只算浏览器回归。
- page-count-mismatch：页数不同，不能只比较有交集的页后算通过。
- dimension-mismatch：图像原尺寸不同，不做自动缩放。
- difference-detected：尺寸相同但有不同像素，逐页给出差异像素比例与通道绝对误差。
- pixel-identical：全部页与所提供 PNG 像素完全相同，仅表示这些样本与该环境的图像一致。
- failed 或报告 fatalError：打开、图片解码、脚本执行、基准字段、路径或哈希等失败。

默认命令用于浏览器回归，缺少基准或检测到版面差异时记录状态，打开或执行失败时退出 1。--require-office 要求所有选定样本都有有效基准且逐页完全一致，否则退出 1。不得把默认命令的退出 0 解释为 Office 保真通过。

像素差异包括字体替代、平台抗锯齿、缩放与布局影响，不等于“兼容百分比”。没有设置任意容差来掩盖差异；是否接受某处差异仍需按页人工确认。

## 后续

取得指定环境的 Office 基准后，先依据逐页报告定位断行、表格和对象位置差异，再推进逐行布局及跨页编辑。新增分节、文本框、浮动对象等样本时，先补来源、授权和结构断言，再接入同一验收入口。

## 本轮实际验证（2026-11-03）

- 固定语料校验通过：8 份文档哈希、4 份合成 DOCX 重建及生成器原内容哈希。
- 语料与基准测试 39 项通过；本次另增 3 项制表位渲染/几何回归，全量前端测试 586 项通过，3 项私人样本测试默认跳过。
- 类型检查、浏览器脚本语法检查与 Vite 生产构建通过。原有 React act 和 Vite 动态导入提示仍存在。
- 用户批准执行后，已在本机 Chromium 147.0.7727.15 通过应用“打开文件”入口打开全部 8 份公开样本，捕获 9 页 PNG；7 份各 1 页，纯图片段落样本 2 页。未捕获到页面运行错误。检查了表格、制表位和图片续页的代表截图。
- 首轮截图发现文字内部原始 Tab 未进入制表位布局，页码不靠右且无点线。现已修复显示占位、缩进坐标及同一行嵌套矩形宽度测量；保留原模型及源 XML。第二轮浏览器复验中，五个页码均与标题同线、点线可见，右边缘偏差最大约 0.111 CSS px。脚本新增真实几何断言，避免仅凭成功截图判定布局正常。
- 当前字体环境仍有 Calibri、Cambria、ＭＳ 明朝缺失提示，已保留在 report.json，不把替代字体截图称为原字体一致。
- 未提供指定 Word/WPS 环境的 PDF/PNG 基准，因此没有 Office 版面验收结果。
- 本轮修改在本地工作区，未提交、推送或发布；没有重新构建 Tauri 安装包。

## 基准找回与对照（2026-11-04）

功能：登记此前已导出的 Word 页面基准，避免重复导出；时间：2026-11-04；作者：lq。

- PDF、PNG 与按 PDF 页创建时间生成的清单位于被 `.gitignore` 忽略的 `output/office-baselines.local/`，共 8 份 PDF、9 张 PNG。清单绑定本仓库 8 份原件 SHA-256；PDF、PNG 文件哈希均已校验。
- 用户确认这些文件由 Microsoft Word 导出。本机 Word 版本为 16.114.26101123，PDF Producer 为 macOS Quartz PDFContext；PDF 创建时间为 2026-11-03 22:14:19–22:14:48 CST。PNG 页面尺寸符合 96 DPI，页数与 Mini Office 浏览器捕获一致。
- 使用 Poppler `pdffonts` 检查，9 页所用字体均为嵌入子集且有 Unicode 映射。检测到 Cambria、MS Mincho、Calibri、Arial、Microsoft YaHei、Apple Color Emoji 和 Times New Roman；PDF 不提供嵌入字体文件的原始发行版本号。
- 与现有 Mini Office PNG 逐页做原尺寸像素比较，9 页尺寸全部一致，但没有一页像素完全相同。不同像素比例范围约 0.18%–15.28%；最大差异出现在纯图片段落样本第 1 页，需优先检查图片尺寸和段落布局。该像素比例包含字体栅格化与实际布局差异，不直接等同兼容率。
- 8 份基准 manifest 均通过格式、源文件绑定和哈希验证（测试 21 项通过）。严格浏览器入口的联机重跑未完成：本轮本地 Vite 连接所需的提权审批服务返回 503；离线逐页像素差异已完成。
- 基准目录保留在本机且不会进入 Git。可在本机 Vite 已启动时运行：`MINI_OFFICE_OFFICE_BASELINES="$PWD/output/office-baselines.local" MINI_OFFICE_TEST_URL=http://127.0.0.1:1420/ node scripts/verify-word-corpus-browser.mjs --require-office`。若开发服务端口不同，替换 URL；不需要再次从 Word 导出。

## 图片差异修复与完整对照（2026-11-04）

功能：修复真实 Word 基准揭示的图片拉伸、纯图片行盒和效果占位差异；时间：2026-11-04；作者：lq。

1. 原图虽为 1×1 像素，OOXML 明确声明 160×80px；原显示采用 `object-fit:contain`，把图案缩成居中的正方形。改为按声明尺寸拉伸，渲染图案与图片框一致。
2. 对纯内嵌图片、单倍自动行距、无网格的段落，清除浏览器为文字下伸部预留的额外空隙；混排、倍数、固定/最小行距和网格保留原规则。
3. 解析并保留 `wp:effectExtent` 四边占位，内嵌行显示纳入占位；重建图片 XML 也写回这些值。此项不实现阴影/发光等效果本身的绘制，也不扩展浮动对象排版算法。
4. 浏览器语料验收新增 18 幅图的实际尺寸、不同高度图片底边、单倍图片行高、效果下方占位和截图边缘着色断言；不只核对图片 DOM 的宽高。基准目录自动发现，HTML 对照页显示每页差异像素比例与通道误差。

| 纯图片段落样本 | 修复前不同像素比例 | 修复后不同像素比例 | 修复前/后通道平均绝对误差（0–255） |
| --- | ---: | ---: | ---: |
| 第 1 页 | 15.286% | 3.433% | 11.874 / 0.748 |
| 第 2 页 | 3.819% | 0.902% | 2.863 / 0.144 |

完整运行使用 Chromium 147.0.7727.15 和本机 1420 开发服务，8 份基准全部自动绑定、9 页完成原尺寸比较，页数/图像尺寸全部匹配，运行错误为 0。所有样本仍有不同像素，严格命令按设计退出 1；这不等于脚本执行失败，也不能声明 Word 像素验收通过。图片样本剩余差异包括替代字体、倍数/混排图片行及文字位置；下一步应逐项核验，不能依靠全局偏移调图。

全量前端测试 588 项通过、3 项私人样本测试按配置跳过；类型检查、Vite 生产构建和脚本语法检查通过。新增回归验证图片尺寸编辑后四边效果占位的保存回读。原有 React act、Vite 动态导入提示仍存在。未验证 Tauri WebKit/WebView2 实机；未重新生成安装包，未提交、推送或发布。

## 字体度量精度与图片倍数行距（2026-11-05）

功能：减少小字号字体边界取整误差，并修复纯图片自动倍数行距的重复下伸部占位；时间：2026-11-05；作者：lq。

- 首先使用当前代码重跑全部语料，复现上轮九页结果；修复前报告保存在本机 `output/playwright/word-corpus-before-metrics.local/`。
- 实测 Calibri/Carlito 11pt 的 Canvas 字体边界在实际字号下为 14px 上升部加 4px 下伸部，直接作为每行高度会积累取整误差。改为在 4096px 参考字号测量后按实际字号换算，当前环境的单行高度为约 17.904px。仅测量边界，不绘制大画布，也未增加字体文件或运行时依赖。实测未支持“混入汉字导致本轮西文度量变高”的初始猜测，因此没有据此改写字体后备栈。
- 无网格、纯内嵌图片的自动倍数行距继续清除文字行盒；将倍数相对单倍增加的字体间距加入每张图片帧的下方占位，并与原效果外扩相加。图片自动换行后每一行分别生效。固定值、最小值、网格、列表及文字混排仍使用原有文字行盒。
- 浏览器新增强制两张图片换成两行的断言：每行增加约 2.686px，两行总高实测 165.344px，计算期望为 165.371px，误差小于 0.1px。保留原有 18 幅图的尺寸、基线及绘制断言。

| 样本/页 | 修复前差异像素 | 修复后差异像素 |
| --- | ---: | ---: |
| br-only-line-size | 0.294% | 0.283% |
| tab-stops-column-edge | 1.285% | 1.205% |
| grid-line-rules | 1.005% | 1.022% |
| para-picture-only-line 第 1 页 | 3.433% | 3.382% |
| para-picture-only-line 第 2 页 | 0.902% | 0.879% |

其余四份旧 DOC 样本的截图未变化。网格样本的不同像素比例略增，仍需按行核对字体基线及网格规则；没有通过全局偏移或容差掩盖该结果。图片第 1 页通道平均绝对误差由 0.748 降为 0.627，第 2 页由 0.144 降为 0.066（范围 0–255）；差异像素比例也包含 PDF 栅格边缘和替代字体差异，不能换算成兼容率。

最终 Chromium 完整对照于 2026-11-05 08:59 CST 生成：8 份样本、9 页，页数和尺寸匹配，执行失败 0，像素完全一致 0。严格命令仍按设计退出 1。全量前端测试 592 项通过、3 项私人样本按配置跳过；类型检查、生产构建、脚本语法检查及 `git diff --check` 通过。原有 React act、Vite 动态导入提示仍存在。

本轮未更改原始文档和 Word 基准，也未验证 Tauri WebKit/WebView2、复杂文档跨页编辑或 Word 打开修改后文件的实际效果；未提交、推送或重新打包。
