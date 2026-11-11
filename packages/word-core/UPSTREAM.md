# Word 定制内核的上游基线

功能：固定 GenOffice Word 源码起点、来源与提取边界；时间：2026-11-08；作者：lq。

当前状态为 **M0–M3 已完成对应阶段范围；2026-11-11 完成工具栏、区域编辑、邮件合并及打印/PDF 的开发入口验收**。本目录为我们持续维护的 GenOffice 衍生定制内核；Mini Office 保留 React 18 和自己的外壳/工具栏。默认产品切换仍按 M5 执行。实际交付与检查见 [M1 执行记录](../../docs/word-core-m1/README.md)及 [M2 执行记录](../../docs/word-core-m2/README.md)。

功能：同步 M1 提取范围、定制与构建状态；时间：2026-11-08；作者：lq。

## 固定来源

- 仓库：[genspark-ai/genoffice](https://github.com/genspark-ai/genoffice)。
- 开发起点：`e1f50e90d246f3ecc6427d2a0af90f04f37810ec`，与当前项目已使用的 DOC 转换和公开语料版本一致。
- 此起点是可复现的源码基线；不声称已核对为用户实际验证的具体版本，也不重新开展上游选型验证。
- [固定提交归档](https://codeload.github.com/genspark-ai/genoffice/tar.gz/e1f50e90d246f3ecc6427d2a0af90f04f37810ec)的本次 SHA-256：`25ddd396692a9d3b047601af007a464d39839d416b89efbe5d916e9035b02be3`，大小 35,540,156 字节。
- 上游 `package-lock.json` 的 SHA-256：`eb780f175d8638d18445ed5a15086a393cdc42af42ab16d2231d71aaf49bac98`；lockfileVersion 为 3。

归档和展开源码保存在仓库忽略目录 `output/word-core-m0.local/`，供开发核对。公共可复核清单在 [M0 基线目录](../../docs/word-core-m0/README.md)。后续构建只能依赖实际提取进项目并接受版本控制的核心代码，不能依赖此忽略目录或旧 `/private/tmp` 副本。

## 已记录的源码范围

| 来源 | 基线处理 | M1 接入约束 |
| --- | --- | --- |
| `packages/docx-engine` | 完整源码、测试及资源哈希 | 保留原始包、锚点、完整模型、解析和补丁保存；不转回旧 `DocModel` |
| `apps/docs/src/renderer/editor` | 编辑扩展、转换和表格等导入清单 | 承接事务、选区、撤销及活动子编辑器；不只挂载默认 Tiptap |
| renderer 的 `pagination-*`、样式和行度量 | 所属文件及引用均已记录 | 成组提取布局规则及 App 协调代码，不替换成旧分页器 |
| `packages/font-metrics` 与 renderer 字体 | 数据、源码和资源哈希 | 区分浏览器可用规则与 Node 字体读取，必要资源单独列明 |
| `packages/zip-gate` | 源码和声明 | 保留必要 ZIP 限制与原始部件处理 |
| `packages/agent-core`、`pptx-engine` | 为 Word 必需的底层引用建立审计材料 | 按实际子路径提取；不引入 AI 服务或 PPT 产品界面 |
| `packages/ui`、`i18n`、`electron-utils` | 完整记录以分析 UI、字体和平台耦合 | 不能把完整 UI 或 Electron 桌面接口直接搬入产品 |
| `packages/doc2docx` | 与现有旧格式转换对应的源码基线 | 保留 `.doc` 的转换和另存语义，现有转换代码不在 M0 覆盖 |
| `apps/docs` 的 App、Ribbon、preload、main | 为提取边界保留来源证据 | 提取必需的协调与平台行为，排除原应用 UI 和外壳 |

`upstream.json` 中有 2,495 条文件记录；它是来源范围，不是最终产品引入名单。`imports.json` 记录 1,543 个 TS/TSX 模块的导入，`dependencies.json` 记录原始包声明和锁文件中的 1,210 条记录。原始依赖清单包含上游开发工具与其他应用耦合，不能直接复制进我们的生产依赖。

M1 已按实际闭包形成 [SOURCE-MANIFEST.json](SOURCE-MANIFEST.json)，共 326 项源码/资源及逐项修改。新增发现的 Word 形状几何底层引用从同一固定归档逐字核对，未加入其他应用的产品界面。正常构建和运行仅使用本目录源码、资源及锁定依赖，不访问 M0 审计副本。

功能：M2 补充同源分页附属区域的提取记录；时间：2026-11-09；作者：lq。新增 `editor/note-dom.ts`、`hf-heights.ts`、`hf-phantom.ts`，当前映射共 329 项；保持同一固定提交、哈希核验及许可。产品宿主对上游 App 中显示协调逻辑的提取和改造见 `packages/editor-word/src/core-layout*.ts`，不引入上游 React App。

功能：M3 在自有 `src/editing/formatting.ts` 增加格式摘要和局部字符/段落事务，现有工具栏直接消费；时间：2026-11-10；作者：lq。分页按钮复用同源 `editor/page-break.ts`，未修改 329 项上游衍生文件，也未引入上游 Ribbon；会话入口补充活动子编辑器命令目标和历史状态。证据见 [M3 记录](../../docs/word-core-m3/README.md)。

## 锁定版本与后续安装

| 关键包 | 固定上游锁文件的解析结果 | M1 处理 |
| --- | --- | --- |
| `@tiptap/core`、`@tiptap/pm`、`@tiptap/extensions` | 均为 `3.31.0` | 配套承接，不能混用不同版本的编辑运行时 |
| `@tiptap/react` | `3.31.0` | 原始记录保留；优先以非 React 核心挂载现有 UI，不自动安装该包装层 |
| React / React DOM | 均为 `19.2.7` | 这是上游锁定结果，不是本项目升级决定；Mini Office 继续 React 18 |
| JSZip | `3.10.1` | 按新核心构建验证后纳入本项目 pnpm 锁文件 |
| fast-xml-parser | `5.10.1` | 原始声明范围为 `^5.3.4`，以锁定结果识别来源，不能混淆声明和解析版本 |
| utif2 | `4.1.0` | 核对 Word 图片解码依赖后决定提取方式 |
| opentype.js | `2.0.0` | 保留实际字体度量语义，不能以字体名称枚举替代 |

M0 阶段没有安装依赖。M1 已将 Tiptap 三个配套包、JSZip、fast-xml-parser、utif2 的精确版本写入本包及项目 `pnpm-lock.yaml`，没有引入 React 19。实际版本由 `scripts/verify-word-core-source.mjs` 核验，上游完整锁文件继续作为原始参考。

## 来源与本地修改记录

功能：M3 第二批的 `editing/search.ts`、`editing/lists.ts` 及会话编号资源管理为本地新增；复用固定上游查找扫描和装饰协议，不修改提取文件。列表资源与正文快照配对保存，相关验证见 `docs/word-core-m3/README.md`；时间：2026-11-10；作者：lq。

功能：M3 第三批的 `editing/styles.ts`、`editing/format-painter.ts` 为本地新增；样式目录用原生文档属性事务共享历史及保存快照，复用上游样式补丁、OOXML 子元素排序和 CSS 生成，不修改提取文件；导入样式仅改名保留原定义 XML。固定行距/底纹等新增属性在本地保存后处理，验证和剩余边界见 M3 记录；时间：2026-11-10；作者：lq。

功能：M3 第四批 `editing/page-layout.ts` 为本地新增，页面及页眉页脚草稿存入原生文档属性事务；保存计划按目标分节补写尺寸/边距/编号，保留未知 XML、装订线和固定负边距。复用固定上游 `hf-sections`、`hf-heights` 与动态域序列化，不修改来源文件；产品显示层同步真实纸张坐标并回填分页总数缓存。边界和证据见 M3 记录；时间：2026-11-10；作者：lq。

功能：M3 第五批 `editing/ruler.ts`、`editing/toc.ts` 及产品标尺接线为本地新增，复用固定上游标题规则、TOC 域生成、字段缓存补写及分页切片；引导点度量复用已有外部重测信号。真实页码保存在匹配当前文档的布局缓存中，序列化只补入捕获快照；不改正文版本、撤销栈或原始解析包，不修改 329 项提取文件。标尺拖动结束才向同一历史提交。范围和限制见 M3 记录；时间：2026-11-11；作者：lq。

功能：M3 第六/七批同源新增提取 `insert-picture.ts`、`new-table.ts`、`table-borders.ts`、`table-ops.ts`、`table-properties.ts`，该批共 334 项。所有来源修改由 `scripts/extract-word-core.mjs` 重放并更新清单：表格段落来源属性/未知 XML 及隐藏纵向续格身份；表格尺寸命令读取实际节净宽；生成器保留未知表格/行/格属性与段落来源。`editing/objects.ts`、`table-source.ts`、`section-width.ts`、`paste.ts`、`picture-files.ts`、`statistics.ts` 及 `serialization/table-paragraph.ts` 为本地模块。分节宽度通过节点装饰传递，避免 DOM 重解析破坏图片选区。粘贴载荷、异步解码和历史边界使用各编辑器独立状态，不使用上游全局暂存。实际验证及复杂导入边界见 M3 记录；时间：2026-11-11；作者：lq。

功能：M3 最终新增同源 `editor/hf-editor.ts`，来源映射共 **335 项**；所有提取文件修改均由 `scripts/extract-word-core.mjs` 重放，`SOURCE-MANIFEST.json` 逐项记录哈希及修改原因；时间：2026-11-11；作者：lq。

- 页眉页脚普通富文本/布局表格直接编辑、变体和继承关系；格内完整段落属性及边框/底纹显示，属性级合并保留未修改的段距零值、字符单位和未知 XML。
- `convert.ts`、`extensions.ts` 及同源 schema 的区域列表/格式和当前焦点改造；`editing/region-history.ts` 让区域、样式目录和正文共享主历史，恢复投影时不提交过时内容。
- 文本框保存快照重新登记新图片媒体，解析预载文本框媒体，新绘图局部声明命名空间；重复保存及重开验证真实图片字节与几何。
- 高级字符效果、段落边框线型、列表/边框样式及样式集；安全删除导入定义时保留隐藏基底并迁移跨部件引用。
- 本地新增 `serialization/mail-merge*.ts`、`merge-fields.ts` 和 XML/段落补写模块，字段与记录包均使用新核心；产品 `core-output.ts` 和 `WordCoreOutput.tsx` 复用同源页面，未引入上游邮件/打印 UI。

最终类型/测试/Web/独立核心检查通过：756 项测试通过、3 项私人样本跳过，425 模块/61 项输出；M3 浏览器 131 项、M2 回归 24 项。范围及支持边界见 [M3 验收记录](../../docs/word-core-m3/README.md)，原生平台和默认交付继续按 M4/M5。

- 原始来源的 LICENSE、NOTICE、字体说明及具体资源文件均有独立哈希。字体不能一律套用仓库根许可；renderer 的字体说明和 `packages/ui` 的共享字体一起保留为资源依据。
- 当前已记录 renderer 与共享 UI 中的 50 个字体文件，尚未决定全部随产品打包；是否采用按 Word 内容渲染用途与实际依赖确定，界面字体不因上游存在就替换我们的 UI。
- M0 没有引入上游生产源码；M1 已逐文件登记“原路径 → 本地路径 → 本地改动”，并为集成代码增加 `lq` 功能与日期注释。原作者、根许可、字体许可和 EMF 许可保留。
- 主要定制为去除 React 翻译 Provider、提取纯数据 UI 子入口、剥离 Node 字体读取、注入图片剪贴板与解析 Worker、提取文档初始化、局部文档 CSS，以及表格纯文字原 XML 保留增强。核心回归见 `tests/word-core-session.test.ts`，行为与限制见 M1 记录。
- 上游版本更新采用有记录的修复合并，不跟随 `main` 自动漂移，也不因命名改成 Mini Office 而失去来源映射。

本文件是工程来源记录，不替代各资源的原始许可文本。实际引入文件及其资源的范围由 M1 清单进一步收敛。
