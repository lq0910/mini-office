# Word 定制内核 M0 执行记录

功能：记录基线整理的实际交付、核验方法和下一阶段入口；时间：2026-11-08；作者：lq。

## 结果

**M0 基线整理已完成。** 本页及 JSON 保存 M0 完成时的历史快照，当时 M1 尚未开始。后续 M1 核心提取现已完成，当前进度见 [M1 记录](../word-core-m1/README.md)。M0 本身对应[开发计划](../word-custom-core-development-plan-20261107.md)的第一个阶段，不切换 Word 编辑器，不改生产解析、编辑、保存或 UI。

| M0 任务 | 实际交付 |
| --- | --- |
| 工作区状态与可恢复基线 | `workspace.json`：`main` / `d965e7bb8decc3050da33056fb007ead81477090`，起点干净，350 个原有受版本控制文件的 SHA-256 |
| 固定来源与资源 | [UPSTREAM.md](../../packages/word-core/UPSTREAM.md)、`upstream.json`：固定提交归档及 2,495 个源码、测试和资源文件的哈希 |
| 依赖和锁定版本 | `dependencies.json`：上游 workspace 包声明、lockfileVersion 3 以及 1,210 条直接/传递包记录 |
| 源码导入审计 | `imports.json`：1,543 个 TS/TSX 文件的静态/动态导入和 re-export，为 M1 拆分依赖提供依据 |
| 现有产品能力 | [功能迁移清单](features.md)、`commands.json`：20 个源码文件中的 288 个静态入口记录，包括 28 个工具栏回调 |
| UI 与公共样本 | [截图清单](../../output/playwright/word-core-m0/manifest.json)：10 张本项目截图及控件状态、样本哈希，浏览器异常为 0 |

此前对话中工作区有未提交修改，但本轮开始时已提交为 `d965e7b`，包含保存、输入法、粘贴等修复及开发计划。本轮以实际状态为准，没有覆盖回退到 `ba23e50`，没有暂存、提交或清理其他工作。

## 证据范围

首次起点在新增 M0 文件前记录，位于忽略目录的 `output/word-core-m0.local/start.json`。可公开复核的版本与源码哈希同步保存在 `workspace.json`。恢复查看使用固定提交的只读 Git 内容，不以执行 `reset` 或覆盖当前工作区作为核验方式。

上游旧临时目录已经残缺，未继续作为来源。重新从 GitHub codeload 取得指定提交的完整归档，记录传输归档哈希及原始文件哈希。源码包有可信固定 URL 和逐文件校验；没有把一个残缺目录的路径或包版本号当作完整来源证明。

归档解包出的全部应用仅保存在忽略目录，供开发审计使用，不加入本项目工作区依赖或发布包。`rootsAreShippingSelection=false` 明确说明：这里记录较宽来源范围是为了分析耦合，不表示我们采用上游全部模块。

依赖记录是原始 workspace 声明及锁文件解析结果，并未声称已经抽取了可编译的 Word 运行时闭包。AST 清单中的混合类型导入、barrel 导出、CSS URL、worker 和平台资源还需要 M1 按实际使用核对。这正是下一阶段的工作，不伪装成 M0 已完成编译。

## 当前 UI 基线

截图由 `scripts/capture-word-core-ui.mjs` 打开仓库公共夹具 `tests/fixtures/docx/contract.docx`，不编辑正文、不保存文件、不读取用户业务文件。记录项目为：

- 浅色主题下的开始、插入、布局、引用、邮件、视图六个标签；
- 布局与属性面板、表格右键菜单；
- 深色主题，以及 900 × 700 窄窗口。

宽窗口为 1440 × 1000，deviceScaleFactor 为 1。浏览器和公共样本的具体版本、SHA-256、可见控件记录在截图清单中。截图已检查开始标签、表格菜单和窄窗口的可读性；它们是现有产品 UI 的基准，不证明 Word 排版已经高保真，也不替代 Tauri WebView 验收。

## 核验

在项目根目录执行：

```bash
# 功能：核验 M0 原有代码、上游来源与 UI 截图；时间：2026-11-08；作者：lq。
node scripts/word-core-baseline.mjs --verify
node --check scripts/word-core-baseline.mjs
node --check scripts/capture-word-core-ui.mjs
git diff --check
```

校验程序比较记录中的 350 个原有文件、2,495 个上游文件和归档哈希，并核验命令来源、UI 样本、提交和截图。本轮实际结果见 [verification.json](verification.json)，另已检查重复记录会失败且不会覆盖原基线。正常迁移从 M1 起会主动修改部分生产文件，因此这是一份“迁移前状态是否仍一致”的工具，不是要求后续阶段永远不能改变源码。

本轮是基线整理，没有修改生产代码或原有锁文件，因此不重复运行全量应用测试和原生打包。实际执行了源码哈希、脚本语法、差异检查和浏览器捕获；不存在本阶段已经完成新核心构建或产品验收的含义。

## 从固定来源重新取得审计副本

归档及展开源码属于可删除的本地开发材料；版本、来源与逐文件哈希由受版本控制清单保留。清理忽略目录后，可按以下命令重新取得同一源码，再核验。命令不会安装依赖或覆盖生产模块。

```bash
# 功能：恢复固定提交的只读审计副本；时间：2026-11-08；作者：lq。
mkdir -p output/word-core-m0.local/upstream
curl --fail --location 'https://codeload.github.com/genspark-ai/genoffice/tar.gz/e1f50e90d246f3ecc6427d2a0af90f04f37810ec' --output output/word-core-m0.local/genoffice-e1f50e90d246f3ecc6427d2a0af90f04f37810ec.tar.gz
tar -xzf output/word-core-m0.local/genoffice-e1f50e90d246f3ecc6427d2a0af90f04f37810ec.tar.gz --strip-components=1 -C output/word-core-m0.local/upstream
node scripts/word-core-baseline.mjs --verify
```

现有基线 JSON 和截图不可被工具自动覆盖。新阶段需要新基线时另建有阶段编号的记录，不能通过重录旧文件把差异隐藏掉。`--record` 与 `--commands` 用于首次建立，记录已存在时会拒绝覆盖。

## M1 的直接输入与已识别结合处

1. 以 `packages/docx-engine` 的完整模型、原始部件和补丁保存为基础，结合 renderer 的编辑转换、扩展与布局提取 `word-core`。
2. 上游 `editor/` 目录的静态导入审计未发现直接 `react` 或 `@tiptap/react` 导入；完整 App 则使用 React 19。这支持先提取非 React 核心的方向，但不能据此宣布整个依赖闭包已兼容现有 React 18。
3. 字体入口含 Node 字体读取，renderer CSS 又引用共享 UI 字体。需要分别接入 Tauri 字体能力和必要渲染资源；不能只复制 `font-metrics/resolve` 就声称接入全部字体能力。
4. `docx-engine` 有对 `pptx-engine`、`agent-core` 的真实底层引用，renderer 又有对 UI、平台和 AI 工作流的耦合。`imports.json` 用于逐项拆分，不能整包删除，也不能把完整其他编辑器搬入产品。
5. 所有工具栏命令直接融合到我们现有入口。28 个回调、动态菜单、快捷键和文件规则见功能清单；不要用保留按钮代替实际迁移。
6. 保存要承接新核心的修订版本和来源基线，统一手动保存、快照和关闭保存；旧 `DocModel` 不能作为新正文的中间转换格式。

当前只有 M0 文档、清单、基线工具与截图发生新增/修改；`word-core` 尚无生产源码，产品默认仍是既有 Word 实现。
