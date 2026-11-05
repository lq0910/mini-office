# Office 兼容性样本集

该目录用于维护真实 Office/WPS 文件的兼容性回归样本。样本不提交敏感业务内容，只提交脱敏后的最小文件；每个样本需要记录来源应用、格式、覆盖能力和预期结果。

## 样本登记规则

每个样本至少覆盖一个高频能力：

- DOCX：页眉页脚、分节、表格合并、图片、段落边框、中文主题字体。
- XLSX：主题字体、数字格式、日期系统、合并单元格、公式、打印区域。
- PPTX：主题色、东亚字体、文本框内边距、形状填充、图片、母版占位符。

验收时记录：

1. 能否打开。
2. 首屏耗时。
3. 页数或幻灯片数量是否一致。
4. 编辑后是否能被 Office/WPS 重新打开。
5. 未修改部件是否保留。
6. 复杂对象是否给出兼容性提示。
7. 保存后相关 `.rels` 与 `[Content_Types].xml` 是否仍能定位并识别保留部件。

## 当前内置样本

| 文件 | 格式 | 覆盖能力 |
| --- | --- | --- |
| `../fixtures/docx/contract.docx` | DOCX | 中文合同、页眉、页脚、图片、表格、页码字段 |
| `../fixtures/xlsx/workbook.xlsx` | XLSX | 中文工作簿、公式、日期、金额格式、合并单元格 |
| `../fixtures/pptx/deck.pptx` | PPTX | 文本、备注、页面复制、对象增量回写 |

后续新增真实样本时，先脱敏并在此表登记，再加入对应的解析、保存和视觉回归测试。

## GenOffice 固定公开语料（2026-11-03）

功能：增加可离线复现的 Word 专项回归入口；作者：lq。

另有 8 份公开 Word 语料，清单与 SHA-256 位于 third_party/genoffice/fixtures/manifest.json。其中 4 份合成 DOCX 保留生成器和许可证据，四份 DOC 的附件来源证据边界单独记录，不把仓库许可类推为附件授权审计完成。

执行 python3 scripts/verify-genoffice-fixtures.py 校验哈希与生成结果。解析与保存回归见 tests/word-genoffice-corpus.test.ts；逐页浏览器验收与外部 Office 基准格式见 docs/word-corpus-acceptance.md，缺失 Office 基准时必须标为待验证。
