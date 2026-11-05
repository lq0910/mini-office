# 固定公开 Word 回归语料

功能：登记样本来源与可复现边界；时间：2026-11-03；作者：lq。

来源仓库与完整提交见 manifest.json，当前固定为 e1f50e90d246f3ecc6427d2a0af90f04f37810ec。文档、纯文本预期、生成脚本及公共 helper 都单独记录 SHA-256，不在测试时联网下载。

| ID | 来源类型 | 覆盖能力 |
| --- | --- | --- |
| formatting | 上游 DOC 测试文档 | 字符格式、段落间距、中文 |
| table | 上游 DOC 测试文档 | 2×3 表格、列宽与边框 |
| page-setup | 上游 DOC 测试文档 | 自定义纸张与页边距 |
| text-basic | 上游 DOC 测试文档 | 中英文混排、Tab、空段落、emoji |
| br-only-line-size | 可重建合成 DOCX | 不同字号的换行 run |
| tab-stops-column-edge | 可重建合成 DOCX | 右对齐点线制表位、段落缩进 |
| grid-line-rules | 可重建合成 DOCX | 网格、倍数/固定/最小行距 |
| para-picture-only-line | 可重建合成 DOCX | 纯图片段落、不同高度、行内关系 |

## 来源与授权边界

- 上游源码的 Apache-2.0 许可与 NOTICE 位于上一层目录，保留原始作者归属。生成器仅增加集成注释，原内容的哈希也已登记。
- 四份合成 DOCX 由对应 build.py 与 _docx.py 生成，仅包含程序化文字及内置 1×1 PNG；按这些第一方生成器的源码许可登记。
- 四份 DOC 属于上游测试文件。上游源码说明使用 LibreOffice 导出，formatting 的测试还注明 python-docx 生成后另存为 Word 97。现有证据不等于每份二进制附件的独立授权说明；清单保留来源未单独确认的状态，不把它们描述为已完成附件授权审计。
- 这里是开发回归语料，不加入应用资源或安装包；不得把仓库中其他附件的授权自动类推为可复用。
- 未复制用户合同、请假单或采购文件；私人语料仍通过环境变量显式选择。

## 离线校验与重建

~~~bash
python3 scripts/verify-genoffice-fixtures.py
pnpm exec vitest run tests/word-genoffice-corpus.test.ts tests/word-fidelity-baseline.test.ts
~~~

第一条检查 8 份文档及所有已登记文本/生成器的哈希，在临时目录重建 4 份合成 DOCX，不覆盖源样本。Python 仅使用标准库。

ZIP 时间戳不作为重建差异；XML 用标准库规范化后逐部件比较。上游前三份无图片的合成文档，其已存二进制比同提交 helper 少一个未使用的 xmlns:r 声明；规范化会消除这个声明差异。其余 XML 结构、文字和二进制媒体仍参与比较，不能把这项结果描述为重建 ZIP 字节完全相同。

## 三层验收

1. 解析结构：对照正文纯文本、字号、制表位、网格、表格和图像关系。
2. 保存往返：无操作返回原字节；修改普通正文后重新打开，未编辑 ZIP 部件内容保持一致，原样本哈希不变。ZIP 目录条目不属于文档部件。
3. 视觉对照：用真实浏览器逐页截图，只有提供绑定原件哈希、应用版本和字体环境的 Word/WPS 基准后才作对照。当前所有样本的 Office 状态均为 pending。

视觉对照命令和外部基准格式见 docs/word-corpus-acceptance.md。上游 truth 数值和 GenOffice 截图只能作为参考，不能代替本项目指定环境的 Word/WPS 验收结果。
