# word-axmath-gpt-skill

将数学 PDF / DOCX 整理为 Word 练习资料，保留可编辑 AxMath 公式。Microsoft Word 和 AxMath 需单独安装，仓库自有代码采用 [MIT](LICENSE)。


## 首轮生成硬规则

**先定规则，再生成 DOCX；不允许把终稿审计当成补救流程。**

- 题号用英文半角普通文字，保留原序号、跳号、星号和逻辑层级；题文顶格，编号后紧接内容，不使用自动编号、编号域、悬挂缩进或 Heading 折叠身份。
- 全文关闭四项段落分页选项；文字与相邻公式之间不加空格，题图和图号分别居中。具体首轮与验收门槛见 [排版规则](docs/LAYOUT_RULES.md)。
- 生成前建立 SOURCE_FINAL_FORMULA_MAP.json；每个源公式填写 source_anchor 和来源 expected_tex；PDF/图片公式列为 source_formula_visual，正文独立变量/表达式列为 added_body_math，并记录准确 source_text。
- **单个字母、数字逐次登记**：一键转换不会识别普通正文中的数学字符。先对照原稿区分题号与数学内容，再按 [单字母/数字操作指南](docs/BODY_SINGLE_CHARACTER_MATH.md) 精确转换，避免漏转、重复和位置漂移。
- 公式必须通过官方 AxMath 宏生成真实 OLE：源 OMML 使用 AMSMML2AM，TeX/正文数学使用 AMSTeX2AM；宏调用前拆分普通说明文字与数学片段。下置 lim 用 \underset{...}{\lim}；大积分用 \int，禁 \smallint。
- 默认保留来源标点；仅用户明确要求时才用专用脚本替换普通正文中的中文句号并记录数量，不触碰公式、OLE、图像或字段。
- 每个源公式和新增数学对象都映射到终稿 OLE；一源多对象允许，但不允许漏映射或重复归属。
- 最终只做一次完整内部导出和一次 Word 逐页验收；问题按定位修复，不反复全书审计。

具体命令和映射格式见 [首轮生成工作流](docs/AXMATH_CONVERSION_WORKFLOW.md)。


## 从这里开始

**[SKILL.md](SKILL.md) 是唯一任务入口。** Codex / GPT 每次处理新文件先读它，并按任务加载下表所需文件。

| 场景 | 使用哪个文件 |
|---|---|
| 分页选项、编号/公式空格、左对齐、图和图号居中、样页验收 | [排版规则与排查入口](docs/LAYOUT_RULES.md) |
| 题目编号、缩进对齐、自动列表、Heading/WPS 折叠箭头 | [唯一编号规范](docs/NUMBERING_AND_OUTLINE.md) |
| 数学内容、题图、留白、分页、只交 Word | [数学资料排版工作流](docs/MATH_LAYOUT_REFERENCE.md) |
| MinerU PDF 转 DOCX | [MinerU API 说明](docs/MINERU_API.md) |
| OfficeMath 转原生 AxMath | [AxMath 转换详细工作流](docs/AXMATH_CONVERSION_WORKFLOW.md) |
| 正文中的单个字母、数字 | [单字符目标与转换](docs/BODY_SINGLE_CHARACTER_MATH.md) |
| Codex 本机直连 | [Codex 入口](docs/CODEX_DIRECT.md) |
| NEXT_ACTION 指定的少数异常 | [定向修复参考](docs/REPAIR_REFERENCE.md) |

**编号硬规则：** 英文半角普通文字题号，各级题文顶格，编号后无空格；标题不要内置 Heading 身份，取消折叠不能只设 `outlineLvl=9`。样页先通过检查，再批量做全文。

**MinerU 方法差异（请按优先级执行）：**
- GPT：在本机直接使用 API。
- Codex：先用当前对话侧边栏已登录 MinerU 网页；排队超 120 秒、60 秒无进展或卡顿则切本机 API。两者均有操作权限。
- 结果都交付 `mineru_original.docx` 进入同一个 AxMath 流程，Token 留在本机忽略的 `secrets\mineru_token.txt`，不输出。

## 主要命令（原稿含 OfficeMath 时）

    python scripts\one_click_convert.py --input "<原稿.docx>" --inspect-only
    python scripts\one_click_convert.py --input "<原稿.docx>"
    python scripts\diagnose_after_conversion.py --source "<冻结源.docx>" --working "<转换稿.docx>" --outdir "<review>"

`--inspect-only` 不创建 READY。正式转换后使用 `READY_FOR_GPT_REVIEW.json` 和 `NEXT_ACTION.json` 继续；新增正文对象之前先诊断。只有 AxMath 和普通文字时直接处理正文目标。仅 Word 交付时禁止 PDF 渲染验收命令；已有断点应恢复，不重复转换或审计。

开发测试位于 `scripts/test_*.py`，CI 不代替 Windows Word/AxMath 真实转换与 Word/WPS 界面验收。

[变更记录](CHANGELOG.md) · [开发贡献](CONTRIBUTING.md) · [安全说明](SECURITY.md) · [历史技术记录](docs/TECHNICAL_BACKGROUND.md)（仅背景，不覆盖当前规则）
