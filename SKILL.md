---
name: word-axmath-conversion
description: Handle math PDF/DOCX formatting, plain-text numbered exercises, non-collapsible headings, and genuine editable AxMath conversion through MinerU and Word.
---

# AxMath 数学资料排版

**每次新任务先读本文件**，再按任务读取下表中的文档。用户本次的范围、格式和文件保留要求优先；不要从旧聊天推测路径或处理方式。

## 本机数学排版：开工前先确定

以下是十三月数学资料的默认格式，用户当前任务另有明确要求时按当前要求执行。适用于新建、转换和修复，**章节标题也在内**。实施与检查步骤见 [排版规则与首轮门槛](docs/LAYOUT_RULES.md)。

- **换行和分页：** 全文所有段落及其继承样式的“孤行控制、与下段同页、段中不分页、段前分页”均关闭。包含标题、题目、小问、图和图号，检查表格、文本框、页眉页脚及脚注等 story；旧设置也须关闭。普通自然换行保留，用真正的 Enter 空段安排留白，不添加手动分页符或靠这些选项排版。
- **题号与缩进：** 实际显示题号先记录，再转成普通文字；禁自动列表、编号域和 Heading 折叠身份。数字用半角阿拉伯数字，题号句点及编号括号用英文半角，如 `1.题干`、`(1)小问`。保留原序号、跳号、星号及逻辑层级。正文、章/节标题和小问从各自正文区或单元格左边界起，首行、悬挂和层级缩进均为零；题号后不插入空格、制表符或占位符。
- **文字、公式、图片：** 正文文字与相邻公式之间不插入空格或制表符；不改公式内部、英文单词及数学表达所需空格。题图及其对应图号分别居中，图号紧贴图片下方；图片和公式不得混为一类，也不能一键把所有段落改为左对齐。
- **首轮门槛：** 冻结来源、盘点题号/题图/公式/样式 → 在生成前定好角色和布局 → 做覆盖标题、跨层级编号、行内公式、图片及图号的代表样页 → 只读检查与实际页面检查通过后再批量生成 → 对最终候选复查全部适用项。结构检测不能替代真实 Word 页面和公式可编辑性检查；每次修改使受影响的验收证据失效。

## 首轮生成硬规则

1. **冻结来源**：以用户指定的最新版为基础，在工作区创建不覆盖原件的冻结副本并记录 SHA-256；先清点题目、数学内容、图片与编号。普通正文和公式的标点默认保留来源；确认属于编号的标点按上面的半角规则处理，当前任务明确指定其他格式时按其要求执行。不全局替换公式或对象；冻结副本和转换输入的衔接见 [转换工作流](docs/AXMATH_CONVERSION_WORKFLOW.md)。
2. **单个字母、数字不能漏**：逐题逐段识别正文中表达数学意义的独立字母、数字，包括已有公式前后的普通文字；每次出现均登记目标。`1.设x=1，求x。` 的题号 `1.` 保留为普通文字，两个 `x` 和数值 `1` 分别是数学目标。不能仅靠一键转换、对象总数或最终审计发现遗漏。具体判断、命令和止损见 **[单字母/数字首轮处理](docs/BODY_SINGLE_CHARACTER_MATH.md)**。
3. **先建完整映射**：生成前用 `math_layout_contract.py map-template` 创建 `SOURCE_FINAL_FORMULA_MAP.json`。逐项填写 `source_formula` 的 `source_anchor`、来源 `expected_tex`；补入 PDF/图片源公式 `source_formula_visual`，以及正文数学目标 `added_body_math`（含准确 `source_text`）。模板不会自动识别正文目标，`additional_target_examples` 不是已登记目标。
4. **只用真实可编辑 AxMath**：已有 OfficeMath/OMML 用官方 `AMSMML2AM`，经来源核定的 TeX 和正文数学目标用官方 `AMSTeX2AM`。中文说明、题号和选项标签留作普通文字；完整数学表达式按原意处理，不逐字拆散。
5. **题号与层级首轮正确**：题号直接写普通文字，保留原序号、跳号和逻辑层级；使用英文半角编号标点，编号后紧接内容，各级题文顶格。自动列表、编号域、悬挂缩进、制表位和 Heading 折叠身份均不使用。方法以 `docs/NUMBERING_AND_OUTLINE.md` 为准。
6. **公式保真**：大积分用 `\int`，禁止 `\smallint`；下置极限用 `\underset{...}{\lim}`。撇号、结构、上下标和运算符按冻结来源确定，不从损坏对象猜原式。
7. **一次生成，一次最终验收**：先确认目标、定位和代表样页，再按目标生成；完成后填写每个目标的终稿对象序号，所有对象恰好归属一次。最终候选只做一次完整内部 TeX 导出与一次真实 Word 逐页检查；发现问题按定位定向修复，不反复进行无证据的全书审计。

## 按任务读取

| 任务 | 必读文件 |
|---|---|
| 分页选项、图片/图号居中、编号后或公式前后空格、正文缩进 | **`docs/LAYOUT_RULES.md`**（含只读检查命令和首轮门槛） |
| 数学讲义、练习册、题图、分页留白 | `docs/MATH_LAYOUT_REFERENCE.md`、`docs/NUMBERING_AND_OUTLINE.md` |
| 自动编号、层级错位、Word/WPS 折叠箭头 | `docs/NUMBERING_AND_OUTLINE.md` |
| PDF → 可编辑 DOCX | `docs/MINERU_API.md` |
| 公式转换、源到终稿映射及验收 | `docs/AXMATH_CONVERSION_WORKFLOW.md` |
| 正文独立字母或数字需要 AxMath | **`docs/BODY_SINGLE_CHARACTER_MATH.md`** |
| Codex 在本机使用 Word/AxMath | `docs/CODEX_DIRECT.md` |
| `NEXT_ACTION.json` 指定异常 | `docs/REPAIR_REFERENCE.md` 中对应类别 |

## 执行顺序

1. 明确本次需要转换还是只排版；仅排版时不要无故转换已有公式。用户只要 Word 或禁止导出 PDF 时，不生成交付或临时 PDF；在真实 Word 页面视图逐页验收。
2. 按来源建立完整公式与正文目标清单，检查题号结构；用少量覆盖单字母、数字、已有公式及排版的样页验证后再制作全文。自动化报告仅辅助定位，不代替按来源判断哪些文字是数学内容。
3. 存在 OfficeMath 时执行下列预检、转换和诊断。`--inspect-only` 只输出预检 JSON，**不会创建 READY 文件**；正式转换后读取 `READY_FOR_GPT_REVIEW.json` 的 `documents` 数组，对单份输入使用 `documents[0].frozen_source`、`documents[0].frozen_source_sha256` 和 `documents[0].working_docx`，确认冻结副本与建图来源哈希一致。诊断在新增正文对象之前执行，按 `NEXT_ACTION.json` 处理异常。

   ```powershell
   python scripts\one_click_convert.py --input "<冻结来源.docx>" --inspect-only
   python scripts\one_click_convert.py --input "<冻结来源.docx>"
   python scripts\diagnose_after_conversion.py --source "<READY.documents[0].frozen_source>" --working "<READY.documents[0].working_docx>" --outdir "<review>"
   ```

   已有运行目录使用 `--resume-run`；不要无故重头转换。若源稿仅有 AxMath 和普通文字，没有 OfficeMath，就以最新版工作副本直接处理正文目标，不运行 OfficeMath 数量诊断。
4. 有正文单字母/数字目标时，按 `docs/BODY_SINGLE_CHARACTER_MATH.md` 的 `body_math.py inspect` → 核对并填写计划 → `body_math.py apply` 执行。计划与工作稿哈希绑定，位置必须由同段唯一上下文定位；输出的终稿对象序号写回映射。其他表达式和原图未识别内容按来源及主工作流处理，无法安全处理的目标要报告。
5. 完整内部导出后，核对源到终稿映射、正文与公式语义，并逐页查看内容和排版。仅 Word 交付时按单字母/数字文档运行正文投影检查；允许 PDF 时，`strict_final_compare.py` 已包含该检查。若来源本身已有 AxMath，按主工作流先取得来源内部内容基线。

## MinerU 与运行边界

- PDF → DOCX：Codex 优先使用当前对话侧边栏已登录的 MinerU 网页；预计排队超过 120 秒、连续 60 秒无进展或页面卡住时，改用本机 API：`python scripts\mineru_api.py parse --input "<原稿.pdf>" --outdir "<工作区>"`。保留 MinerU 原始输出；本机 Token 不得输出或提交。
- Word/AxMath 操作使用本机独立任务进程，不关闭或终止用户原有 Word/WPS。宏批次、断点、异常处理和内部 TeX 门禁遵循 `docs/AXMATH_CONVERSION_WORKFLOW.md`；不全局修改 OLE 大小、预览或页面数量。
- 用户要求仅保留最终版时，验收后按其指令清理本任务生成的中间交付物；保留明确要求保留的原件及用户已有文件。说明实际验证范围和未完成项。
