---
name: word-axmath-conversion
description: Handle math PDF/DOCX formatting, plain-text numbered exercises, non-collapsible headings, and genuine editable AxMath conversion through MinerU and Word.
---

# AxMath 数学资料排版

**每次新任务先读本文件**，再按任务读取下表中的文档。用户本次的范围、格式和文件保留要求优先；不要从旧聊天推测路径或处理方式。

## 本机数学排版：开工前先确定

以下默认标准适用于用户授权的**整册数学练习资料排版**；当前任务另有要求时以当前要求为准。只改标题、题号或其他局部时，按授权范围编辑，保留用户已调好的字体、缩进、空段和分页；仅排版不转换已有公式。完整步骤见 [练习册排版标准](docs/MATH_LAYOUT_REFERENCE.md)。

- **字体和缩进：** 文字默认楷体，正文左右各缩进 3 个字符，首行与悬挂缩进为零。用户说明已设置好或要求保留时，记录并保留原设置，不重设；图、表、标题、页眉等按实际角色保留或设计，不用空格模拟缩进。
- **题号：** 先读取实际显示值，再转换为普通文字；不用自动编号、编号域和折叠式 Heading。编号用半角数字、英文括号，编号或右括号后紧跟一个英文句点，且编号前后不额外加空格，如 `1.题干`、`例3.题干`、`(1).小问`、`*19.(1).小问`。保留原序号、星号和逻辑层级。
- **留白和同页：** 按用户给定的参考册确定各类题目和独立小题的书写空间；页首可见内容从正文区第一行开始。完整主问题干、小问、题图和图注连同所需留白要同页；物理上放不下时不得擅自降低要求，按工作流说明冲突并请用户选择处理方式。
- **标题层级：** 先盘点章、节及下级标题实际分几层，再为每层指定统一且有区别的字号、加粗和样式；同层全文一致，封面单独处理。标题尽量与其后的第一题同页。
- **换行和分页：** 整册重排时关闭全文所有段落及继承样式的“孤行控制、与下段同页、段中不分页、段前分页”，检查各 Word story；保留自然换行，用真正的 Enter 空段安排留白与换页，不添加手动分页符。实施方法见 [排版规则](docs/LAYOUT_RULES.md)。
- **文字、公式、图片：** 正文文字与相邻行内公式之间不插入空格或制表符；不改公式内部、英文单词及数学表达所需空格。题图及图号分别居中，图号紧贴图片下方；保留已有公式和图像的内容、对象及关系。
- **验收与交付：** 先盘点来源及参考 → 做涵盖多层标题、主问题号、小题、题图和不同题型的样页 → 逐题逐页检查最终 Word，核对真实页首、整题同页、每题留白和四项分页有效值 → 核对目标路径及交付副本哈希。不能用缩略图、计数或只检查页首标号替代实际验收。

## 公式转换首轮硬规则（仅在要求创建或转换公式时执行）

1. **冻结来源**：以用户指定的最新版为基础，在工作区创建不覆盖原件的冻结副本并记录 SHA-256；先清点题目、数学内容、图片与编号。普通正文和公式的标点默认保留来源；确认属于编号的标点按上面的半角规则处理，当前任务明确指定其他格式时按其要求执行。不全局替换公式或对象；冻结副本和转换输入的衔接见 [转换工作流](docs/AXMATH_CONVERSION_WORKFLOW.md)。
2. **单个字母、数字不能漏**：逐题逐段识别正文中表达数学意义的独立字母、数字，包括已有公式前后的普通文字；每次出现均登记目标。`1.设x=1，求x。` 的题号 `1.` 保留为普通文字，两个 `x` 和数值 `1` 分别是数学目标。不能仅靠一键转换、对象总数或最终审计发现遗漏。具体判断、命令和止损见 **[单字母/数字首轮处理](docs/BODY_SINGLE_CHARACTER_MATH.md)**。
3. **先建完整映射**：生成前用 `math_layout_contract.py map-template` 创建 `SOURCE_FINAL_FORMULA_MAP.json`。逐项填写 `source_formula` 的 `source_anchor`、来源 `expected_tex`；补入 PDF/图片源公式 `source_formula_visual`，以及正文数学目标 `added_body_math`（含准确 `source_text`）。模板不会自动识别正文目标，`additional_target_examples` 不是已登记目标。
4. **只用真实可编辑 AxMath**：已有 OfficeMath/OMML 用官方 `AMSMML2AM`，经来源核定的 TeX 和正文数学目标用官方 `AMSTeX2AM`。中文说明、题号和选项标签留作普通文字；完整数学表达式按原意处理，不逐字拆散。
5. **题号与层级首轮正确**：题号直接写普通文字，保留原序号、跳号和逻辑层级；使用英文半角括号及编号后的句点，编号后紧接内容。保持规定的正文左右缩进，不使用自动列表、编号域、悬挂缩进、编号制表位和 Heading 折叠身份。方法以 `docs/NUMBERING_AND_OUTLINE.md` 为准。
6. **公式保真**：大积分用 `\int`，禁止 `\smallint`；下置极限用 `\underset{...}{\lim}`。撇号、结构、上下标和运算符按冻结来源确定，不从损坏对象猜原式。
7. **一次生成，一次最终验收**：先确认目标、定位和代表样页，再按目标生成；完成后填写每个目标的终稿对象序号，所有对象恰好归属一次。最终候选只做一次完整内部 TeX 导出与一次真实 Word 逐页检查；发现问题按定位定向修复，不反复进行无证据的全书审计。

## 按任务读取

| 任务 | 必读文件 |
|---|---|
| 分页选项、图片/图号居中、编号后或公式前后空格、正文缩进 | **`docs/LAYOUT_RULES.md`**（含只读检查命令和首轮门槛） |
| 数学讲义、练习册、参考册留白、标题层级、题图及整题分页 | **`docs/MATH_LAYOUT_REFERENCE.md`**、`docs/NUMBERING_AND_OUTLINE.md` |
| 自动编号、层级错位、Word/WPS 折叠箭头 | `docs/NUMBERING_AND_OUTLINE.md` |
| PDF → 可编辑 DOCX | `docs/MINERU_API.md` |
| 公式转换、源到终稿映射及验收 | `docs/AXMATH_CONVERSION_WORKFLOW.md` |
| 正文独立字母或数字需要 AxMath | **`docs/BODY_SINGLE_CHARACTER_MATH.md`** |
| Codex 在本机使用 Word/AxMath | `docs/CODEX_DIRECT.md` |
| `NEXT_ACTION.json` 指定异常 | `docs/REPAIR_REFERENCE.md` 中对应类别 |

## 执行顺序

1. 明确本次需要转换还是只排版；仅排版时不要无故转换已有公式。用户只要 Word 或禁止导出 PDF 时，不生成交付或临时 PDF；在真实 Word 页面视图逐页验收。
2. 纯排版按 [练习册排版标准](docs/MATH_LAYOUT_REFERENCE.md) 盘点题组、参考留白和标题并做样页，保留原公式。需要创建或转换公式时才按来源建立完整公式与正文目标清单，样页覆盖单字母、数字和已有公式。自动化报告仅辅助定位，不代替按来源判断。
3. 仅在本次要求转换公式且存在 OfficeMath 时执行下列预检、转换和诊断。`--inspect-only` 只输出预检 JSON，**不会创建 READY 文件**；正式转换后读取 `READY_FOR_GPT_REVIEW.json` 的 `documents` 数组，对单份输入使用 `documents[0].frozen_source`、`documents[0].frozen_source_sha256` 和 `documents[0].working_docx`，确认冻结副本与建图来源哈希一致。诊断在新增正文对象之前执行，按 `NEXT_ACTION.json` 处理异常。

   ```powershell
   python scripts\one_click_convert.py --input "<冻结来源.docx>" --inspect-only
   python scripts\one_click_convert.py --input "<冻结来源.docx>"
   python scripts\diagnose_after_conversion.py --source "<READY.documents[0].frozen_source>" --working "<READY.documents[0].working_docx>" --outdir "<review>"
   ```

   已有运行目录使用 `--resume-run`；不要无故重头转换。若源稿仅有 AxMath 和普通文字，没有 OfficeMath，就以最新版工作副本直接处理正文目标，不运行 OfficeMath 数量诊断。
4. 本次要求公式转换且有正文单字母/数字目标时，按 `docs/BODY_SINGLE_CHARACTER_MATH.md` 的 `body_math.py inspect` → 核对并填写计划 → `body_math.py apply` 执行。计划与工作稿哈希绑定，位置必须由同段唯一上下文定位；输出的终稿对象序号写回映射。其他表达式和原图未识别内容按来源及主工作流处理，无法安全处理的目标要报告。
5. 公式转换任务在完整内部导出后，核对源到终稿映射、正文与公式语义，并逐页查看内容和排版。仅 Word 交付时按单字母/数字文档运行正文投影检查；允许 PDF 时，`strict_final_compare.py` 已包含该检查。若来源本身已有 AxMath，按主工作流先取得来源内部内容基线。仅排版任务按练习册排版标准检查，不触发公式内部导出。

## MinerU 与运行边界

- PDF → DOCX：Codex 优先使用当前对话侧边栏已登录的 MinerU 网页；预计排队超过 120 秒、连续 60 秒无进展或页面卡住时，改用本机 API：`python scripts\mineru_api.py parse --input "<原稿.pdf>" --outdir "<工作区>"`。保留 MinerU 原始输出；本机 Token 不得输出或提交。
- Word/AxMath 操作使用本机独立任务进程，不关闭或终止用户原有 Word/WPS。宏批次、断点、异常处理和内部 TeX 门禁遵循 `docs/AXMATH_CONVERSION_WORKFLOW.md`；不全局修改 OLE 大小、预览或页面数量。
- 用户要求仅保留最终版时，验收后按其指令清理本任务生成的中间交付物；保留明确要求保留的原件及用户已有文件。说明实际验证范围和未完成项。
