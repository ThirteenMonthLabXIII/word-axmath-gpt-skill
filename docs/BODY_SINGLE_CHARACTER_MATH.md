# 正文单个字母、数字：首轮处理

只在本次要求数学内容使用 AxMath 时执行。先读 [SKILL.md](../SKILL.md) 和 [转换工作流](AXMATH_CONVERSION_WORKFLOW.md)。一键转换只识别原有 OfficeMath；它不会自动把正文中的 `x` 或 `1` 识别成公式。`map-template` 也不会发现普通文字中的数学目标，必须逐项补入。

## 1. 先按来源确定每个目标

以用户指定的最新版为基础冻结来源，先取得候选清单：

```powershell
python scripts\body_math.py inspect --input "<冻结来源.docx>" --report "<工作区>\SOURCE_BODY_INSPECT.json"
```

报告按 Word 主文档段落列出 `candidates`：大小写字母、数字等每次出现都有 `candidate_id` 和上下文。`text_offset_utf16` 只标记报告中的文字位置，**不能当作 Word Range 起点**。候选还包含题号、英文单词和已有公式内的字符，必须由原稿和对象类型判断；报告不覆盖图片、其他 story 或所有特殊符号。

对照原 PDF/Word 逐题查看题干、小问、选项和表格正文，逐段处理候选并记录 `candidate_id → 普通文字数学目标 / 已有公式 / 完整表达式的一部分 / 非数学文字 / 未决`。每段无目标也记为已看过；数学目标须包含每一次出现，尤其是已有公式前后的普通文字。原图中的文字是否需要单独编辑，以当前任务要求为准；不把图中的字猜成正文。对未列出的来源位置与特殊字符仍要结合原稿补全。

按数学意义判断边界，不用全局正则替换：

| 来源片段 | 目标 |
|---|---|
| `1.设x=1，求x。` 中的题号 `1.` | 普通文字，保留原编号。 |
| 同句中两个独立的 `x` 和条件中的数值 `1` | 三个不同目标；各自记录位置，分别生成 AxMath。 |
| `A. 点A在直线l上` 中的选项标记 `A.` | 普通文字。 |
| 同句中的点 `A` 和直线 `l` | 表达数学意义，各为目标。 |

中文说明、题号、小问号、选项标签、图号、页码及普通英文单词按原意保留为文字。多位数或 `x+1` 等本来作为一个完整数学表达式时按完整表达式处理，不拆成逐字公式。已有 AxMath 保留并映射；已有 OfficeMath 由主转换链处理，不重复新增。上下文无法判断时对照原稿并记录未决项，不能猜测或宣称完成。

希腊字母同样按数学意义登记，例如原文 `α` 的 `source_text` 为 `α`，经来源核实的 `expected_tex` 为 `\alpha`。计划和正文投影会比较单字符与 TeX 的语义；不认识或不能确认的符号须标记未决，不用相似字母代替。

生成前在 `SOURCE_FINAL_FORMULA_MAP.json` 的 `targets` 中，为每个普通文字数学目标增加 `added_body_math` 条目；`additional_target_examples` 只是示例，不会自动成为目标：

```json
{
  "target_id": "BODY-001",
  "kind": "added_body_math",
  "source_anchor": "第1题，‘设’后、‘=1’前的x",
  "source_text": "x",
  "expected_tex": "x",
  "final_ordinals": []
}
```

`source_text` 必须是冻结来源中的原始文字，`expected_tex` 从来源确定；重复出现分别登记。对扫描资料中未被识别为文字的内容用 `source_formula_visual`，记录 PDF 页码和区域，不虚构 `source_text`。先核对整份目标清单，再开始转换；清单漏掉的目标不会被映射检查自动发现。

## 2. 在当前工作稿中精确转换

若有 OfficeMath，先按主转换链完成其批量转换、诊断和必要定向修复，再使用交接文件里的 `working_docx`。诊断必须在新增正文对象之前运行，因为它按原有公式数量和正文比较。若最新版已有 AxMath 且无 OfficeMath，直接以其工作副本作为输入，不运行只处理 OfficeMath 的一键转换或其数量诊断。

在隔离工作目录运行下面的命令，获取**当前工作稿**的 Word 段落和输入 SHA-256，用来定位上一步已确定的目标。若工作稿与步骤 1 的冻结来源路径和哈希完全相同，直接复用 `SOURCE_BODY_INSPECT.json`，无需重复读取。

```powershell
python scripts\body_math.py inspect --input "<当前工作稿.docx>" --report "<工作区>\BODY_INSPECT.json"
```

根据报告中的 `paragraph_index`、段落原文和逐题目标清单，填写 `BODY_PLAN.json`。`before` 和 `after` 是目标在**当前工作稿同一段内紧邻的原文**；三者拼接后必须在该段恰好出现一次。重复时延长上下文；不要填写旧版段落号、页码或估计的字符偏移。每个 `target_id` 必须与公式映射中的目标一致。

```json
{
  "schema": "axmath-body-math-plan/v1",
  "input_sha256": "填写BODY_INSPECT.json中的input_sha256",
  "targets": [
    {
      "target_id": "BODY-001",
      "paragraph_index": 1,
      "before": "设",
      "source_text": "x",
      "after": "=1",
      "expected_tex": "x"
    }
  ]
}
```

确认全部目标和输入哈希后运行一次；输出路径和报告路径必须是工作区内尚不存在的新文件：

```powershell
python scripts\body_math.py apply --input "<当前工作稿.docx>" --plan "<工作区>\BODY_PLAN.json" --output "<工作区>\body_converted.docx" --report "<工作区>\BODY_APPLY.json"
```

该入口在独立后台 Word 中以官方 `AMSTeX2AM` 生成真实 AxMath，先删除经定位核实的原字符，再在同一位置插入对象，按文末向前处理；此前的对象会使 Word 内部位置变化，因此不能自行用文本长度推算 Word 位置。完全相同且大小写一致的 TeX 复用公式样本，但每处生成独立的 AxMath 对象。它检查每段未选中文字、段落/对象数量、保存后的结果及输入哈希。

加入正文对象前，先为工作稿中原有 AxMath 的每个来源目标记录其已核实的工作稿序号（可在映射条目增加 `baseline_ordinals`）。完成后按下列方式填写最终序号：

1. 将 `BODY_APPLY.json` 每行 `final_ordinal` 写入同 `target_id` 的正文目标。
2. 取 `1..axmath_after`，删除上述所有新对象序号。剩余序列必须有 `axmath_before` 项，按顺序分别对应操作前的第 1、2、… 个 AxMath；据此把每个来源目标的 `baseline_ordinals` 换成 `final_ordinals`。例如操作前有 3 个对象，新对象最终为 2、5，原 1、2、3 依次对应最终 1、3、4。
3. 若来源目标与操作前对象尚未完整对应，或数量不符，停止并定位；不要按旧序号猜。多次拆分的来源目标逐个转换其已核实的工作稿序号。

本入口把整份计划作为一个批次，末尾才保存并验证，没有批内自动续跑。失败时停止，检查报告和保留的 `.partial.docx`；部分文件可能只是输入副本或已保存结果，不能当完成稿。确认原因并核实输入哈希后，才可从未改动的输入、经复核的计划和新的输出路径重试；未解决的异常如实报告。

本入口只替换主文档中**一个 BMP 字母或数字的普通文字**；上下文不能跨对象或控制字符，目标不能处于域、超链接或内容控件中，且输入不能有残余 OfficeMath 或修订。多字表达式、扫描图中未识别的字符及其他 story（如页眉、脚注、文本框）需按来源另行处理。无法安全处理时记录未完成项；不要把主文档报告当成覆盖所有位置。

## 3. 一次最终验收

填写最终稿路径、哈希、总数和全部 `final_ordinals`；每个对象恰有一个来源。单字目标的内部 TeX 必须与该位置的来源字符一致。仅 Word 交付时，在主工作流的 `math_layout_contract.py validate` 之后，再运行正文投影检查：

```powershell
python scripts\body_math_text_contract.py --source "<映射中的冻结来源.docx>" --final "<最终候选.docx>" --formula-map "<工作区>\SOURCE_FINAL_FORMULA_MAP.json" --out "<工作区>\BODY_TEXT_VALIDATION.json"
```

两个报告均需 `passed=true`。正文投影将已映射对象还原为对应 `source_text` 后比较逐段文字，检查漏删原字符、位置错乱和无关文字变化；它不负责发现未登记目标，也不检查对象内部语义。允许 PDF 时，`strict_final_compare.py` 已集成该检查，不再单独重复运行。最后一次逐页对照来源，并在 Word 中确认代表性对象可由 AxMath 编辑；发现问题只定位相关目标修复。正文以外未处理项、无法确认的来源及实际检查范围应如实报告。
