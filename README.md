# Word 格式自动化生成工具

这个目录提供一个 Word 格式自动化工具链：

1. 从格式要求文件创建或复用格式包。
2. 识别原始论文 `.docx` 的结构。
3. 格式化 Word 并输出格式化、工作流和校验报告。

如果已有结构化 JSON 规格，也可以继续使用旧的“生成 formatter 再处理 `.docx`”方式。

长期目标是把它升级为一个 AI agent skill。后续开发任务见 [TASKS.md](TASKS.md)。

## v0.13 格式要求编译器（Schema v2.1）

`format_compiler.py` 专门处理“格式要求识别”阶段，把 `.txt`、`.md`、`.json`、`.doc`、`.docx` 或 `.pdf` 编译成唯一、规整的 `format_spec.json`。旧版 `.doc` 需要系统安装 `antiword`，且只能可靠保留文字证据；若格式本身也承载要求，优先转换为 `.docx`：

```bash
python3 format_compiler.py compile \
  --source path/to/format.pdf \
  --source path/to/extra.docx \
  --name "某大学本科毕业论文格式 2026" \
  --description "学校规范与学院补充要求" \
  --output-dir out/format-recognition
```

输出只有三个正式产物：

- `format_spec.json`：Schema v2 的确定性执行规范。
- `recognition_report.json`：来源证据、覆盖率、冲突、保留项和校验详情。
- `recognition_report.md`：供人工检查的识别摘要。

Schema v2.1 对每个已识别属性允许四种唯一动作：`set`、`preserve`、`remove`、`conditional`。规范没有要求、无法识别或存在未解决冲突时使用 `preserve`，不会静默填入宋体、小四等常见默认值。相同目标和属性出现冲突时，正式规范只保留一个 `preserve` 结果，候选值和证据写入识别报告；不同适用条件下的值会合并为一个 `conditional` 动作，多条件同时命中时仍强制保持原格式。

v2.1 新增四类可规整能力：

- `content.template` 使用字段、字面量、空格和引导符 token 描述编号与内容模板，不把自然语言说明留给执行器。
- `section.position`、`section.relative_position`、`section.relative_to` 描述章节顺序。
- `text_span` 选择器通过父目标、捕获组或排除目标定位段内不同格式的文字片段。
- `conditional` 使用固定上下文字段和运算符表达学位类型、培养类型、语言例外及“存在内容时”等条件；条件输入未知、未命中或同时命中多个不一致分支时均默认 `preserve`。

生成 AI 分析请求：

```bash
python3 format_compiler.py ai-request \
  --source path/to/format.pdf \
  --output ai_request.json
```

AI 必须返回带来源块 ID 的候选规则。将候选规则交回编译器：

```bash
python3 format_compiler.py compile \
  --source path/to/format.pdf \
  --ai-candidates ai_candidates.json \
  --output-dir out/format-recognition
```

校验已有 v2 规范：

```bash
python3 format_spec_validator_v2.py \
  out/format-recognition/format_spec.json \
  --report out/format-recognition/schema_validation.json
```

校验失败时，报告中的 `repair_request` 可直接返回 AI 做定点修复；AI 不得补充无来源格式值。当前 Word 格式化引擎尚未适配 Schema v2，会明确拒绝执行 v2 识别产物。旧 pipeline 在执行器适配完成前继续使用旧规格，避免产生错误格式化结果。

仓库内 Skill 位于 `skills/compile-format-requirements/`。它要求 AI 对每个来源块进行分类、引用真实证据，并在最多三轮“编译—校验—定点修复”后将仍无法确定的局部规则保留原格式。

## 端到端入口：模块 0 到模块 5

`format_pipeline.py` 会串联格式包复用、格式解析、formatter 生成、论文结构识别、格式化执行和格式校验：

```bash
python3 format_pipeline.py run \
  --input raw.docx \
  --format-source path/to/format.txt \
  --format-source path/to/extra.docx \
  --description "武汉科技大学 本科毕业论文 2024" \
  --name "武汉科技大学本科毕业论文格式 2024" \
  --package-id wust_thesis_2024 \
  --formats-dir formats \
  --output-dir out
```

运行后会在 `out/` 下输出：

- `formatted.docx`
- `paper_structure.md` / `paper_structure.json` / `structure_report.json`
- `workflow_report.json` / `workflow_report.md`
- `format_report.json` / `format_report.md`
- `validation_report.json` / `validation_report.md`
- `pipeline_report.json` / `pipeline_report.md`

如果格式要求命中已有格式包，pipeline 会直接复用；否则会在 `formats/` 下创建新格式包，包含 `manifest.json`、`format_spec.md`、`format_spec.json`、`formatter.py` 和 `source/` 原始格式要求文件副本。

## 模块 0：格式包复用判断

模块 0 已实现为 `format_registry.py`。它用于判断一组格式要求文件是否已经生成过格式包，如果命中即可跳过格式解析和脚本生成。

检查格式要求文件是否命中已有格式包：

```bash
python3 format_registry.py check --formats-dir formats examples/format_spec.example.json
```

带用户描述进行元数据匹配：

```bash
python3 format_registry.py check \
  --formats-dir formats \
  --description "武汉科技大学 本科毕业论文 2024" \
  path/to/format.pdf
```

为新格式包生成 manifest 模板：

```bash
python3 format_registry.py manifest-template \
  --id wust_thesis_2024 \
  --name "武汉科技大学本科毕业论文格式 2024" \
  --keyword 武汉科技大学 \
  --keyword 本科毕业论文 \
  path/to/format.pdf
```

示例格式包位于 `formats/example_general/`。

## 模块 1：格式要求解析

模块 1 已实现为 `format_parser.py`。它把格式要求文件解析成：

- `format_spec.md`：给人检查的标准格式说明。
- `format_spec.json`：给后续模块执行的结构化格式规范。
- `parse_report.json`：解析报告，包含来源、告警、冲突和待澄清字段。

解析格式要求文件：

```bash
python3 format_parser.py parse \
  --output-dir formats/my_format \
  --description "武汉科技大学 本科毕业论文 2024" \
  path/to/format.txt path/to/extra.docx
```

当前支持：

- `.txt` / `.md` / `.json`：直接读取文本。
- `.docx`：提取段落和表格文本。
- `.pdf`：提取普通文本；扫描版 PDF 会产生告警，暂不做 OCR。

模块 1 会把无法确定的字段写入 `unknowns`，把多文件不一致的字段写入 `conflicts`，不会静默猜测或覆盖。

## 模块 2：通用格式引擎

模块 2 已实现为 `format_engine.py`。它读取 `format_spec.json` 并格式化 `.docx`：

```bash
python3 format_engine.py \
  --spec formats/example_general/format_spec.json \
  --input raw.docx \
  --output formatted.docx \
  --report format_report.json
```

`generate_formatter.py` 仍保留原入口，但现在生成的是薄封装脚本：

```bash
python3 generate_formatter.py \
  --spec formats/example_general/format_spec.json \
  --output formats/example_general/formatter.py
```

生成后的 formatter 用法：

```bash
python3 formats/example_general/formatter.py raw.docx formatted.docx \
  --report format_report.json
```

当前引擎支持页面、正文、标题和表格基础格式。`openxml_patches/` 已提供高级补丁第一版，支持三线表、页眉页脚、图表题、参考文献、数学字体和公式段落对齐；缺少规则或文档中没有适用对象时会写入 `skipped_patches`，不会中断基础格式化。公式编号当前是保守实现，复杂公式矩阵和 Word 原生 `eqArr` 重排仍需后续增强。

## 模块 3：论文结构识别

模块 3 已实现为 `paper_structure.py`。它只读取原始论文 `.docx`，不修改 Word，输出论文逻辑结构：

- `paper_structure.md`：给人检查的结构预览。
- `paper_structure.json`：给后续模块使用的结构数据。
- `structure_report.json`：识别报告和统计信息。

识别论文结构：

```bash
python3 paper_structure.py analyze \
  --input raw.docx \
  --output-md paper_structure.md \
  --output-json paper_structure.json \
  --report structure_report.json
```

当前支持标题层级、摘要、关键词、目录、图题、表题、参考文献、致谢、附录等基础识别。无法可靠识别的复杂对象会进入 `preserve`，后续模块默认保留原样。

## 模块 4：工作流门控与 formatter 启动

模块 4 已实现为 `format_workflow.py`。它检查前置产物是否齐全，并在齐全时调用已有 formatter：

```text
raw.docx + format_spec.json/formatter.py + paper_structure.json
        ↓
检查产物是否齐全
        ↓
齐全：调用 formatter.py 或 format_engine.py
缺失：返回 blocked 报告和应回到的模块
```

模块 4 不自动补齐缺失产物，不重新解析格式要求，也不重新识别论文结构。

使用格式包运行：

```bash
python3 format_workflow.py run \
  --input raw.docx \
  --format-package formats/example_general \
  --structure paper_structure.json \
  --output formatted.docx \
  --workflow-report-json workflow_report.json \
  --workflow-report-md workflow_report.md \
  --format-report-json format_report.json \
  --format-report-md format_report.md
```

也可以显式传入 `format_spec.json` 和 `formatter.py`；如果不传 `formatter.py`，模块 4 会在 `format_spec.json` 存在时回退调用 `format_engine.py`：

```bash
python3 format_workflow.py run \
  --input raw.docx \
  --spec formats/example_general/format_spec.json \
  --structure paper_structure.json \
  --output formatted.docx \
  --workflow-report-json workflow_report.json \
  --workflow-report-md workflow_report.md \
  --format-report-json format_report.json \
  --format-report-md format_report.md
```

缺少 `raw.docx`、`format_spec.json`、可用 formatter/engine 或 `paper_structure.json` 时，命令不会执行格式化，会输出 blocked 工作流报告和应回到的模块。

## 模块 5：格式合规性校验与报告

模块 5 已实现为 `format_validator.py`。它不修改 Word，只检查 `formatted.docx` 是否符合 `format_spec.json`，并输出校验报告：

```bash
python3 format_validator.py validate \
  --input formatted.docx \
  --spec formats/example_general/format_spec.json \
  --structure paper_structure.json \
  --report-json validation_report.json \
  --report-md validation_report.md
```

校验结果分为：

- `pass`：规则明确且实际格式符合要求。
- `warn`：规则未知、对象缺失或当前版本无法可靠校验。
- `fail`：规则明确但实际格式不符合要求。

当前支持页面大小、方向、页边距、正文段落、标题段落、表格单元格、图题/表题基础段落格式、页眉页脚距离，以及模块 2 OpenXML 高级补丁的第一版校验：

- 三线表边框。
- 页眉文本和页脚页码字段。
- 图题/表题缩进、对齐、字体字号。
- 参考文献对齐、悬挂缩进和 NBSP 清理。
- 数学字体设置。
- 公式段落保守对齐校验。

复杂公式右侧编号、矩阵和多行公式的 Word 原生 `eqArr` 布局仍以 `warn` 记录。

## 安装依赖

```bash
python3 -m pip install -r requirements.txt
```

## 第一步：准备格式要求

复制示例规格：

```bash
cp examples/format_spec.example.json my_format.json
```

然后修改 `my_format.json`。核心字段：

- `page`：页面大小、方向、页边距。
- `default`：普通正文格式。
- `rules`：按段落位置、开头、包含文本或正则表达式匹配特殊段落，例如标题。
- `tables`：表格内文字格式。

常用匹配方式：

```json
{"paragraph_index": 0}
{"starts_with": "第一章"}
{"contains": "摘要"}
{"regex": "^\\d+\\.\\d+"}
```

## 第二步：生成格式化脚本

```bash
python3 generate_formatter.py --spec my_format.json --output generated_formatter.py
```

## 第三步：处理 Word 文件

```bash
python3 generated_formatter.py raw.docx formatted.docx
```

也可以查看脚本里固化的格式要求：

```bash
python3 generated_formatter.py raw.docx formatted.docx --show-spec
```

## 注意

- 当前工具支持 `.docx`，不直接支持旧版 `.doc`。
- 该工具主要处理页面、段落、字体、标题规则和表格文字格式。
- 如果原始 Word 里有复杂文本框、页眉页脚、脚注、域代码或嵌入对象，可能需要扩展脚本逻辑。
