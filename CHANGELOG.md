# 版本说明

## v0.12.0（开发中）- 2026-08-03

### 新增

- 新增格式要求编译入口 `format_compiler.py`，独立生成 Schema v2 的唯一、规整 `format_spec.json`。
- 新增 `schemas/format_spec.schema.json`、论文对象词表和格式属性词表。
- 统一使用 `set`、`preserve`、`remove` 表达属性动作；没有要求、无法识别或冲突未解决时保持原格式。
- 新增统一来源块模型，保留 TXT/Markdown 行号、DOCX 段落样式和表格位置、PDF 页码等证据。
- 新增旧版 Word `.doc` 文本提取支持；通过可选的 `antiword` 读取并明确报告样式信息损失。
- 新增中文字号、长度单位、磅值、字符缩进和行距规整。
- 新增格式要求抽取正则注册表与论文对象定位正则注册表。
- 新增 Schema、数值、单位、正则、对象和属性验证器，并在失败时生成机器可读的 AI 定点修复请求。
- 新增 AI 候选导入协议；所有候选必须引用真实来源块 ID。
- 新增来源块全量分类协议，区分格式要求、解释、示例、无关内容和未解决内容，避免把说明文字误计为遗漏。
- 新增局部未解决片段协议；同一来源块中可同时保留已规整规则和当前属性词表尚不支持的要求。
- 根据华东理工大学真实规范扩充封面/扉页子对象、中英文图表题、分区页码、奇偶页页眉、双面打印、关键词数量与分隔符等对象和属性。
- 新增识别报告，记录覆盖率、候选规则、冲突、保留项、来源证据和验证轮次。
- 新增仓库内 `compile-format-requirements` Skill，约束 AI 证据使用、缺省保留策略和最多三轮定点修复。

### 兼容性

- Schema v2 暂不交给现有 Word 格式化引擎执行；旧引擎会明确拒绝 v2，避免误格式化。
- 旧版模块 0–5 pipeline 保持现有规格和行为，后续版本再适配 Schema v2。

## v0.11.0 - 2026-07-14

### 新增

- 完成模块 5 高级 OpenXML 规则校验第一版。
- 支持校验三线表 OpenXML 边框：表格顶线/底线、内部线、表头单元格底线。
- 支持校验页眉页脚 OpenXML：header/footer relationship、section reference、页眉文本和页脚 `PAGE` 字段。
- 支持校验图题/表题补丁结果：对齐、字体字号、首行缩进和 keep-with-next 可读状态。
- 支持校验参考文献补丁结果：对齐、悬挂缩进和 NBSP 清理。
- 支持校验数学字体：`settings.xml` 中的 `m:mathFont` 和已有数学 run 字体。
- 支持公式编号保守校验：检测公式/编号段落对齐，并对复杂右侧编号 `eqArr` 强校验缺失给出 `warn`。
- 新增高级校验测试，覆盖高级规则 pass、破坏三线表 fail、破坏页眉文本 fail。

### 调整

- 图题/表题校验期望首行缩进为 0，和模块 2 的 `captions` 补丁行为保持一致。
- Markdown 校验报告新增 `caption`、`references` 和 `equations` 分组。

### 验证

- `python3 -m unittest discover -s tests -v`
- `python3 -m py_compile format_registry.py format_parser.py format_engine.py generate_formatter.py paper_structure.py format_workflow.py format_validator.py format_pipeline.py format_thesis.py openxml_patches/__init__.py tests/test_format_validator.py`

## v0.10.0 - 2026-07-14

### 新增

- 完成模块 2 OpenXML 高级补丁第一版：`openxml_patches/` 不再只是占位注册框架。
- 支持按 `openxml_patches` 显式配置执行补丁，也支持从 `format_spec.json` 自动推断三线表、页眉页脚、图表题、参考文献、数学字体和公式编号补丁。
- 新增三线表补丁：设置表格顶线/底线、首行单元格底线，并移除左右和内部边框。
- 新增页眉页脚补丁：支持简单页眉文本和居中页码字段。
- 新增图表题补丁：支持图题/表题段落居中、字体字号、去首行缩进和 keep-with-next。
- 新增参考文献补丁：支持参考文献段落对齐、悬挂缩进和 NBSP 清理。
- 新增数学字体补丁：写入 `settings.xml` 数学字体，并设置已有数学 run 字体。
- 新增公式编号保守补丁：对检测到的公式或编号段落执行居中对齐；复杂 `eqArr` 重排留待后续增强。
- 新增 `tests/test_openxml_patches.py`，覆盖三线表、页眉页脚、图表题、参考文献、数学字体和公式段落对齐。

### 调整

- `format_engine.py` 会把 patch 内部非致命错误写入格式化报告的 `errors`，但不破坏已完成的基础格式化输出。

### 验证

- `python3 -m unittest discover -s tests -v`
- `python3 -m py_compile format_registry.py format_parser.py format_engine.py generate_formatter.py paper_structure.py format_workflow.py format_validator.py format_pipeline.py format_thesis.py openxml_patches/__init__.py tests/test_openxml_patches.py`

## v0.9.0 - 2026-07-14

### 新增

- 完成 M5 端到端最小闭环：新增 `format_pipeline.py`。
- 支持一个命令串联模块 0 到模块 5：格式包复用/创建、格式解析、formatter 生成、论文结构识别、格式化执行和格式校验。
- 新格式要求未命中已有格式包时，会自动创建格式包目录，写入 `manifest.json`、`format_spec.md`、`format_spec.json`、`formatter.py` 和 `source/` 原始来源副本。
- 支持输出 `pipeline_report.json` 与 `pipeline_report.md`，汇总格式包动作、工作流状态、校验状态和全部产物路径。
- 新增端到端 pipeline 单元测试和 CLI 测试。

### 修复

- 修正模块 1 新式 spec 下标题首行缩进被正文 `first_line_indent_chars` 覆盖的问题；明确 `first_line_indent_cm` 优先于字符缩进。

### 验证

- `python3 -m unittest discover -s tests -v`
- `python3 -m py_compile format_registry.py format_parser.py format_engine.py generate_formatter.py paper_structure.py format_workflow.py format_validator.py format_pipeline.py format_thesis.py openxml_patches/__init__.py`

## v0.8.0 - 2026-07-14

### 新增

- 完成模块 5 的第一版实现：`format_validator.py`。
- 支持读取 `formatted.docx`、`format_spec.json` 和 `paper_structure.json`，输出 `validation_report.json` 与 `validation_report.md`。
- 支持 `pass`、`warn`、`fail`、`failed` 状态分级。
- 支持校验页面大小、方向、页边距、正文段落、标题段落、表格单元格和图题/表题基础段落格式。
- 支持页眉页脚距离校验；复杂页眉页脚内容先以 `warn` 记录。
- 兼容旧版 `default/rules/tables` spec 和模块 1 新版 `body/headings/tables` spec。
- 新增模块 5 单元测试。

### 验证

- `PYTHONPATH=/tmp/word_format_deps python3 -m unittest discover -s tests -v`
- `PYTHONPATH=/tmp/word_format_deps python3 -m py_compile format_registry.py format_parser.py format_engine.py generate_formatter.py paper_structure.py format_workflow.py format_validator.py format_thesis.py openxml_patches/__init__.py`

## v0.7.0 - 2026-07-14

### 新增

- 完成模块 4 的第一版实现：`format_workflow.py`。
- 支持检查 `raw.docx`、格式包或 `format_spec.json`、可用 formatter/engine、`paper_structure.json` 和输出目录。
- 前置产物缺失时返回 blocked 工作流报告，并用 `return_to` 指明应回到 `input`、`module_1`、`module_2` 或 `module_3`。
- 前置产物齐全时优先调用格式包或显式传入的 `formatter.py`。
- 无专用 `formatter.py` 但存在 `format_spec.json` 时，回退调用 `format_engine.py`。
- 支持输出 `workflow_report.json`、`workflow_report.md`、`format_report.json` 和 `format_report.md`。
- 新增模块 4 单元测试。

### 验证

- `PYTHONPATH=/tmp/word_format_deps python3 -m unittest discover -s tests -v`
- `PYTHONPATH=/tmp/word_format_deps python3 -m py_compile format_registry.py format_parser.py format_engine.py generate_formatter.py paper_structure.py format_workflow.py format_thesis.py openxml_patches/__init__.py`

## v0.6.1 - 2026-07-14

### 调整

- 修正模块 4 的任务定义：模块 4 是“工作流门控与 formatter 启动器”，不是重新实现格式化规则的执行器。
- 明确模块 4 不自动补齐模块 1、模块 2 或模块 3 的缺失产物。
- 明确前置产物缺失时应输出 blocked 工作流报告，并提示应回到哪个模块。
- 明确前置产物齐全时优先调用已有 `formatter.py`，必要时回退调用 `format_engine.py`。

## v0.6.0 - 2026-07-14

### 新增

- 完成模块 3 的第一版实现：`paper_structure.py`。
- 支持从原始 `.docx` 读取段落、表格和基础对象信息。
- 支持基于 Word 样式、编号模式和文本模式识别标题层级。
- 支持识别摘要、关键词、目录、参考文献、致谢和附录等特殊区段。
- 支持识别图题、表题，并为表题建立近似表格索引关联。
- 支持输出 `paper_structure.md`、`paper_structure.json` 和 `structure_report.json`。
- 对绘图、公式和未分类段落写入 `preserve`，供后续模块保留原样。
- 新增模块 3 单元测试。

### 验证

- `PYTHONPATH=/tmp/word_format_deps python3 -m unittest discover -s tests -v`
- `PYTHONPATH=/tmp/word_format_deps python3 -m py_compile paper_structure.py`

## v0.5.0 - 2026-07-14

### 新增

- 完成模块 2 的第一版实现：`format_engine.py`。
- 将 Word 格式化能力抽为通用引擎，支持旧版示例 spec 和模块 1 生成的新式 `format_spec.json`。
- `generate_formatter.py` 改为生成薄封装 formatter，由封装脚本加载 `format_spec.json` 并调用通用引擎。
- 新增 `openxml_patches/` patch 注册框架，预留页眉页脚、三线表、图表题、公式编号、数学字体、参考文献等高级补丁。
- 支持事务性执行：在临时 docx 上处理，成功后写出；失败时不创建或覆盖输出文件。
- 支持 `format_report.json`，记录 applied、skipped、skipped_patches 和 errors。
- 新增模块 2 单元测试。

### 验证

- `PYTHONPATH=/tmp/word_format_deps python3 -m unittest discover -s tests -v`
- `PYTHONPATH=/tmp/word_format_deps python3 -m py_compile format_registry.py format_parser.py format_engine.py generate_formatter.py format_thesis.py`

## v0.4.0 - 2026-07-14

### 新增

- 完成模块 1 的第一版实现：`format_parser.py`。
- 支持把 `.txt`、`.md`、`.json`、`.docx`、可提取文本的 `.pdf` 格式要求文件解析为 `format_spec.md`、`format_spec.json` 和 `parse_report.json`。
- 定义第一版 `format_spec.json` 输出结构，覆盖 metadata、page、body、headings、tables、figures、equations、references、headers_footers、derived_rules、conflicts、unknowns 和 validation_rules。
- 支持 TXT 多编码读取，DOCX 段落/表格提取，PDF 文本提取和扫描版告警。
- 支持常见页面、正文、标题、图题、表题、三线表、公式、参考文献和页眉页脚规则的确定性抽取。
- 支持推导式规则记录，例如“比正文小一号”和“标题逐级递减”。
- 支持多文件字段冲突记录，避免静默覆盖。
- 支持缺失关键字段进入 `unknowns`，避免凭空猜测。
- 新增模块 1 单元测试。

### 验证

- `python3 -m unittest discover -s tests -v`
- `python3 -m py_compile format_registry.py format_parser.py generate_formatter.py format_thesis.py`
- `python3 format_parser.py parse --output-dir /tmp/format_parser_demo --description "通用公文格式示例" examples/format_spec.example.json`

## v0.3.0 - 2026-07-14

### 新增

- 完成模块 0 的代码设计与第一版实现：`format_registry.py`。
- 支持格式要求文件 SHA-256 指纹计算。
- 支持多文件格式要求集的顺序无关组合 hash。
- 支持扫描 `formats/` 目录下的格式包 manifest。
- 支持基于 `combined_source_hash` 的精确复用匹配。
- 支持基于描述、文件名和 manifest 关键词的元数据匹配。
- 预留语义相似匹配接口，后续可接入 embedding。
- 未命中且信息不足时返回格式澄清问题。
- 新增 `formats/example_general/` 示例格式包。
- 新增模块 0 单元测试。

### 验证

- `python3 -m unittest discover -s tests -v`
- `python3 format_registry.py check --formats-dir formats examples/format_spec.example.json`

## v0.2.0 - 2026-07-14

### 新增

- 新增 AI agent skill 开发任务文档 `TASKS.md`，明确从格式要求解析、格式包复用、论文结构识别到 Word 格式化执行的完整模块设计。
- 新增武科大论文格式化参考实现 `format_thesis.py`，作为后续拆分 OpenXML 高级补丁和论文格式化模块的样例。
- 在 `README.md` 中加入长期 skill 化开发任务入口。

### 调整

- 将后续开发路线明确拆成 M1 到 M7 里程碑，优先从格式包复用骨架开始。
- 明确把 Markdown 作为人类可读和结构辅助产物，而不是唯一保真中间格式。

## v0.1.0 - 2026-07-13

首个可用版本，提供从格式要求到 Word 格式化脚本的基础自动化流程。

### 新增

- 支持用 JSON 描述 Word 格式要求。
- 支持根据格式要求生成独立的 Python 格式化脚本。
- 支持处理 `.docx` 文件并输出格式化后的 `.docx`。
- 支持页面大小、页面方向、页边距设置。
- 支持正文默认字体、字号、加粗、斜体、对齐、行距、首行缩进和段前段后间距。
- 支持按段落位置、开头文本、包含文本和正则表达式匹配标题或特殊段落。
- 支持表格内文字格式和表头行加粗。
- 提供示例格式规格和基础使用说明。

### 限制

- 当前只直接支持 `.docx`，旧版 `.doc` 需要先转换。
- 暂未覆盖复杂文本框、页眉页脚、脚注、域代码和嵌入对象。
- 当前还不是最终的 Codex/AI agent skill 形态，后续会在此基础上封装成 skill。
