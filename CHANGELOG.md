# 版本说明

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
