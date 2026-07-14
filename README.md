# Word 格式自动化生成工具

这个目录提供一个两步工具链：

1. 用 JSON 描述你的 Word 格式要求。
2. 根据 JSON 自动生成一个独立的 Python 格式化脚本，再用该脚本处理原始 `.docx`。

长期目标是把它升级为一个 AI agent skill。后续开发任务见 [TASKS.md](TASKS.md)。

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

当前引擎支持页面、正文、标题和表格基础格式。高级 OpenXML 补丁已有注册框架，尚未实现的补丁会写入 `skipped_patches`，不会中断基础格式化。

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
