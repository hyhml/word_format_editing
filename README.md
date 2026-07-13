# Word 格式自动化生成工具

这个目录提供一个两步工具链：

1. 用 JSON 描述你的 Word 格式要求。
2. 根据 JSON 自动生成一个独立的 Python 格式化脚本，再用该脚本处理原始 `.docx`。

长期目标是把它升级为一个 AI agent skill。后续开发任务见 [TASKS.md](/home/hyhml/codex/word/TASKS.md)。

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
