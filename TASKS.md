# AI Agent Skill 开发任务

目标：把当前 Word 格式自动化原型升级为一个可复用的 AI agent skill。工具应能读取格式要求文件，复用或生成格式包，识别原始论文结构，并输出符合要求的 `.docx`。

## 总流程

```text
格式要求文件集 + 用户描述
        ↓
0. 格式文件整合与格式包复用判断
        ↓
命中已有格式包 ───────────────┐
        ↓                    │
1. 格式要求解析               │
        ↓                    │
2. 格式化脚本/引擎配置生成     │
        ↓                    │
保存新格式包                  │
        ↓                    │
原始论文.docx ────────────────┘
        ↓
3. 论文结构识别
        ↓
4. Word 格式化执行
        ↓
5. 格式合规性校验与报告
        ↓
formatted.docx
```

暂不实现“差异预览与人工微调”作为核心流程；后续可作为增强功能加入。

## 需求来源

- 当前项目原型：`generate_formatter.py`
- 武科大论文专用参考实现：`format_thesis.py`
- 用户提出的目标流程：
  - 格式要求可来自 `.txt`、`.docx`、`.pdf`
  - 格式要求先转成标准 Markdown 文档
  - 再生成针对该格式的 Python 脚本
  - 原始论文 Word 先被识别成带标题层级的 Markdown/结构文件
  - 最后生成符合格式要求的 Word 论文
- 设计约束：
  - 已生成过的格式应直接复用
  - 支持多格式要求文件组成一个“格式要求集”
  - 支持命令式规则和约束式/推导式规则
  - 结构识别失败时优先保留原样，不破坏内容
  - 生成与执行过程要有事务性、校验和错误报告

## 模块 0：格式文件整合与格式包复用判断

### 目标

判断用户提供的格式要求是否已经处理过。如果命中已有格式包，跳过模块 1 和模块 2，直接进入论文结构识别与格式化执行。

### 输入

- 一个或多个格式要求文件：`.txt`、`.docx`、`.pdf`
- 可选用户描述：学校、期刊、年份、格式名称、补充说明

### 输出

- 命中时：已有格式包路径
- 未命中时：标准化后的格式要求集描述，交给模块 1

### 数据结构

格式包目录建议：

```text
formats/
  wust_thesis_2024/
    manifest.json
    format_spec.md
    format_spec.json
    formatter.py
    source/
      source_01.pdf
      source_02.docx
```

`manifest.json` 至少包含：

```json
{
  "id": "wust_thesis_2024",
  "name": "武汉科技大学本科毕业论文格式 2024",
  "version": "0.1.0",
  "created_at": "2026-07-14",
  "source_hashes": [],
  "combined_source_hash": "",
  "keywords": ["武汉科技大学", "本科毕业论文", "2024"],
  "format_spec": "format_spec.json",
  "formatter": "formatter.py"
}
```

### 子任务

- [x] 实现格式文件哈希计算：单文件 hash 和多文件组合 hash。
- [x] 实现 `formats/` 格式包索引扫描。
- [x] 实现精确匹配：`combined_source_hash` 命中即复用。
- [x] 实现元数据匹配：学校、期刊、年份、关键词。
- [x] 预留语义相似匹配接口，先不强依赖 embedding。
- [x] 未提供用户描述且解析信心不足时，输出需要澄清的问题。

### 验收标准

- 同一组格式文件重复输入时，能稳定命中已有格式包。
- 文件顺序不同但集合相同时，组合 hash 不变。
- 未命中时能返回明确的“需要新建格式包”状态。

## 模块 1：格式要求解析

### 目标

把格式要求文件集解析成两份文件：

- `format_spec.md`：给人看的标准格式要求说明
- `format_spec.json`：给程序执行的结构化格式规范

### 输入

- 模块 0 输出的格式要求集
- 用户描述或澄清信息

### 输出

- `format_spec.md`
- `format_spec.json`

### 关键能力

- `.txt`：直接读取文本。
- `.docx`：优先用 Pandoc 或 python-docx 提取文本和表格。
- `.pdf`：优先文本提取；扫描版后续接 OCR。
- 多文件要求要能合并，冲突处要标记来源和置信度。

### `format_spec.json` 建议结构

```json
{
  "metadata": {
    "name": "",
    "institution": "",
    "document_type": "",
    "year": "",
    "sources": []
  },
  "page": {},
  "styles": {},
  "sections": {},
  "headings": [],
  "paragraphs": {},
  "tables": {},
  "figures": {},
  "equations": {},
  "references": {},
  "headers_footers": {},
  "derived_rules": [],
  "validation_rules": []
}
```

### 子任务

- [x] 定义 `format_spec.json` schema 第一版。
- [x] 实现 TXT 格式要求提取。
- [x] 实现 DOCX 格式要求提取。
- [x] 实现 PDF 文本提取。
- [x] 实现格式要求归一化到 `format_spec.md`。
- [x] 实现从 `format_spec.md` 到 `format_spec.json` 的规则抽取。
- [x] 增加 `derived_rules`，支持“比正文小一号”“逐级递减”等推导式规则。
- [x] 增加冲突记录，例如同一字段在两个来源中不一致。

### 验收标准

- 能从一份简单格式要求文本生成可读 `format_spec.md`。
- 能生成可被模块 2 使用的 `format_spec.json`。
- 对缺失字段不瞎猜，能标记 `unknown` 或提出澄清。

## 模块 2：格式化脚本/引擎配置生成

### 目标

根据 `format_spec.json` 生成可执行的格式化能力。优先发展为“通用引擎 + 配置注入”，减少每个格式包重复生成大量代码。

### 输入

- `format_spec.json`

### 输出

- 格式包中的 `formatter.py`
- 或通用 `format_engine.py` 可读取的配置

### 设计方向

```text
format_engine.py       通用执行引擎
format_spec.json       格式配置
formatter.py           薄封装，负责加载配置和特殊补丁
openxml_patches/       复杂 Word XML 补丁模块
```

### 子任务

- [x] 从当前 `generate_formatter.py` 抽出通用 engine。
- [x] 保留生成 `formatter.py` 的能力，但让它尽量只加载配置。
- [x] 将 `format_thesis.py` 中可复用逻辑拆成 OpenXML patch 候选模块：
  - 页眉页脚
  - 三线表
  - 图表题
  - 公式编号
  - 数学字体
  - 参考文献
- [x] 实现事务性执行：永远在临时目录和临时文件上工作。
- [x] 执行失败时保留原文件不变，并输出错误报告。

### 验收标准

- 用一个简单 `format_spec.json` 可以格式化 `.docx`。
- 格式化失败不会破坏原始文件。
- 特殊补丁可以按配置启用或禁用。

## 模块 3：论文结构识别

### 目标

识别原始论文 `.docx` 的逻辑结构，输出给人检查的 Markdown 和给程序执行的 JSON。

### 输入

- `raw.docx`
- 可选：格式包中的结构规则

### 输出

- `paper_structure.md`
- `paper_structure.json`

### 识别对象

- 标题层级
- 摘要和关键词
- 目录
- 正文章节
- 图和图题
- 表和表题
- 公式和公式编号
- 参考文献
- 致谢
- 附录
- 脚注、尾注、长引文等可扩展特殊块

### 子任务

- [x] 实现 `.docx` 基础段落读取。
- [x] 实现标题层级识别：基于 Word 样式、编号模式和文本模式。
- [x] 实现图题、表题识别。
- [x] 实现参考文献区段识别。
- [x] 输出 `paper_structure.md` 供人工阅读。
- [x] 输出 `paper_structure.json` 供模块 4 使用。
- [x] 对不能识别的复杂对象标记为 `preserve`，后续保留原样。

### 验收标准

- 对普通论文 Word，能识别主要章节层级。
- 识别失败不会删除或重写未知内容。
- `paper_structure.json` 能定位回原 docx 中的段落或对象。

## 模块 4：工作流门控与 formatter 启动

### 目标

检查前置模块产物是否齐全；如果缺失，返回缺失项和应回到的模块；如果齐全，调用已有 `formatter.py` 或 `format_engine.py` 完成格式化。

模块 4 不自动补齐缺失产物，不重新解析格式要求，不重新识别论文结构，也不重新实现格式化规则。

### 输入

- `raw.docx`
- 格式包目录，或显式传入 `format_spec.json` 和 `formatter.py`
- `paper_structure.json`
- 输出目录

### 输出

- `formatted.docx`
- `format_report.json`
- `format_report.md`
- `workflow_report.json`
- `workflow_report.md`

### 执行原则

- 只做门控检查和执行编排。
- 缺少前置产物时立即停止，报告应回到哪个模块。
- 前置产物齐全时，优先调用格式包中的 `formatter.py`。
- 如果没有专用 `formatter.py`，但有 `format_spec.json`，可回退调用 `format_engine.py`。
- 不自动调用模块 1、模块 2 或模块 3 补齐缺失文件。

### 子任务

- [x] 检查 `raw.docx` 是否存在。
- [x] 检查格式包或 `format_spec.json` 是否存在。
- [x] 检查 `formatter.py` 是否存在；不存在时检查是否可使用 `format_engine.py`。
- [x] 检查 `paper_structure.json` 是否存在。
- [x] 缺失任一必要产物时输出 blocked 状态和 `return_to` 模块提示。
- [x] 前置产物齐全时调用 `formatter.py` 或 `format_engine.py`。
- [x] 汇总 formatter/engine 的 `format_report.json`。
- [x] 输出 `workflow_report.json` 和 `workflow_report.md`。

### 验收标准

- 前置产物缺失时不执行格式化，并明确返回缺失项和应回到的模块。
- 前置产物齐全时能调用已有 formatter 输出 `.docx` 和格式报告。
- 工作流报告能说明使用了哪个格式包、哪个 formatter、哪个结构文件，以及最终状态。

## 模块 5：格式合规性校验与报告

### 目标

格式化完成后，用独立校验器检查关键格式是否达标。复用旧格式包时也必须执行校验，防止格式包过期或错误复用。

### 输入

- `formatted.docx`
- `format_spec.json`
- `paper_structure.json`

### 输出

- `validation_report.md`
- `validation_report.json`

### 子任务

- [ ] 校验页面大小和页边距。
- [ ] 校验正文样式。
- [ ] 校验标题样式。
- [ ] 校验表格格式。
- [ ] 校验图表题格式。
- [ ] 校验页眉页脚。
- [ ] 给出 pass/warn/fail 三级结果。

### 验收标准

- 对格式化后的文档能输出明确合规报告。
- 复用旧格式包时，如果校验失败，能提示格式包可能过时或匹配错误。

## 里程碑

### M1：格式包复用骨架

- [x] 建立 `formats/` 目录规范。
- [x] 实现格式文件 hash 和 manifest 匹配。
- [x] 能判断“复用旧格式包”还是“创建新格式包”。

### M2：基础格式规范解析

- [x] 支持 TXT/DOCX/PDF 文本抽取。
- [x] 生成 `format_spec.md` 和基础 `format_spec.json`。
- [x] 不追求复杂规则完整覆盖。

### M3：通用格式引擎

- [x] 将当前 `generate_formatter.py` 升级为通用 engine。
- [x] 支持页面、正文、标题、表格基础格式。

### M4：论文结构识别

- [x] 从 `raw.docx` 输出 `paper_structure.md/json`。
- [x] 支持标题、图题、表题、参考文献基础识别。

### M5：端到端最小闭环

- 输入一份格式要求和一份原始论文。
- 自动创建或复用格式包。
- 输出格式化 Word 和报告。

### M6：OpenXML 高级补丁

- 从 `format_thesis.py` 中拆出可复用补丁。
- 支持页眉页脚、三线表、公式编号、数学字体等复杂规则。

### M7：封装为 AI agent skill

- 编写 `SKILL.md`。
- 定义 skill 的输入、输出、工作流和失败处理。
- 提供示例任务和测试样例。

## 当前下一步

M1、M2、M3、M4（论文结构识别）和模块 4（工作流门控与 formatter 启动）已完成。下一步建议实现模块 5：格式合规性校验与报告。

模块 5 的第一批任务：

1. 读取 `formatted.docx`、`format_spec.json` 和 `paper_structure.json`。
2. 校验页面大小、页边距、正文、标题和表格基础格式。
3. 输出 `validation_report.json` 和 `validation_report.md`。
4. 用 pass/warn/fail 标记关键规则是否达标。
