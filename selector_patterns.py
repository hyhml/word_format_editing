"""Reusable fallback regex selectors that may enter canonical format specs."""

from __future__ import annotations


SELECTOR_PATTERNS: dict[str, dict[str, object]] = {
    "heading.level_1": {
        "patterns": [r"^第[一二三四五六七八九十百]+章(?:\s|$)", r"^\d+\s+\S"],
        "positive": ["第一章 绪论", "1 绪论"],
        "negative": ["第一节 研究背景", "1.1 研究背景"],
        "priority": 80,
    },
    "heading.level_2": {
        "patterns": [r"^第[一二三四五六七八九十百]+节(?:\s|$)", r"^\d+\.\d+(?!\.\d)\s*\S"],
        "positive": ["第一节 背景", "1.1 研究背景"],
        "negative": ["第一章 绪论", "1.1.1 研究方法"],
        "priority": 75,
    },
    "heading.level_3": {
        "patterns": [r"^\d+\.\d+\.\d+\s*\S"],
        "positive": ["1.1.1 研究方法"],
        "negative": ["1.1 研究背景"],
        "priority": 70,
    },
    "abstract.zh.heading": {
        "patterns": [r"^(?:摘\s*要|中文摘要)$"],
        "positive": ["摘要", "摘 要"],
        "negative": ["摘要正文"],
        "priority": 90,
    },
    "abstract.en.heading": {
        "patterns": [r"^(?:Abstract|ABSTRACT)$"],
        "positive": ["Abstract", "ABSTRACT"],
        "negative": ["Abstract text"],
        "priority": 90,
    },
    "keywords.zh": {
        "patterns": [r"^关键词\s*[：:]"],
        "positive": ["关键词：格式识别；论文"],
        "negative": ["本文关键词包括"],
        "priority": 85,
    },
    "keywords.en": {
        "patterns": [r"^(?:Key\s*words|Keywords)\s*[：:]"],
        "positive": ["Keywords: formatting; thesis"],
        "negative": ["The keywords are"],
        "priority": 85,
    },
    "figure.caption": {
        "patterns": [r"^图\s*\d+(?:[-—.]\d+)*\s*\S", r"^(?:Figure|Fig\.)\s*\d+(?:[-—.]\d+)*\s*\S"],
        "positive": ["图 2-1 系统结构", "Figure 2.1 System"],
        "negative": ["如图所示"],
        "priority": 75,
    },
    "table.caption": {
        "patterns": [r"^表\s*\d+(?:[-—.]\d+)*\s*\S", r"^Table\s*\d+(?:[-—.]\d+)*\s*\S"],
        "positive": ["表 2-1 实验结果", "Table 2.1 Results"],
        "negative": ["下表说明"],
        "priority": 75,
    },
    "references.heading": {
        "patterns": [r"^参考文献$", r"^References$"],
        "positive": ["参考文献", "References"],
        "negative": ["参考文献综述"],
        "priority": 90,
    },
    "references.entry": {
        "patterns": [r"^\[\d+\]\s*\S", r"^\d+[.、]\s*\S"],
        "positive": ["[1] 张三. 论文题名", "1. Zhang. Title"],
        "negative": ["第一章 绪论"],
        "priority": 65,
    },
    "appendix.heading": {
        "patterns": [r"^附录\s*[A-Z一二三四五六七八九十\d]*(?:\s|$)", r"^Appendix\s*[A-Z\d]*(?:\s|$)"],
        "positive": ["附录 A 数据", "Appendix A Data"],
        "negative": ["参见附录A"],
        "priority": 85,
    },
    "acknowledgements.heading": {
        "patterns": [r"^(?:致谢|谢辞)$"],
        "positive": ["致谢"],
        "negative": ["致谢词正文"],
        "priority": 90,
    },
}
