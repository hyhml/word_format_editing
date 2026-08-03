"""Deterministic patterns used only while compiling requirement sources."""

from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass(frozen=True)
class RequirementPattern:
    id: str
    category: str
    pattern: str

    def compile(self) -> re.Pattern[str]:
        return re.compile(self.pattern, re.IGNORECASE)


PATTERNS = (
    RequirementPattern("paper_a4", "page", r"(?:A\s*4|Ａ\s*4|A\s*４)(?:纸|页面|幅面)?"),
    RequirementPattern("orientation_portrait", "page", r"(?:纵向|portrait)"),
    RequirementPattern("orientation_landscape", "page", r"(?:横向|landscape)"),
    RequirementPattern("margin_top", "page", r"(?:上边距|顶部|上)[^。；;\n0-9]{0,12}(\d+(?:\.\d+)?)\s*(cm|厘米|mm|毫米|in|英寸)"),
    RequirementPattern("margin_bottom", "page", r"(?:下边距|底部|下)[^。；;\n0-9]{0,12}(\d+(?:\.\d+)?)\s*(cm|厘米|mm|毫米|in|英寸)"),
    RequirementPattern("margin_left", "page", r"(?:左边距|左)[^。；;\n0-9]{0,12}(\d+(?:\.\d+)?)\s*(cm|厘米|mm|毫米|in|英寸)"),
    RequirementPattern("margin_right", "page", r"(?:右边距|右)[^。；;\n0-9]{0,12}(\d+(?:\.\d+)?)\s*(cm|厘米|mm|毫米|in|英寸)"),
    RequirementPattern("font_size_cn", "font", r"(小初|初号|小一|一号|小二|二号|小三|三号|小四|四号|小五|五号|小六|六号|七号|八号)(?:字|号)?"),
    RequirementPattern("font_size_pt", "font", r"(\d+(?:\.\d+)?)\s*(?:pt|磅)"),
    RequirementPattern("line_spacing_multiple", "paragraph", r"(\d+(?:\.\d+)?)\s*倍(?:行距)?"),
    RequirementPattern("line_spacing_fixed", "paragraph", r"固定值?[^0-9]{0,8}(\d+(?:\.\d+)?)\s*(?:pt|磅)"),
    RequirementPattern("first_line_chars", "paragraph", r"首行缩进[^。；;\n]{0,12}?([0-9]+|一|二|两|三|四)\s*(?:个)?字符"),
    RequirementPattern("three_line_table", "table", r"(?:三线表|三线制)"),
    RequirementPattern("page_number_roman", "pagination", r"(?:页码|编页)[^。；;\n]{0,30}(?:小写罗马|罗马数字)"),
    RequirementPattern("page_number_arabic", "pagination", r"(?:页码|编页)[^。；;\n]{0,30}(?:阿拉伯数字|Arabic)"),
    RequirementPattern("page_break_before", "pagination", r"(?:另起一页|另页开始|分页开始)"),
    RequirementPattern("keep_with_next", "paragraph", r"(?:与下段同页|与下段保持同页)"),
)

PATTERN_BY_ID = {pattern.id: pattern for pattern in PATTERNS}
