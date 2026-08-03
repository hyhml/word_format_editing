"""Normalize common Chinese format values into canonical units and enums."""

from __future__ import annotations

import re
from typing import Any


CN_FONT_SIZES_PT = {
    "初号": 42,
    "小初": 36,
    "一号": 26,
    "小一": 24,
    "二号": 22,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "小四": 12,
    "五号": 10.5,
    "小五": 9,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5,
}

CHINESE_NUMBERS = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6}


def normalize_font_size(value: str | int | float) -> float | int | None:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return value if value > 0 else None
    text = str(value).strip()
    # Prefer the more specific "小X" names when overlapping text such as
    # "小四号" also contains the shorter semantic token "四号".
    for name in sorted(CN_FONT_SIZES_PT, key=lambda item: (-len(item), 0 if item.startswith("小") else 1)):
        if name in text:
            return CN_FONT_SIZES_PT[name]
    match = re.search(r"(\d+(?:\.\d+)?)\s*(?:pt|磅)", text, re.IGNORECASE)
    return float(match.group(1)) if match else None


def normalize_length_cm(value: str | int | float, default_unit: str | None = None) -> float | None:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) if default_unit == "cm" else None
    match = re.search(r"(-?\d+(?:\.\d+)?)\s*(cm|厘米|mm|毫米|in|英寸)", str(value), re.IGNORECASE)
    if not match:
        return None
    number = float(match.group(1))
    unit = match.group(2).lower()
    if unit in {"cm", "厘米"}:
        return number
    if unit in {"mm", "毫米"}:
        return number / 10
    return number * 2.54


def normalize_length_pt(value: str | int | float) -> float | None:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    match = re.search(r"(-?\d+(?:\.\d+)?)\s*(?:pt|磅)", str(value), re.IGNORECASE)
    return float(match.group(1)) if match else None


def normalize_character_count(value: str | int | float) -> float | None:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value)
    match = re.search(r"(\d+(?:\.\d+)?)\s*(?:个)?字符", text)
    if match:
        return float(match.group(1))
    for character, number in CHINESE_NUMBERS.items():
        if re.search(fr"{character}\s*(?:个)?字符", text):
            return float(number)
    return None


def normalize_alignment(text: str) -> str | None:
    mappings = (("两端对齐", "justify"), ("分散对齐", "distributed"), ("居中", "center"), ("右对齐", "right"), ("左对齐", "left"), ("顶格", "left"))
    return next((value for phrase, value in mappings if phrase in text), None)


def normalize_line_spacing(text: str) -> tuple[str, float] | None:
    multiple = re.search(r"(\d+(?:\.\d+)?)\s*倍(?:行距)?", text)
    if multiple:
        return "multiple", float(multiple.group(1))
    fixed = re.search(r"固定值?[^0-9]{0,8}(\d+(?:\.\d+)?)\s*(?:pt|磅)", text, re.IGNORECASE)
    if fixed:
        return "fixed", float(fixed.group(1))
    at_least = re.search(r"最小值?[^0-9]{0,8}(\d+(?:\.\d+)?)\s*(?:pt|磅)", text, re.IGNORECASE)
    if at_least:
        return "at_least", float(at_least.group(1))
    if "单倍行距" in text:
        return "single", 1.0
    return None
