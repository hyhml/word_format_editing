#!/usr/bin/env python3
"""Compile requirement sources into canonical format_spec schema v2."""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Iterable

from format_normalization import (
    CN_FONT_SIZES_PT,
    normalize_alignment,
    normalize_character_count,
    normalize_font_size,
    normalize_length_cm,
    normalize_line_spacing,
)
from format_ontology import PROPERTY_DEFINITIONS, SCHEMA_VERSION, TARGETS, ontology_summary
from format_spec_validator_v2 import validate_spec
from requirement_patterns import PATTERN_BY_ID
from requirement_source import RequirementSource, SourceBlock, extract_requirement_sources
from selector_patterns import SELECTOR_PATTERNS


EAST_ASIA_FONTS = ("宋体", "黑体", "楷体", "仿宋", "隶书", "微软雅黑", "方正小标宋")
LATIN_FONTS = ("Times New Roman", "Arial", "Calibri", "Cambria")


@dataclass(frozen=True)
class Candidate:
    target: str
    property: str
    value: Any
    evidence_id: str
    unit: str | None = None
    method: str = "deterministic"
    confidence: float = 0.8


def _set_action(candidate: Candidate, evidence_ids: list[str]) -> dict[str, Any]:
    action: dict[str, Any] = {
        "action": "set",
        "value": candidate.value,
        "evidence_ids": evidence_ids,
        "method": candidate.method,
        "confidence": candidate.confidence,
    }
    if candidate.unit:
        action["unit"] = candidate.unit
    return action


def empty_spec(sources: list[RequirementSource], name: str, description: str) -> dict[str, Any]:
    institution = None
    document_type = None
    year = None
    metadata_text = f"{name} {description} " + " ".join(
        [*(Path(source.path).stem for source in sources), *(block.text for source in sources for block in source.blocks)]
    )
    institution_match = re.search(r"([\u4e00-\u9fff]{2,30}(?:大学|学院|学校|研究院|期刊|出版社))", metadata_text)
    year_match = re.search(r"(20\d{2})", metadata_text)
    if institution_match:
        institution = institution_match.group(1)
    if "本科" in metadata_text and "论文" in metadata_text:
        document_type = "本科毕业论文"
    elif "硕士" in metadata_text and "论文" in metadata_text:
        document_type = "硕士论文"
    elif "期刊" in metadata_text or "投稿" in metadata_text:
        document_type = "期刊论文"
    if year_match:
        year = year_match.group(1)
    return {
        "schema_version": SCHEMA_VERSION,
        "metadata": {
            "name": name or description or Path(sources[0].path).stem,
            "institution": institution,
            "document_type": document_type,
            "year": year,
            "description": description,
            "sources": [source.metadata() for source in sources],
        },
        "policies": {
            "unspecified_property_action": "preserve",
            "conflict_action": "preserve",
            "selector_priority": ["explicit_location", "structure", "style", "regex", "body_fallback"],
        },
        "document": {"properties": {}},
        "targets": {},
        "selectors": {},
        "recognition": {"status": "success", "compiler": "format_compiler_v2", "requires_review": False},
    }


def target_for_block(block: SourceBlock) -> str | None:
    text = block.text
    heading_patterns = (
        ("heading.level_1", r"一级标题|章标题"),
        ("heading.level_2", r"二级标题|节标题"),
        ("heading.level_3", r"三级标题"),
        ("heading.level_4", r"四级标题"),
        ("heading.level_5", r"五级标题"),
        ("heading.level_6", r"六级标题"),
    )
    for target, pattern in heading_patterns:
        if re.search(pattern, text):
            return target
    if re.search(r"中文摘要标题|摘要标题", text):
        return "abstract.zh.heading"
    if re.search(r"英文摘要标题|Abstract标题", text, re.IGNORECASE):
        return "abstract.en.heading"
    if re.search(r"中文摘要|摘要正文", text):
        return "abstract.zh.body"
    if re.search(r"英文摘要|Abstract正文", text, re.IGNORECASE):
        return "abstract.en.body"
    if re.search(r"中文关键词", text):
        return "keywords.zh"
    if re.search(r"英文关键词|Keywords", text, re.IGNORECASE):
        return "keywords.en"
    if re.search(r"封面", text):
        return "cover"
    if re.search(r"题名页|扉页", text):
        return "title_page"
    if re.search(r"原创性声明", text):
        return "originality_statement"
    if re.search(r"授权(?:使用)?声明", text):
        return "authorization_statement"
    if re.search(r"目录标题", text):
        return "table_of_contents.heading"
    if re.search(r"一级目录", text):
        return "table_of_contents.level_1"
    if re.search(r"二级目录", text):
        return "table_of_contents.level_2"
    if re.search(r"三级目录", text):
        return "table_of_contents.level_3"
    if re.search(r"图目录", text):
        return "list_of_figures"
    if re.search(r"表目录", text):
        return "list_of_tables"
    if re.search(r"图题|图标题|图名|图注", text):
        return "figure.caption"
    if re.search(r"表题|表标题|表名", text):
        return "table.caption"
    if re.search(r"表格|三线表|表内", text):
        return "table.body"
    if re.search(r"公式编号", text):
        return "equation.number"
    if "公式" in text:
        return "equation"
    if re.search(r"参考文献(?:条目|正文|内容)", text):
        return "references.entry"
    if re.search(r"参考文献标题", text):
        return "references.heading"
    if re.search(r"致谢正文", text):
        return "acknowledgements.body"
    if re.search(r"致谢标题", text):
        return "acknowledgements.heading"
    if re.search(r"附录正文", text):
        return "appendix.body"
    if re.search(r"附录标题", text):
        return "appendix.heading"
    if re.search(r"脚注", text):
        return "footnote"
    if re.search(r"尾注", text):
        return "endnote"
    if re.search(r"正文|主体文字|普通段落", text):
        return "body.paragraph"
    return None


def has_multiple_format_scopes(text: str) -> bool:
    """Return true when one block assigns different formats to sub-objects."""
    conflicting_markers = (
        ("附录标题", "附录内容"),
        ("论文题目", "其余"),
        ("表头", "表内"),
        ("表头", "表注"),
        ("图标题", "图注"),
        ("页眉", "其余"),
    )
    return any(first in text and second in text for first, second in conflicting_markers)


def add_candidate(items: list[Candidate], target: str, property_name: str, value: Any, block: SourceBlock, unit: str | None = None, confidence: float = 0.8) -> None:
    if value is None:
        return
    items.append(Candidate(target, property_name, value, block.id, unit, "deterministic", confidence))


def deterministic_candidates(sources: list[RequirementSource]) -> list[Candidate]:
    candidates: list[Candidate] = []
    for source in sources:
        for block in source.blocks:
            text = re.sub(r"\s+", " ", block.text).strip()
            target = target_for_block(block)

            if PATTERN_BY_ID["paper_a4"].compile().search(text):
                add_candidate(candidates, "document", "page.paper_size", "A4", block, confidence=0.98)
            if PATTERN_BY_ID["orientation_portrait"].compile().search(text):
                add_candidate(candidates, "document", "page.orientation", "portrait", block, confidence=0.9)
            if PATTERN_BY_ID["orientation_landscape"].compile().search(text):
                add_candidate(candidates, "document", "page.orientation", "landscape", block, confidence=0.9)

            handled_margins: set[str] = set()
            paired_vertical = re.search(
                r"上[、,，和及]下[^。；;\n]{0,16}?分别[^0-9]{0,8}"
                r"(\d+(?:\.\d+)?)\s*(cm|厘米|mm|毫米|in|英寸)\s*(?:和|、|,|，)\s*"
                r"(\d+(?:\.\d+)?)\s*(cm|厘米|mm|毫米|in|英寸)",
                text,
                re.IGNORECASE,
            )
            if paired_vertical:
                top = normalize_length_cm("".join(paired_vertical.groups()[:2]))
                bottom = normalize_length_cm("".join(paired_vertical.groups()[2:]))
                add_candidate(candidates, "document", "page.margin_top_cm", top, block, "cm", 0.96)
                add_candidate(candidates, "document", "page.margin_bottom_cm", bottom, block, "cm", 0.96)
                handled_margins.update({"page.margin_top_cm", "page.margin_bottom_cm"})

            margin_patterns = {
                "page.margin_top_cm": PATTERN_BY_ID["margin_top"].pattern,
                "page.margin_bottom_cm": PATTERN_BY_ID["margin_bottom"].pattern,
                "page.margin_left_cm": PATTERN_BY_ID["margin_left"].pattern,
                "page.margin_right_cm": PATTERN_BY_ID["margin_right"].pattern,
                "page.header_distance_cm": r"页眉[^0-9]{0,12}(\d+(?:\.\d+)?)\s*(cm|厘米|mm|毫米|in|英寸)",
                "page.footer_distance_cm": r"页脚[^0-9]{0,12}(\d+(?:\.\d+)?)\s*(cm|厘米|mm|毫米|in|英寸)",
            }
            for property_name, pattern in margin_patterns.items():
                if property_name in handled_margins:
                    continue
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    value = normalize_length_cm("".join(match.groups()))
                    add_candidate(candidates, "document", property_name, value, block, "cm", 0.92)

            if target and not has_multiple_format_scopes(text):
                for font in EAST_ASIA_FONTS:
                    if font in text:
                        add_candidate(candidates, target, "font.east_asia", font, block, confidence=0.9)
                        break
                for font in LATIN_FONTS:
                    if font.lower() in text.lower():
                        add_candidate(candidates, target, "font.latin", font, block, confidence=0.9)
                        break
                size = normalize_font_size(text)
                if size:
                    add_candidate(candidates, target, "font.size_pt", size, block, "pt", 0.9)
                alignment = normalize_alignment(text)
                if alignment:
                    add_candidate(candidates, target, "paragraph.alignment", alignment, block, confidence=0.88)
                spacing = normalize_line_spacing(text)
                if spacing:
                    spacing_type, spacing_value = spacing
                    add_candidate(candidates, target, "paragraph.line_spacing.type", spacing_type, block, confidence=0.9)
                    unit = "multiple" if spacing_type in {"multiple", "single"} else "pt"
                    add_candidate(candidates, target, "paragraph.line_spacing.value", spacing_value, block, unit, 0.9)
                indent = normalize_character_count(text) if "首行缩进" in text else None
                if indent is not None:
                    add_candidate(candidates, target, "paragraph.first_line_indent_chars", indent, block, "character", 0.92)
                if re.search(r"不加粗|非粗体", text):
                    add_candidate(candidates, target, "font.bold", False, block, confidence=0.9)
                elif re.search(r"加粗|粗体", text):
                    add_candidate(candidates, target, "font.bold", True, block, confidence=0.9)
                if "斜体" in text and "不斜体" not in text:
                    add_candidate(candidates, target, "font.italic", True, block, confidence=0.85)

            if PATTERN_BY_ID["three_line_table"].compile().search(text):
                add_candidate(candidates, "table", "table.style", "three_line", block, confidence=0.95)
            if re.search(r"表题[^。；;\n]{0,30}(?:表上|上方|表前)", text):
                add_candidate(candidates, "table.caption", "caption.position", "above", block, confidence=0.9)
            if re.search(r"图题[^。；;\n]{0,30}(?:图下|下方|图后)", text):
                add_candidate(candidates, "figure.caption", "caption.position", "below", block, confidence=0.9)
            if re.search(r"悬挂缩进", text) and "参考文献" in text:
                length = normalize_length_cm(text)
                if length is not None:
                    add_candidate(candidates, "references.entry", "references.hanging_indent_cm", length, block, "cm", 0.85)
            if re.search(r"(?:小写罗马|罗马数字)", text) and re.search(r"页码|编页", text):
                add_candidate(candidates, "document.pagination", "numbering.format", "lower_roman", block, confidence=0.85)
            if "阿拉伯数字" in text and re.search(r"页码|编页", text):
                add_candidate(candidates, "document.pagination", "numbering.format", "arabic", block, confidence=0.85)
    return candidates


BLOCK_CLASSIFICATIONS = {"requirement", "explanation", "example", "irrelevant", "unresolved"}
UNRESOLVED_REASONS = {"ambiguous", "missing_dependency", "unsupported_property", "source_incomplete"}


def load_ai_candidates(path: Path | None) -> tuple[list[Candidate], list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    if path is None:
        return [], [], [], []
    data = json.loads(path.read_text(encoding="utf-8"))
    raw_candidates = data.get("candidates", []) if isinstance(data, dict) else []
    candidates = []
    classifications = []
    unresolved_items = []
    errors = []
    for index, item in enumerate(raw_candidates):
        try:
            target = str(item["target"])
            property_name = str(item["property"])
            evidence_ids = item["evidence_ids"]
            if target != "document" and target not in TARGETS:
                raise ValueError(f"未知target: {target}")
            if property_name not in PROPERTY_DEFINITIONS:
                raise ValueError(f"未知property: {property_name}")
            if not isinstance(evidence_ids, list) or not evidence_ids:
                raise ValueError("evidence_ids必须是非空数组")
            for evidence_id in evidence_ids:
                candidates.append(
                    Candidate(
                        target=target,
                        property=property_name,
                        value=item["value"],
                        evidence_id=str(evidence_id),
                        unit=item.get("unit"),
                        method="ai",
                        confidence=float(item.get("confidence", 0.7)),
                    )
                )
        except (KeyError, TypeError, ValueError) as exc:
            errors.append({"index": index, "message": str(exc), "candidate": item})
    for index, item in enumerate(data.get("block_classifications", []) if isinstance(data, dict) else []):
        try:
            evidence_id = str(item["evidence_id"])
            classification = str(item["classification"])
            if classification not in BLOCK_CLASSIFICATIONS:
                raise ValueError(f"未知classification: {classification}")
            classifications.append(
                {"evidence_id": evidence_id, "classification": classification, "notes": str(item.get("notes", ""))}
            )
        except (KeyError, TypeError, ValueError) as exc:
            errors.append({"index": index, "message": str(exc), "block_classification": item})
    for index, item in enumerate(data.get("unresolved_items", []) if isinstance(data, dict) else []):
        try:
            evidence_id = str(item["evidence_id"])
            reason = str(item["reason"])
            text = str(item["text"])
            if reason not in UNRESOLVED_REASONS:
                raise ValueError(f"未知unresolved reason: {reason}")
            unresolved_items.append(
                {
                    "evidence_id": evidence_id,
                    "text": text,
                    "reason": reason,
                    "notes": str(item.get("notes", "")),
                }
            )
        except (KeyError, TypeError, ValueError) as exc:
            errors.append({"index": index, "message": str(exc), "unresolved_item": item})
    return candidates, classifications, unresolved_items, errors


def merge_candidates(spec: dict[str, Any], candidates: Iterable[Candidate]) -> list[dict[str, Any]]:
    grouped: dict[tuple[str, str], list[Candidate]] = {}
    for candidate in candidates:
        grouped.setdefault((candidate.target, candidate.property), []).append(candidate)
    conflicts = []
    for (target, property_name), items in sorted(grouped.items()):
        unique: dict[str, list[Candidate]] = {}
        for item in items:
            value = float(item.value) if isinstance(item.value, (int, float)) and not isinstance(item.value, bool) else item.value
            key = json.dumps({"value": value, "unit": item.unit}, ensure_ascii=False, sort_keys=True)
            unique.setdefault(key, []).append(item)
        rule = spec["document"] if target == "document" else spec["targets"].setdefault(target, {"properties": {}})
        if len(unique) == 1:
            same_items = next(iter(unique.values()))
            selected = sorted(same_items, key=lambda item: item.confidence, reverse=True)[0]
            evidence_ids = sorted({item.evidence_id for item in same_items})
            rule["properties"][property_name] = _set_action(selected, evidence_ids)
            continue
        conflict = {
            "target": target,
            "property": property_name,
            "candidates": [asdict(item) for item in items],
            "resolution": "preserve",
        }
        conflicts.append(conflict)
        rule["properties"][property_name] = {
            "action": "preserve",
            "reason": "unresolved_conflict",
            "evidence_ids": sorted({item.evidence_id for item in items}),
        }
    return conflicts


def install_selectors(spec: dict[str, Any]) -> None:
    for target in sorted(spec["targets"]):
        definition = SELECTOR_PATTERNS.get(target)
        if not definition:
            continue
        level_match = re.fullmatch(r"heading\.level_(\d+)", target)
        selector: dict[str, Any] = {
            "structure_type": "heading" if level_match else target,
            "fallback_regex": list(definition["patterns"]),
            "priority": int(definition["priority"]),
        }
        if level_match:
            selector["structure_level"] = int(level_match.group(1))
            selector["style_names"] = [f"Heading {level_match.group(1)}", f"标题 {level_match.group(1)}"]
        spec["selectors"][target] = selector


def source_coverage(
    sources: list[RequirementSource],
    candidates: list[Candidate],
    classifications: list[dict[str, Any]],
    unresolved_items: list[dict[str, Any]],
) -> dict[str, Any]:
    all_blocks = [block for source in sources for block in source.blocks]
    mapped_ids = {candidate.evidence_id for candidate in candidates}
    classification_by_id = {item["evidence_id"]: item["classification"] for item in classifications}
    covered_ids = mapped_ids | set(classification_by_id)
    mapped = [block.id for block in all_blocks if block.id in mapped_ids]
    unmapped = [block.id for block in all_blocks if block.id not in covered_ids]
    unresolved_item_ids = {item["evidence_id"] for item in unresolved_items}
    unresolved = [
        block.id
        for block in all_blocks
        if classification_by_id.get(block.id) == "unresolved"
        or (classification_by_id.get(block.id) == "requirement" and block.id not in mapped_ids)
        or block.id in unresolved_item_ids
    ]
    ratio = len(covered_ids & {block.id for block in all_blocks}) / len(all_blocks) if all_blocks else 0.0
    return {
        "total_blocks": len(all_blocks),
        "mapped_blocks": len(mapped),
        "classified_blocks": len([block for block in all_blocks if block.id in classification_by_id]),
        "unmapped_blocks": unmapped,
        "unresolved_blocks": unresolved,
        "mapped_ratio": round(ratio, 4),
    }


def render_report(report: dict[str, Any]) -> str:
    lines = [
        "# 格式要求识别报告",
        "",
        f"- 状态：{report['status']}",
        f"- Schema：{report['schema_version']}",
        f"- 来源文件：{len(report['sources'])}",
        f"- 候选规则：{report['candidate_count']}",
        f"- 冲突：{len(report['conflicts'])}",
        f"- 局部未解决要求：{len(report.get('unresolved_items', []))}",
        f"- 映射覆盖率：{report['coverage']['mapped_ratio']:.2%}",
        "",
        "## 来源警告",
        "",
    ]
    if report.get("source_warnings"):
        lines.extend(f"- {warning}" for warning in report["source_warnings"])
    else:
        lines.append("- 无")
    lines.extend([
        "",
        "## 校验",
        "",
    ])
    if report["validation"]["errors"]:
        lines.extend(f"- 失败：{item['path']} — {item['message']}" for item in report["validation"]["errors"])
    else:
        lines.append("- Schema、数值、正则和语义校验通过。")
    lines.extend(["", "## 冲突", ""])
    if report["conflicts"]:
        lines.extend(f"- {item['target']}.{item['property']}：保持原格式" for item in report["conflicts"])
    else:
        lines.append("- 无")
    lines.extend(["", "## 局部未解决要求", ""])
    if report.get("unresolved_items"):
        for item in report["unresolved_items"]:
            notes = f"；{item['notes']}" if item.get("notes") else ""
            lines.append(f"- {item['evidence_id']}：{item['text']}（{item['reason']}{notes}）")
    else:
        lines.append("- 无")
    lines.extend(["", "## 未映射来源块", ""])
    unmapped = report["coverage"]["unmapped_blocks"]
    lines.extend(f"- {item}" for item in unmapped) if unmapped else lines.append("- 无")
    lines.append("")
    return "\n".join(lines)


def compile_sources(source_paths: list[Path], name: str = "", description: str = "", ai_candidates_path: Path | None = None) -> tuple[dict[str, Any], dict[str, Any], list[RequirementSource]]:
    sources = extract_requirement_sources(source_paths)
    spec = empty_spec(sources, name, description)
    deterministic = deterministic_candidates(sources)
    ai_candidates, classifications, unresolved_items, ai_errors = load_ai_candidates(ai_candidates_path)
    valid_evidence_ids = {block.id for source in sources for block in source.blocks}
    accepted_ai_candidates = []
    for candidate in ai_candidates:
        if candidate.evidence_id not in valid_evidence_ids:
            ai_errors.append(
                {
                    "message": f"AI候选引用不存在的证据ID: {candidate.evidence_id}",
                    "candidate": asdict(candidate),
                }
            )
        else:
            accepted_ai_candidates.append(candidate)
    ai_candidates = accepted_ai_candidates
    accepted_classifications = []
    seen_classifications: dict[str, str] = {}
    for item in classifications:
        evidence_id = item["evidence_id"]
        if evidence_id not in valid_evidence_ids:
            ai_errors.append({"message": f"AI分类引用不存在的证据ID: {evidence_id}", "block_classification": item})
            continue
        previous = seen_classifications.get(evidence_id)
        if previous is not None and previous != item["classification"]:
            ai_errors.append({"message": f"同一来源块存在冲突分类: {evidence_id}", "block_classification": item})
            continue
        seen_classifications[evidence_id] = item["classification"]
        accepted_classifications.append(item)
    classifications = accepted_classifications
    accepted_unresolved_items = []
    for item in unresolved_items:
        if item["evidence_id"] not in valid_evidence_ids:
            ai_errors.append(
                {"message": f"未解决项引用不存在的证据ID: {item['evidence_id']}", "unresolved_item": item}
            )
        else:
            accepted_unresolved_items.append(item)
    unresolved_items = accepted_unresolved_items
    candidates = [*deterministic, *ai_candidates]
    conflicts = merge_candidates(spec, candidates)
    install_selectors(spec)
    validation = validate_spec(spec)
    source_warnings = [warning for source in sources for warning in source.warnings]
    requires_review = bool(conflicts or ai_errors or unresolved_items or validation["warnings"] or source_warnings)
    coverage = source_coverage(sources, candidates, classifications, unresolved_items)
    requires_review = requires_review or bool(coverage["unmapped_blocks"] or coverage["unresolved_blocks"])
    status = "blocked" if validation["errors"] else ("warn" if requires_review else "success")
    spec["recognition"] = {"status": status, "compiler": "format_compiler_v2", "requires_review": requires_review}
    validation = validate_spec(spec)
    finalized_rules = []
    preserved_properties = []
    for target, rule in [("document", spec["document"]), *spec["targets"].items()]:
        for property_name, action in rule["properties"].items():
            item = {"target": target, "property": property_name, **action}
            if action["action"] == "preserve":
                preserved_properties.append(item)
            else:
                finalized_rules.append(item)
    evidence_catalog = {
        block.id: {
            "source_id": block.source_id,
            "type": block.type,
            "page": block.page,
            "paragraph_index": block.paragraph_index,
            "table_index": block.table_index,
            "row": block.row,
            "column": block.column,
            "text": block.text,
            "style": block.style,
        }
        for source in sources
        for block in source.blocks
    }
    report = {
        "status": status,
        "schema_version": SCHEMA_VERSION,
        "sources": [source.metadata() for source in sources],
        "source_warnings": source_warnings,
        "candidate_count": len(candidates),
        "deterministic_candidate_count": len(deterministic),
        "ai_candidate_count": len(ai_candidates),
        "ai_candidate_errors": ai_errors,
        "block_classifications": classifications,
        "unresolved_items": unresolved_items,
        "recognized_rules": finalized_rules,
        "preserved_properties": preserved_properties,
        "conflicts": conflicts,
        "coverage": coverage,
        "evidence_catalog": evidence_catalog,
        "validation": validation,
        "validation_rounds": [{"round": 1, "status": validation["status"], "errors": validation["errors"]}],
    }
    return spec, report, sources


def write_outputs(output_dir: Path, spec: dict[str, Any], report: dict[str, Any], sources: list[RequirementSource]) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "format_spec.json").write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
    (output_dir / "recognition_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    (output_dir / "recognition_report.md").write_text(render_report(report), encoding="utf-8")


def ai_request(sources: list[RequirementSource]) -> dict[str, Any]:
    return {
        "instruction": "依据来源证据补充确定性编译器未覆盖的格式规则；不得猜测缺省值。",
        "output_contract": {
            "candidates": [
                {
                    "target": "固定对象词表中的值或document",
                    "property": "固定属性词表中的值",
                    "value": "规整后的具体值",
                    "unit": "属性需要时填写标准单位",
                    "evidence_ids": ["至少一个来源块ID"],
                    "confidence": "0到1",
                }
            ],
            "block_classifications": [
                {
                    "evidence_id": "每个未由候选规则使用的来源块ID",
                    "classification": "requirement|explanation|example|irrelevant|unresolved",
                    "notes": "简短理由"
                }
            ],
            "unresolved_items": [
                {
                    "evidence_id": "部分规则无法规整时的来源块ID",
                    "text": "无法规整的原文片段",
                    "reason": "ambiguous|missing_dependency|unsupported_property|source_incomplete",
                    "notes": "为何必须保持原格式或等待扩展"
                }
            ]
        },
        "ontology": ontology_summary(),
        "sources": [source.metadata() for source in sources],
        "blocks": [block.to_dict() for source in sources for block in source.blocks],
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compile requirement files into canonical format_spec schema v2.")
    subparsers = parser.add_subparsers(dest="command", required=True)
    compile_parser = subparsers.add_parser("compile")
    compile_parser.add_argument("--source", type=Path, action="append", required=True)
    compile_parser.add_argument("--output-dir", type=Path, required=True)
    compile_parser.add_argument("--name", default="")
    compile_parser.add_argument("--description", default="")
    compile_parser.add_argument("--ai-candidates", type=Path)
    request_parser = subparsers.add_parser("ai-request")
    request_parser.add_argument("--source", type=Path, action="append", required=True)
    request_parser.add_argument("--output", type=Path, required=True)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.command == "compile":
        spec, report, sources = compile_sources(args.source, args.name, args.description, args.ai_candidates)
        write_outputs(args.output_dir, spec, report, sources)
        print(f"已生成: {args.output_dir / 'format_spec.json'}")
        print(f"已生成: {args.output_dir / 'recognition_report.json'}")
        print(f"已生成: {args.output_dir / 'recognition_report.md'}")
        if report["status"] == "blocked":
            raise SystemExit(1)
        return
    sources = extract_requirement_sources(args.source)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(ai_request(sources), ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已生成: {args.output}")


if __name__ == "__main__":
    main()
