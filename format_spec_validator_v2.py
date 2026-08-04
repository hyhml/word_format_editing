#!/usr/bin/env python3
"""Validate canonical format_spec schema v2 and emit repair feedback."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Iterable

from jsonschema import Draft202012Validator

from format_ontology import (
    CONDITION_FIELD_DEFINITIONS,
    CONDITION_FIELDS,
    CONDITION_OPERATORS,
    PROPERTY_DEFINITIONS,
    PROPERTY_NAMES,
    SCHEMA_VERSION,
    TARGETS,
    TEMPLATE_FIELDS,
)


SCHEMA_PATH = Path(__file__).resolve().parent / "schemas" / "format_spec.schema.json"
UNIT_ALIASES = {
    "pt": {"pt"},
    "cm": {"cm"},
    "percent": {"percent"},
    "character": {"character"},
    "multiple_or_pt": {"multiple", "pt"},
    "line": {"line"},
}


def load_schema() -> dict[str, Any]:
    return json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))


def error_item(error_type: str, path: Iterable[Any], message: str, received: Any = None) -> dict[str, Any]:
    item = {"error_type": error_type, "path": ".".join(str(part) for part in path), "message": message}
    if received is not None:
        item["received"] = received
    return item


def validate_template(value: Any, path: list[str]) -> list[dict[str, Any]]:
    errors = []
    if not isinstance(value, dict) or set(value) != {"segments"} or not isinstance(value.get("segments"), list) or not value["segments"]:
        return [error_item("invalid_template", path, "内容模板必须只包含非空 segments 数组", value)]
    for index, segment in enumerate(value["segments"]):
        segment_path = [*path, "segments", str(index)]
        if not isinstance(segment, dict):
            errors.append(error_item("invalid_template_segment", segment_path, "模板片段必须是对象", segment))
            continue
        kind = segment.get("kind")
        if kind == "field":
            field = segment.get("field")
            if set(segment) != {"kind", "field"} or not isinstance(field, str) or field not in TEMPLATE_FIELDS:
                errors.append(error_item("invalid_template_segment", segment_path, "field 片段包含未知字段或多余属性", segment))
        elif kind == "literal":
            text = segment.get("text")
            if set(segment) != {"kind", "text"} or not isinstance(text, str) or not text or len(text) > 50:
                errors.append(error_item("invalid_template_segment", segment_path, "literal 片段需要1至50字符的 text", segment))
        elif kind == "spacer":
            count = segment.get("count")
            if (
                set(segment) != {"kind", "count", "unit"}
                or not isinstance(count, int)
                or isinstance(count, bool)
                or not 1 <= count <= 20
                or segment.get("unit") not in {"character", "space"}
            ):
                errors.append(error_item("invalid_template_segment", segment_path, "spacer 需要1至20的 count 和 character/space单位", segment))
        elif kind == "leader":
            character = segment.get("character")
            if set(segment) != {"kind", "character"} or not isinstance(character, str) or len(character) != 1:
                errors.append(error_item("invalid_template_segment", segment_path, "leader 需要单个 character", segment))
        elif kind == "tab":
            if set(segment) != {"kind", "alignment"} or segment.get("alignment") not in {"left", "center", "right"}:
                errors.append(error_item("invalid_template_segment", segment_path, "tab 需要 left/center/right alignment", segment))
        else:
            errors.append(error_item("invalid_template_segment", segment_path, "未知模板片段类型", kind))
    return errors


def validate_condition(condition: Any, path: list[str]) -> list[dict[str, Any]]:
    errors = []
    if not isinstance(condition, dict):
        return [error_item("invalid_condition", path, "条件必须是对象", condition)]
    logical_keys = [key for key in ("all", "any", "not") if key in condition]
    if logical_keys:
        if len(logical_keys) != 1 or len(condition) != 1:
            return [error_item("invalid_condition", path, "逻辑条件只能包含 all、any 或 not 之一", condition)]
        key = logical_keys[0]
        children = condition[key]
        if key == "not":
            return validate_condition(children, [*path, key])
        if not isinstance(children, list) or not children:
            return [error_item("invalid_condition", [*path, key], f"{key} 必须是非空条件数组", children)]
        for index, child in enumerate(children):
            errors.extend(validate_condition(child, [*path, key, str(index)]))
        return errors

    field = condition.get("field")
    operator = condition.get("operator")
    allowed_keys = {"field", "operator", "target", "value"}
    if (
        set(condition) - allowed_keys
        or not isinstance(field, str)
        or field not in CONDITION_FIELDS
        or not isinstance(operator, str)
        or operator not in CONDITION_OPERATORS
    ):
        return [error_item("invalid_condition", path, "条件字段、运算符或属性无效", condition)]
    definition = CONDITION_FIELD_DEFINITIONS[field]
    target = condition.get("target")
    if definition.get("requires_target"):
        if not isinstance(target, str) or target not in TARGETS:
            errors.append(error_item("invalid_condition_target", [*path, "target"], "目标状态条件必须引用已知 target", target))
    elif "target" in condition:
        errors.append(error_item("unexpected_condition_target", [*path, "target"], f"{field} 不接受 target", target))
    if operator == "exists":
        if "value" in condition:
            errors.append(error_item("unexpected_condition_value", [*path, "value"], "exists 条件不接受 value", condition.get("value")))
        return errors
    if "value" not in condition:
        errors.append(error_item("missing_condition_value", [*path, "value"], f"{operator} 条件必须提供 value"))
        return errors
    value = condition["value"]
    values = value if operator in {"in", "not_in"} else [value]
    if operator in {"in", "not_in"} and (not isinstance(value, list) or not value):
        errors.append(error_item("invalid_condition_value", [*path, "value"], f"{operator} 的 value 必须是非空数组", value))
        return errors
    expected_type = definition["type"]
    for item in values:
        valid = True
        if expected_type == "boolean":
            valid = isinstance(item, bool)
        elif expected_type == "string":
            valid = isinstance(item, str) and bool(item.strip())
        elif expected_type == "enum":
            valid = isinstance(item, str) and item in definition["values"]
        if not valid:
            errors.append(error_item("invalid_condition_value", [*path, "value"], f"{field} 的条件值无效", item))
    return errors


def validate_property_value(name: str, value: Any, unit: Any, has_unit: bool, path: list[str]) -> list[dict[str, Any]]:
    errors = []
    definition = PROPERTY_DEFINITIONS.get(name)
    if definition is None:
        return [error_item("unknown_property", path, f"未知格式属性: {name}", name)]
    value_type = definition["type"]
    valid_type = True
    if value_type == "string":
        valid_type = isinstance(value, str) and bool(value.strip())
    elif value_type == "boolean":
        valid_type = isinstance(value, bool)
    elif value_type == "integer":
        valid_type = isinstance(value, int) and not isinstance(value, bool)
    elif value_type == "number":
        valid_type = isinstance(value, (int, float)) and not isinstance(value, bool)
    elif value_type == "color":
        valid_type = isinstance(value, str) and bool(re.fullmatch(r"#[0-9A-Fa-f]{6}", value))
    elif value_type == "enum":
        valid_type = isinstance(value, str) and value in definition["values"]
    elif value_type == "enum_list":
        valid_type = (
            isinstance(value, list)
            and bool(value)
            and all(isinstance(item, str) for item in value)
            and len(value) == len(set(value))
            and all(item in definition["values"] for item in value)
        )
    elif value_type == "target":
        valid_type = isinstance(value, str) and value in TARGETS
    elif value_type == "template":
        template_errors = validate_template(value, [*path, "value"])
        if template_errors:
            return template_errors

    if not valid_type:
        errors.append(error_item("invalid_value", [*path, "value"], f"{name} 的值不符合 {value_type} 约束", value))
        return errors

    if value_type in {"integer", "number"}:
        if "minimum" in definition and value < definition["minimum"]:
            errors.append(error_item("out_of_range", [*path, "value"], f"{name} 不能小于 {definition['minimum']}", value))
        if "maximum" in definition and value > definition["maximum"]:
            errors.append(error_item("out_of_range", [*path, "value"], f"{name} 不能大于 {definition['maximum']}", value))

    expected_unit = definition.get("unit")
    if expected_unit:
        allowed_units = UNIT_ALIASES[expected_unit]
        if unit not in allowed_units:
            errors.append(
                error_item(
                    "invalid_unit",
                    [*path, "unit"],
                    f"{name} 的单位必须是 {sorted(allowed_units)}",
                    unit,
                )
            )
    elif has_unit:
        errors.append(error_item("unexpected_unit", [*path, "unit"], f"{name} 不接受单位", unit))
    return errors


def validate_property(name: str, action: dict[str, Any], path: list[str]) -> list[dict[str, Any]]:
    if name not in PROPERTY_DEFINITIONS:
        return [error_item("unknown_property", path, f"未知格式属性: {name}", name)]
    if action.get("action") == "set":
        return validate_property_value(name, action.get("value"), action.get("unit"), "unit" in action, path)
    if action.get("action") != "conditional":
        return []
    errors = []
    seen_conditions = set()
    for index, case in enumerate(action.get("cases", [])):
        if not isinstance(case, dict):
            continue
        case_path = [*path, "cases", str(index)]
        condition = case.get("when")
        condition_key = json.dumps(condition, ensure_ascii=False, sort_keys=True)
        if condition_key in seen_conditions:
            errors.append(error_item("duplicate_condition", [*case_path, "when"], "同一属性不能包含重复条件", condition))
        seen_conditions.add(condition_key)
        errors.extend(validate_condition(condition, [*case_path, "when"]))
        errors.extend(validate_property_value(name, case.get("value"), case.get("unit"), "unit" in case, case_path))
    fallback = action.get("fallback")
    if isinstance(fallback, dict) and fallback.get("action") == "set":
        errors.extend(validate_property_value(name, fallback.get("value"), fallback.get("unit"), "unit" in fallback, [*path, "fallback"]))
    return errors


def validate_selector(target: str, selector: dict[str, Any], path: list[str]) -> list[dict[str, Any]]:
    errors = []
    for index, pattern in enumerate(selector.get("fallback_regex", [])):
        regex_path = [*path, "fallback_regex", str(index)]
        if not isinstance(pattern, str):
            errors.append(error_item("invalid_regex", regex_path, "正则必须是字符串", pattern))
            continue
        if len(pattern) > 500:
            errors.append(error_item("unsafe_regex", regex_path, "正则长度不能超过500个字符", pattern))
            continue
        try:
            re.compile(pattern)
        except re.error as exc:
            errors.append(error_item("invalid_regex", regex_path, str(exc), pattern))
        if pattern in {".*", "^.*$", ".+", "^.+$"}:
            errors.append(error_item("unsafe_regex", regex_path, "不允许匹配任意文本的回退正则", pattern))
        capture_group = selector.get("capture_group")
        if isinstance(capture_group, int):
            try:
                if re.compile(pattern).groups < capture_group:
                    errors.append(error_item("invalid_capture_group", [*path, "capture_group"], "正则不存在指定捕获组", capture_group))
            except re.error:
                pass
    if selector.get("match_scope") == "text_span":
        within_target = selector.get("within_target")
        if not isinstance(within_target, str) or within_target not in TARGETS:
            errors.append(error_item("invalid_within_target", [*path, "within_target"], "文本片段选择器必须引用已知父 target", within_target))
        excluded_targets = selector.get("exclude_targets", [])
        for index, excluded_target in enumerate(excluded_targets if isinstance(excluded_targets, list) else []):
            if not isinstance(excluded_target, str) or excluded_target not in TARGETS:
                errors.append(error_item("invalid_exclude_target", [*path, "exclude_targets", str(index)], "排除项必须引用已知 target", excluded_target))
        if not selector.get("fallback_regex") and not excluded_targets:
            errors.append(error_item("missing_span_locator", path, "文本片段选择器必须提供定位正则或排除目标"))
    elif "within_target" in selector or "capture_group" in selector or "exclude_targets" in selector:
        errors.append(error_item("invalid_span_selector", path, "within_target/capture_group 只能用于 text_span", selector))
    return errors


def validate_spec(spec: dict[str, Any]) -> dict[str, Any]:
    """Return a stable machine-readable validation and AI repair report."""
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []

    validator = Draft202012Validator(load_schema())
    for issue in sorted(validator.iter_errors(spec), key=lambda item: list(item.absolute_path)):
        errors.append(error_item("schema_error", issue.absolute_path, issue.message, issue.instance))

    if spec.get("schema_version") != SCHEMA_VERSION:
        return {"status": "failed", "schema_version": SCHEMA_VERSION, "errors": errors, "warnings": warnings}

    target_rules: list[tuple[str, Any, list[str]]] = [("document", spec.get("document"), ["document"])]
    for target, rule in spec.get("targets", {}).items() if isinstance(spec.get("targets"), dict) else []:
        path = ["targets", target]
        if target not in TARGETS:
            errors.append(error_item("unknown_target", path, f"未知论文对象: {target}", target))
        target_rules.append((target, rule, path))

    for _target, rule, path in target_rules:
        if not isinstance(rule, dict) or not isinstance(rule.get("properties"), dict):
            continue
        for name, action in rule["properties"].items():
            if name not in PROPERTY_NAMES:
                errors.append(error_item("unknown_property", [*path, "properties", name], f"未知格式属性: {name}", name))
                continue
            if isinstance(action, dict):
                errors.extend(validate_property(name, action, [*path, "properties", name]))
        properties = rule["properties"]
        active_relative_position = isinstance(properties.get("section.relative_position"), dict) and properties["section.relative_position"].get("action") != "preserve"
        active_relative_target = isinstance(properties.get("section.relative_to"), dict) and properties["section.relative_to"].get("action") != "preserve"
        active_absolute_position = isinstance(properties.get("section.position"), dict) and properties["section.position"].get("action") != "preserve"
        if active_relative_position != active_relative_target:
            errors.append(error_item("incomplete_section_relation", [*path, "properties"], "section.relative_position 与 section.relative_to 必须成对出现"))
        if active_absolute_position and active_relative_position:
            errors.append(error_item("conflicting_section_position", [*path, "properties"], "绝对章节位置与相对章节位置不能同时设置"))
        relative_target = properties.get("section.relative_to", {})
        if (
            _target != "document"
            and isinstance(relative_target, dict)
            and relative_target.get("action") == "set"
            and relative_target.get("value") == _target
        ):
            errors.append(error_item("self_referencing_section", [*path, "properties", "section.relative_to", "value"], "章节不能相对于自身定位", _target))

    selectors = spec.get("selectors", {})
    if isinstance(selectors, dict):
        for target, selector in selectors.items():
            path = ["selectors", target]
            if target not in TARGETS:
                errors.append(error_item("unknown_selector_target", path, f"选择器指向未知对象: {target}", target))
            if isinstance(selector, dict):
                errors.extend(validate_selector(target, selector, path))

    unresolved = []
    for target, rule, _path in target_rules:
        if not isinstance(rule, dict):
            continue
        for name, action in rule.get("properties", {}).items():
            if isinstance(action, dict) and action.get("action") == "preserve" and action.get("reason") not in {
                "not_specified",
                "explicitly_preserve",
            }:
                unresolved.append({"target": target, "property": name, "reason": action.get("reason")})
    if unresolved:
        warnings.append({"warning_type": "unresolved_preserve", "message": "部分属性因未解决问题而保持原格式", "items": unresolved})

    return {
        "status": "failed" if errors else ("warn" if warnings else "success"),
        "schema_version": SCHEMA_VERSION,
        "errors": errors,
        "warnings": warnings,
        "repair_request": {
            "instruction": "只修复列出的字段；不得补充无来源格式值；无法确定的局部属性改为 preserve。",
            "errors": errors,
        }
        if errors
        else None,
    }


def load_and_validate(path: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    def reject_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"format_spec 包含重复JSON字段: {key}")
            result[key] = value
        return result

    data = json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_duplicate_pairs)
    if not isinstance(data, dict):
        raise ValueError("format_spec 顶层必须是对象")
    return data, validate_spec(data)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate canonical format_spec schema v2.")
    parser.add_argument("spec", type=Path)
    parser.add_argument("--report", type=Path)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    _spec, report = load_and_validate(args.spec)
    output = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(output, encoding="utf-8")
    print(output)
    if report["status"] == "failed":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
