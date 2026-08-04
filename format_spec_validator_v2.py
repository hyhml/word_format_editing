#!/usr/bin/env python3
"""Validate canonical format_spec schema v2 and emit repair feedback."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Iterable

from jsonschema import Draft202012Validator

from format_ontology import PROPERTY_DEFINITIONS, PROPERTY_NAMES, SCHEMA_VERSION, TARGETS


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


def validate_property(name: str, action: dict[str, Any], path: list[str]) -> list[dict[str, Any]]:
    errors = []
    definition = PROPERTY_DEFINITIONS.get(name)
    if definition is None:
        return [error_item("unknown_property", path, f"未知格式属性: {name}", name)]
    if action.get("action") != "set":
        return errors

    value = action.get("value")
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
        valid_type = value in definition["values"]

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
        if action.get("unit") not in allowed_units:
            errors.append(
                error_item(
                    "invalid_unit",
                    [*path, "unit"],
                    f"{name} 的单位必须是 {sorted(allowed_units)}",
                    action.get("unit"),
                )
            )
    elif "unit" in action:
        errors.append(error_item("unexpected_unit", [*path, "unit"], f"{name} 不接受单位", action.get("unit")))
    return errors


def validate_selector(target: str, selector: dict[str, Any], path: list[str]) -> list[dict[str, Any]]:
    errors = []
    for index, pattern in enumerate(selector.get("fallback_regex", [])):
        regex_path = [*path, "fallback_regex", str(index)]
        if len(pattern) > 500:
            errors.append(error_item("unsafe_regex", regex_path, "正则长度不能超过500个字符", pattern))
            continue
        try:
            re.compile(pattern)
        except re.error as exc:
            errors.append(error_item("invalid_regex", regex_path, str(exc), pattern))
        if pattern in {".*", "^.*$", ".+", "^.+$"}:
            errors.append(error_item("unsafe_regex", regex_path, "不允许匹配任意文本的回退正则", pattern))
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
