#!/usr/bin/env python3
"""Generate a thin Word formatter wrapper for a JSON format specification."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from textwrap import dedent


SUPPORTED_SPEC_KEYS = {"default", "body", "page", "headings", "tables"}


def load_spec(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as handle:
        spec = json.load(handle)

    if not isinstance(spec, dict):
        raise ValueError("规格文件顶层必须是 JSON 对象")

    if not (SUPPORTED_SPEC_KEYS & set(spec)):
        raise ValueError(
            "规格文件不包含可识别字段；至少需要 default/body/page/headings/tables 之一"
        )

    if "rules" in spec:
        if not isinstance(spec.get("rules", []), list):
            raise ValueError("字段 rules 必须是数组")
        for index, rule in enumerate(spec.get("rules", [])):
            if "match" not in rule or "format" not in rule:
                raise ValueError(f"rules[{index}] 必须包含 match 和 format")

    return spec


def find_repo_root(start: Path) -> Path:
    for candidate in [start.resolve(), *start.resolve().parents]:
        if (candidate / "format_engine.py").is_file():
            return candidate
    return start.resolve()


def build_formatter_source(spec_path: Path, repo_root: Path) -> str:
    spec_literal = str(spec_path.resolve())
    repo_root_literal = str(repo_root.resolve())
    template = '''
#!/usr/bin/env python3
"""Thin Word formatter wrapper generated from a JSON specification."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


SPEC_PATH = Path(__SPEC_PATH__)
GENERATED_REPO_ROOT = Path(__REPO_ROOT__)


def find_repo_root(start: Path) -> Path:
    for candidate in [start, *start.parents]:
        if (candidate / "format_engine.py").is_file():
            return candidate
    if (GENERATED_REPO_ROOT / "format_engine.py").is_file():
        return GENERATED_REPO_ROOT
    raise RuntimeError("无法找到 format_engine.py；请在项目仓库内运行此 formatter。")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Format a .docx file according to the generated specification.")
    parser.add_argument("input", type=Path, help="原始 .docx 文件")
    parser.add_argument("output", type=Path, help="输出 .docx 文件")
    parser.add_argument("--report", type=Path, default=None, help="输出格式化报告 JSON")
    parser.add_argument("--show-spec", action="store_true", help="打印当前 formatter 使用的格式规格")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.show_spec:
        print(SPEC_PATH.read_text(encoding="utf-8"))
        return

    repo_root = find_repo_root(Path(__file__).resolve().parent)
    sys.path.insert(0, str(repo_root))
    from format_engine import format_document, load_spec

    spec = load_spec(SPEC_PATH)
    report = format_document(args.input, args.output, spec, args.report, SPEC_PATH)
    if report["status"] == "failed":
        if args.report is None:
            print(json.dumps(report, ensure_ascii=False, indent=2))
        raise SystemExit(1)
    print(f"已生成: {args.output}")
    if args.report:
        print(f"已生成: {args.report}")


if __name__ == "__main__":
    main()
'''
    return (
        dedent(template)
        .replace("__SPEC_PATH__", repr(spec_literal))
        .replace("__REPO_ROOT__", repr(repo_root_literal))
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate a thin Python Word formatter wrapper.")
    parser.add_argument("--spec", type=Path, required=True, help="JSON 格式要求文件")
    parser.add_argument("--output", type=Path, required=True, help="生成的 Python 脚本路径")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    load_spec(args.spec)
    source = build_formatter_source(args.spec, find_repo_root(Path(__file__).parent))
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(source, encoding="utf-8")
    args.output.chmod(0o755)
    print(f"已生成格式化脚本: {args.output}")


if __name__ == "__main__":
    main()
