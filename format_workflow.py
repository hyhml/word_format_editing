#!/usr/bin/env python3
"""Workflow gate and formatter launcher for module 4."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import traceback
from pathlib import Path
from typing import Any


def repo_root() -> Path:
    return Path(__file__).resolve().parent


def load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    return data if isinstance(data, dict) else {}


def resolve_inputs(args: argparse.Namespace) -> dict[str, Path | None]:
    format_package = args.format_package.resolve() if args.format_package else None
    spec = args.spec.resolve() if args.spec else None
    formatter = args.formatter.resolve() if args.formatter else None

    if format_package is not None:
        if spec is None:
            spec = format_package / "format_spec.json"
        if formatter is None:
            formatter = format_package / "formatter.py"

    return {
        "input": args.input.resolve(),
        "format_package": format_package,
        "spec": spec,
        "formatter": formatter,
        "structure": args.structure.resolve(),
        "output": args.output.resolve(),
        "workflow_report_json": args.workflow_report_json.resolve(),
        "workflow_report_md": args.workflow_report_md.resolve(),
        "format_report_json": args.format_report_json.resolve(),
        "format_report_md": args.format_report_md.resolve(),
    }


def missing_item(artifact: str, return_to: str, message: str) -> dict[str, str]:
    return {"artifact": artifact, "return_to": return_to, "message": message}


def can_use_engine(spec: Path | None) -> bool:
    return spec is not None and spec.is_file() and (repo_root() / "format_engine.py").is_file()


def ensure_writable_dir(directory: Path, report: dict[str, Any]) -> None:
    try:
        directory.mkdir(parents=True, exist_ok=True)
        probe = directory / ".workflow_write_probe"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
    except Exception as exc:
        report["errors"].append({"message": f"目录不可写: {directory}: {exc}"})


def initial_report(paths: dict[str, Path | None]) -> dict[str, Any]:
    return {
        "status": "pending",
        "input": str(paths["input"]),
        "format_package": str(paths["format_package"]) if paths["format_package"] else None,
        "spec": str(paths["spec"]) if paths["spec"] else None,
        "formatter": str(paths["formatter"]) if paths["formatter"] else None,
        "structure": str(paths["structure"]),
        "output": str(paths["output"]),
        "executor": None,
        "format_report": str(paths["format_report_json"]),
        "missing": [],
        "errors": [],
    }


def gate(paths: dict[str, Path | None], report: dict[str, Any]) -> None:
    input_path = paths["input"]
    spec = paths["spec"]
    formatter = paths["formatter"]
    structure = paths["structure"]
    output = paths["output"]

    if input_path is None or not input_path.is_file():
        report["missing"].append(
            missing_item("raw.docx", "input", "缺少原始论文 Word 文件。")
        )
    elif input_path.suffix.lower() != ".docx":
        report["missing"].append(
            missing_item("raw.docx", "input", "当前只支持 .docx 输入。")
        )

    if spec is None or not spec.is_file():
        report["missing"].append(
            missing_item("format_spec.json", "module_1", "缺少格式规范，请先运行格式要求解析。")
        )

    formatter_exists = formatter is not None and formatter.is_file()
    if not formatter_exists and not can_use_engine(spec):
        report["missing"].append(
            missing_item("formatter.py", "module_2", "缺少 formatter.py，且无法回退到 format_engine.py。")
        )

    if structure is None or not structure.is_file():
        report["missing"].append(
            missing_item("paper_structure.json", "module_3", "缺少论文结构文件，请先运行论文结构识别。")
        )

    writable_dirs = {
        output.parent,
        paths["workflow_report_json"].parent,
        paths["workflow_report_md"].parent,
        paths["format_report_json"].parent,
        paths["format_report_md"].parent,
    }
    for directory in writable_dirs:
        ensure_writable_dir(directory, report)

    if report["missing"]:
        report["status"] = "blocked"
    elif report["errors"]:
        report["status"] = "failed"


def run_formatter(paths: dict[str, Path | None], report: dict[str, Any]) -> None:
    formatter = paths["formatter"]
    spec = paths["spec"]
    input_path = paths["input"]
    output = paths["output"]
    format_report_json = paths["format_report_json"]

    if formatter is not None and formatter.is_file():
        command = [
            sys.executable,
            str(formatter),
            str(input_path),
            str(output),
            "--report",
            str(format_report_json),
        ]
        report["executor"] = "formatter"
    else:
        command = [
            sys.executable,
            str(repo_root() / "format_engine.py"),
            "--spec",
            str(spec),
            "--input",
            str(input_path),
            "--output",
            str(output),
            "--report",
            str(format_report_json),
        ]
        report["executor"] = "format_engine"

    completed = subprocess.run(
        command,
        cwd=repo_root(),
        text=True,
        capture_output=True,
    )
    report["command"] = command
    report["stdout"] = completed.stdout
    report["stderr"] = completed.stderr
    if completed.returncode != 0:
        report["status"] = "failed"
        report["errors"].append(
            {
                "message": f"formatter 执行失败，退出码 {completed.returncode}",
                "stderr": completed.stderr,
            }
        )
        return

    if not output.is_file():
        report["status"] = "failed"
        report["errors"].append({"message": "formatter 执行成功但未生成输出 docx。"})
        return

    report["status"] = "success"


def render_format_report_md(format_report_json: Path, format_report_md: Path) -> None:
    if not format_report_json.is_file():
        format_report_md.write_text("# 格式化报告\n\n- 未生成 format_report.json\n", encoding="utf-8")
        return
    data = load_json(format_report_json)
    lines = [
        "# 格式化报告",
        "",
        f"- 状态：{data.get('status', 'unknown')}",
        f"- 输入：{data.get('input', '')}",
        f"- 输出：{data.get('output', '')}",
        "",
        "## 已应用",
        "",
    ]
    applied = data.get("applied", [])
    if applied:
        lines.extend(f"- {item}" for item in applied)
    else:
        lines.append("- 无")
    lines.extend(["", "## 跳过", ""])
    skipped = data.get("skipped", []) + data.get("skipped_patches", [])
    if skipped:
        lines.extend(f"- {item}" for item in skipped)
    else:
        lines.append("- 无")
    if data.get("errors"):
        lines.extend(["", "## 错误", ""])
        lines.extend(f"- {item.get('message', item)}" for item in data["errors"])
    lines.append("")
    format_report_md.write_text("\n".join(lines), encoding="utf-8")


def render_workflow_report_md(report: dict[str, Any]) -> str:
    lines = [
        "# 工作流执行报告",
        "",
        f"- 状态：{report['status']}",
        f"- 输入：{report['input']}",
        f"- 输出：{report['output']}",
        f"- 格式包：{report.get('format_package')}",
        f"- 格式规范：{report.get('spec')}",
        f"- Formatter：{report.get('formatter')}",
        f"- 结构文件：{report.get('structure')}",
        f"- 执行器：{report.get('executor')}",
        "",
        "## 缺失项",
        "",
    ]
    if report["missing"]:
        for item in report["missing"]:
            lines.append(f"- {item['artifact']} -> {item['return_to']}：{item['message']}")
    else:
        lines.append("- 无")

    lines.extend(["", "## 错误", ""])
    if report["errors"]:
        for error in report["errors"]:
            lines.append(f"- {error.get('message', error)}")
    else:
        lines.append("- 无")
    lines.append("")
    return "\n".join(lines)


def write_reports(paths: dict[str, Path | None], report: dict[str, Any]) -> None:
    workflow_json = paths["workflow_report_json"]
    workflow_md = paths["workflow_report_md"]
    format_report_json = paths["format_report_json"]
    format_report_md = paths["format_report_md"]

    workflow_json.parent.mkdir(parents=True, exist_ok=True)
    workflow_md.parent.mkdir(parents=True, exist_ok=True)
    format_report_md.parent.mkdir(parents=True, exist_ok=True)

    workflow_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    workflow_md.write_text(render_workflow_report_md(report), encoding="utf-8")
    if report["status"] == "success" or format_report_json.is_file():
        render_format_report_md(format_report_json, format_report_md)


def run_workflow(args: argparse.Namespace) -> dict[str, Any]:
    paths = resolve_inputs(args)
    report = initial_report(paths)
    try:
        gate(paths, report)
        if report["status"] == "pending":
            run_formatter(paths, report)
    except Exception as exc:
        report["status"] = "failed"
        report["errors"].append({"message": str(exc), "traceback": traceback.format_exc()})
    write_reports(paths, report)
    return report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Gate and launch the Word formatting workflow.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    run = subparsers.add_parser("run", help="run workflow gate and formatter")
    run.add_argument("--input", type=Path, required=True, help="原始 .docx")
    run.add_argument("--format-package", type=Path, default=None, help="格式包目录")
    run.add_argument("--spec", type=Path, default=None, help="format_spec.json")
    run.add_argument("--formatter", type=Path, default=None, help="formatter.py")
    run.add_argument("--structure", type=Path, required=True, help="paper_structure.json")
    run.add_argument("--output", type=Path, required=True, help="输出 formatted.docx")
    run.add_argument("--workflow-report-json", type=Path, required=True)
    run.add_argument("--workflow-report-md", type=Path, required=True)
    run.add_argument("--format-report-json", type=Path, required=True)
    run.add_argument("--format-report-md", type=Path, required=True)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.command == "run":
        report = run_workflow(args)
        print(f"已生成: {args.workflow_report_json}")
        print(f"已生成: {args.workflow_report_md}")
        if report["status"] == "success":
            print(f"已生成: {args.output}")
            print(f"已生成: {args.format_report_json}")
            print(f"已生成: {args.format_report_md}")
        if report["status"] != "success":
            raise SystemExit(1 if report["status"] == "failed" else 2)
        return
    raise RuntimeError(f"未知命令: {args.command}")


if __name__ == "__main__":
    main()
