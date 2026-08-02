#!/usr/bin/env python3
"""End-to-end pipeline for Word format editing.

This module wires modules 0 through 5 into one command:

1. Match or create a reusable format package.
2. Analyze the raw paper structure.
3. Run the formatter workflow.
4. Validate the formatted document.
5. Write a pipeline summary report.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import traceback
from datetime import date
from pathlib import Path
from typing import Any, Iterable

from format_parser import parse_format_requirements, write_outputs as write_parser_outputs
from format_registry import build_manifest_template, description_tokens, match_format_package
from format_validator import validate_document
from format_workflow import run_workflow
from generate_formatter import build_formatter_source, find_repo_root, load_spec as validate_formatter_spec
from paper_structure import analyze_docx, write_outputs as write_structure_outputs


DEFAULT_OUTPUT_NAME = "formatted.docx"


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def copy_sources(source_paths: Iterable[Path], source_dir: Path) -> list[str]:
    source_dir.mkdir(parents=True, exist_ok=True)
    copied = []
    used_names: set[str] = set()
    for index, source in enumerate(source_paths, start=1):
        source = source.expanduser().resolve()
        suffix = source.suffix
        base_name = f"source_{index:02d}{suffix}" if suffix else f"source_{index:02d}"
        target = source_dir / base_name
        while target.name in used_names:
            target = source_dir / f"source_{index:02d}_{len(used_names)}{suffix}"
        shutil.copyfile(source, target)
        used_names.add(target.name)
        copied.append(str(target.relative_to(source_dir.parent)))
    return copied


def package_id_slug(text: str) -> str:
    tokens = re.findall(r"[A-Za-z0-9]+", text.lower())
    slug = "_".join(tokens[:8]).strip("_")
    return slug[:60]


def generate_package_id(
    explicit_id: str,
    name: str,
    description: str,
    combined_hash: str,
) -> str:
    if explicit_id.strip():
        candidate = package_id_slug(explicit_id)
    else:
        candidate = package_id_slug(name or description)
    if not candidate:
        candidate = f"format_{combined_hash[:8]}"
    return candidate


def unique_package_dir(formats_dir: Path, package_id: str) -> Path:
    package_dir = formats_dir / package_id
    if not package_dir.exists():
        return package_dir
    for index in range(2, 1000):
        candidate = formats_dir / f"{package_id}_{index}"
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"无法为格式包生成唯一目录: {package_id}")


def keywords_from_text(text: str) -> list[str]:
    return sorted(description_tokens(text))


def create_format_package(
    source_paths: list[Path],
    formats_dir: Path,
    description: str,
    name: str,
    package_id: str,
    combined_hash: str,
) -> Path:
    package_name = name.strip() or description.strip() or package_id or f"格式包 {combined_hash[:8]}"
    package_id = generate_package_id(package_id, package_name, description, combined_hash)
    package_dir = unique_package_dir(formats_dir, package_id)
    package_dir.mkdir(parents=True, exist_ok=False)

    spec, spec_md, parse_report = parse_format_requirements(
        source_paths=source_paths,
        description=description,
        name=package_name,
    )
    write_parser_outputs(package_dir, spec, spec_md, parse_report)

    formatter_path = package_dir / "formatter.py"
    validate_formatter_spec(package_dir / "format_spec.json")
    formatter_source = build_formatter_source(package_dir / "format_spec.json", find_repo_root(Path(__file__).parent))
    formatter_path.write_text(formatter_source, encoding="utf-8")
    formatter_path.chmod(0o755)

    copied_sources = copy_sources(source_paths, package_dir / "source")
    manifest = build_manifest_template(
        package_id=package_dir.name,
        name=package_name,
        source_paths=source_paths,
        keywords=keywords_from_text(f"{package_name} {description}"),
    )
    manifest["created_at"] = date.today().isoformat()
    manifest["source_hashes"] = [
        {**source_hash, "path": copied_sources[index]}
        for index, source_hash in enumerate(manifest["source_hashes"])
    ]
    manifest["metadata"] = {
        "institution": spec.get("metadata", {}).get("institution"),
        "document_type": spec.get("metadata", {}).get("document_type"),
        "year": spec.get("metadata", {}).get("year"),
    }
    write_json(package_dir / "manifest.json", manifest)
    return package_dir


def artifact_paths(output_dir: Path) -> dict[str, Path]:
    return {
        "formatted_docx": output_dir / DEFAULT_OUTPUT_NAME,
        "paper_structure_md": output_dir / "paper_structure.md",
        "paper_structure_json": output_dir / "paper_structure.json",
        "structure_report_json": output_dir / "structure_report.json",
        "workflow_report_json": output_dir / "workflow_report.json",
        "workflow_report_md": output_dir / "workflow_report.md",
        "format_report_json": output_dir / "format_report.json",
        "format_report_md": output_dir / "format_report.md",
        "validation_report_json": output_dir / "validation_report.json",
        "validation_report_md": output_dir / "validation_report.md",
        "pipeline_report_json": output_dir / "pipeline_report.json",
        "pipeline_report_md": output_dir / "pipeline_report.md",
    }


def initial_report(args: argparse.Namespace, paths: dict[str, Path]) -> dict[str, Any]:
    return {
        "status": "pending",
        "input": str(args.input.expanduser().resolve()),
        "format_sources": [str(path.expanduser().resolve()) for path in args.format_source],
        "formats_dir": str(args.formats_dir.expanduser().resolve()),
        "output_dir": str(args.output_dir.expanduser().resolve()),
        "description": args.description,
        "name": args.name,
        "package_action": None,
        "format_package": None,
        "registry": None,
        "workflow_status": None,
        "validation_status": None,
        "artifacts": {key: str(value) for key, value in paths.items()},
        "errors": [],
    }


def workflow_args(input_path: Path, package_dir: Path, paths: dict[str, Path]) -> argparse.Namespace:
    return argparse.Namespace(
        input=input_path,
        format_package=package_dir,
        spec=None,
        formatter=None,
        structure=paths["paper_structure_json"],
        output=paths["formatted_docx"],
        workflow_report_json=paths["workflow_report_json"],
        workflow_report_md=paths["workflow_report_md"],
        format_report_json=paths["format_report_json"],
        format_report_md=paths["format_report_md"],
    )


def render_pipeline_report_md(report: dict[str, Any]) -> str:
    lines = [
        "# 端到端格式化报告",
        "",
        f"- 状态：{report['status']}",
        f"- 输入：{report['input']}",
        f"- 输出目录：{report['output_dir']}",
        f"- 格式包动作：{report.get('package_action')}",
        f"- 格式包：{report.get('format_package')}",
        f"- 工作流状态：{report.get('workflow_status')}",
        f"- 校验状态：{report.get('validation_status')}",
        "",
        "## 格式要求来源",
        "",
    ]
    lines.extend(f"- {source}" for source in report["format_sources"])
    lines.extend(["", "## 产物", ""])
    for key, value in report["artifacts"].items():
        lines.append(f"- {key}: {value}")
    lines.extend(["", "## 错误", ""])
    if report["errors"]:
        for error in report["errors"]:
            lines.append(f"- {error.get('message', error)}")
    else:
        lines.append("- 无")
    lines.append("")
    return "\n".join(lines)


def final_status(workflow_status: str | None, validation_status: str | None) -> str:
    if workflow_status != "success":
        return "failed"
    if validation_status in {"failed", "fail"}:
        return "failed"
    if validation_status == "warn":
        return "warn"
    return "success"


def run_pipeline(args: argparse.Namespace) -> dict[str, Any]:
    output_dir = args.output_dir.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = artifact_paths(output_dir)
    report = initial_report(args, paths)

    try:
        input_path = args.input.expanduser().resolve()
        source_paths = [path.expanduser().resolve() for path in args.format_source]
        formats_dir = args.formats_dir.expanduser().resolve()

        registry_result = match_format_package(
            source_paths,
            formats_dir=formats_dir,
            description=args.description,
            metadata_threshold=args.metadata_threshold,
        )
        report["registry"] = registry_result.to_dict()

        if registry_result.package is not None:
            package_dir = registry_result.package.package_dir.resolve()
            report["package_action"] = "reused"
        else:
            package_dir = create_format_package(
                source_paths=source_paths,
                formats_dir=formats_dir,
                description=args.description,
                name=args.name,
                package_id=args.package_id,
                combined_hash=registry_result.combined_source_hash,
            ).resolve()
            report["package_action"] = "created"

        report["format_package"] = str(package_dir)

        structure, structure_md, structure_report = analyze_docx(input_path)
        write_structure_outputs(
            paths["paper_structure_json"],
            paths["paper_structure_md"],
            paths["structure_report_json"],
            structure,
            structure_md,
            structure_report,
        )

        workflow_report = run_workflow(workflow_args(input_path, package_dir, paths))
        report["workflow_status"] = workflow_report.get("status")

        if workflow_report.get("status") == "success":
            validation_report = validate_document(
                paths["formatted_docx"],
                package_dir / "format_spec.json",
                paths["paper_structure_json"],
                paths["validation_report_json"],
                paths["validation_report_md"],
            )
            report["validation_status"] = validation_report.get("status")
        else:
            report["validation_status"] = "skipped"
            report["errors"].extend(workflow_report.get("errors", []))

        report["status"] = final_status(report["workflow_status"], report["validation_status"])
    except Exception as exc:
        report["status"] = "failed"
        report["errors"].append({"message": str(exc), "traceback": traceback.format_exc()})

    write_json(paths["pipeline_report_json"], report)
    paths["pipeline_report_md"].write_text(render_pipeline_report_md(report), encoding="utf-8")
    return report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the full Word format editing pipeline.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    run = subparsers.add_parser("run", help="run modules 0 through 5")
    run.add_argument("--input", type=Path, required=True, help="原始论文 .docx")
    run.add_argument(
        "--format-source",
        type=Path,
        action="append",
        required=True,
        help="格式要求文件，可重复传入",
    )
    run.add_argument("--output-dir", type=Path, required=True, help="输出目录")
    run.add_argument("--formats-dir", type=Path, default=Path("formats"), help="格式包目录")
    run.add_argument("--description", default="", help="学校、期刊、年份或格式说明")
    run.add_argument("--name", default="", help="新建格式包名称")
    run.add_argument("--package-id", default="", help="新建格式包 id；不传则自动生成")
    run.add_argument("--metadata-threshold", type=float, default=0.6, help="元数据复用匹配阈值")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.command == "run":
        report = run_pipeline(args)
        print(f"已生成: {args.output_dir / 'pipeline_report.json'}")
        print(f"已生成: {args.output_dir / 'pipeline_report.md'}")
        if report["status"] in {"failed"}:
            raise SystemExit(1)
        return
    raise RuntimeError(f"未知命令: {args.command}")


if __name__ == "__main__":
    main()
