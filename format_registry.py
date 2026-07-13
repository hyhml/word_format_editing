#!/usr/bin/env python3
"""Format package registry for module 0.

This module decides whether a set of format requirement files has already
been converted into a reusable format package.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable


MANIFEST_NAME = "manifest.json"
DEFAULT_FORMATS_DIR = Path("formats")


class RegistryError(ValueError):
    """Raised when a format registry input or manifest is invalid."""


@dataclass(frozen=True)
class SourceFingerprint:
    path: str
    sha256: str
    size_bytes: int


@dataclass(frozen=True)
class FormatPackage:
    package_dir: Path
    manifest: dict[str, Any]

    @property
    def id(self) -> str:
        return str(self.manifest.get("id", self.package_dir.name))

    @property
    def name(self) -> str:
        return str(self.manifest.get("name", self.id))

    @property
    def combined_source_hash(self) -> str:
        return str(self.manifest.get("combined_source_hash", ""))

    @property
    def keywords(self) -> list[str]:
        values = self.manifest.get("keywords", [])
        if not isinstance(values, list):
            return []
        return [str(value) for value in values if str(value).strip()]


@dataclass(frozen=True)
class MatchResult:
    status: str
    sources: list[SourceFingerprint]
    combined_source_hash: str
    package: FormatPackage | None = None
    match_type: str | None = None
    score: float = 0.0
    needs_clarification: list[str] = field(default_factory=list)
    candidates: list[dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        package_data = None
        if self.package is not None:
            package_data = {
                "id": self.package.id,
                "name": self.package.name,
                "path": str(self.package.package_dir),
                "manifest": str(self.package.package_dir / MANIFEST_NAME),
            }

        return {
            "status": self.status,
            "match_type": self.match_type,
            "score": self.score,
            "combined_source_hash": self.combined_source_hash,
            "sources": [
                {
                    "path": source.path,
                    "sha256": source.sha256,
                    "size_bytes": source.size_bytes,
                }
                for source in self.sources
            ],
            "package": package_data,
            "needs_clarification": list(self.needs_clarification),
            "candidates": list(self.candidates),
        }


def sha256_file(path: Path) -> str:
    """Return a SHA-256 digest for one source file."""
    if not path.is_file():
        raise RegistryError(f"格式要求文件不存在或不是文件: {path}")

    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def fingerprint_sources(paths: Iterable[Path]) -> list[SourceFingerprint]:
    """Fingerprint source files while preserving the caller's display paths."""
    fingerprints = []
    seen = set()
    for raw_path in paths:
        path = raw_path.expanduser().resolve()
        if path in seen:
            continue
        seen.add(path)
        fingerprints.append(
            SourceFingerprint(
                path=str(path),
                sha256=sha256_file(path),
                size_bytes=path.stat().st_size,
            )
        )

    if not fingerprints:
        raise RegistryError("至少需要提供一个格式要求文件")

    return fingerprints


def combined_source_hash(fingerprints: Iterable[SourceFingerprint]) -> str:
    """Return an order-independent hash for a format requirement file set."""
    source_hashes = sorted(source.sha256 for source in fingerprints)
    digest = hashlib.sha256()
    digest.update(b"word-format-source-set-v1\n")
    for source_hash in source_hashes:
        digest.update(source_hash.encode("ascii"))
        digest.update(b"\n")
    return digest.hexdigest()


def load_manifest(path: Path) -> dict[str, Any]:
    try:
        with path.open("r", encoding="utf-8") as handle:
            manifest = json.load(handle)
    except json.JSONDecodeError as exc:
        raise RegistryError(f"manifest 不是合法 JSON: {path}: {exc}") from exc

    if not isinstance(manifest, dict):
        raise RegistryError(f"manifest 顶层必须是对象: {path}")

    required = {"id", "name", "combined_source_hash", "format_spec", "formatter"}
    missing = sorted(required - set(manifest))
    if missing:
        raise RegistryError(f"manifest 缺少字段 {missing}: {path}")

    return manifest


def scan_format_packages(formats_dir: Path = DEFAULT_FORMATS_DIR) -> list[FormatPackage]:
    """Load all valid format packages under formats_dir."""
    if not formats_dir.exists():
        return []
    if not formats_dir.is_dir():
        raise RegistryError(f"formats 路径不是目录: {formats_dir}")

    packages = []
    for manifest_path in sorted(formats_dir.glob(f"*/{MANIFEST_NAME}")):
        manifest = load_manifest(manifest_path)
        package = FormatPackage(package_dir=manifest_path.parent, manifest=manifest)
        validate_package_files(package)
        packages.append(package)
    return packages


def validate_package_files(package: FormatPackage) -> None:
    for key in ("format_spec", "formatter"):
        rel_path = package.manifest.get(key)
        if not rel_path:
            raise RegistryError(f"{package.package_dir}: manifest 缺少 {key}")
        target = package.package_dir / str(rel_path)
        if not target.is_file():
            raise RegistryError(f"{package.package_dir}: {key} 指向的文件不存在: {target}")


def description_tokens(description: str) -> set[str]:
    """Extract simple Chinese/English/year tokens for metadata matching."""
    tokens = set()
    for token in re.findall(r"[\u4e00-\u9fffA-Za-z0-9]+", description.lower()):
        token = token.strip()
        if len(token) >= 2:
            tokens.add(token)
    return tokens


def package_metadata_text(package: FormatPackage) -> str:
    pieces = [
        package.id,
        package.name,
        str(package.manifest.get("institution", "")),
        str(package.manifest.get("document_type", "")),
        str(package.manifest.get("year", "")),
    ]
    pieces.extend(package.keywords)
    return " ".join(piece for piece in pieces if piece)


def metadata_score(package: FormatPackage, query_text: str) -> float:
    query_tokens = description_tokens(query_text)
    if not query_tokens:
        return 0.0

    metadata_text = package_metadata_text(package).lower()
    metadata_tokens = description_tokens(metadata_text)
    if not metadata_tokens:
        return 0.0

    hits = set()
    for token in query_tokens:
        if token in metadata_tokens or token in metadata_text:
            hits.add(token)

    keyword_hits = 0
    for keyword in package.keywords:
        keyword_norm = keyword.lower().strip()
        if keyword_norm and keyword_norm in query_text.lower():
            keyword_hits += 1

    token_score = len(hits) / max(len(query_tokens), 1)
    keyword_bonus = min(keyword_hits * 0.15, 0.45)
    return min(token_score + keyword_bonus, 1.0)


def semantic_candidates(
    packages: Iterable[FormatPackage],
    query_text: str,
) -> list[dict[str, Any]]:
    """Placeholder for future embedding-based semantic matching."""
    _ = list(packages)
    _ = query_text
    return []


def clarification_questions(has_description: bool) -> list[str]:
    questions = [
        "请提供格式名称，例如学校、期刊或单位名称。",
        "请提供格式年份或版本，例如 2024 版、修订版。",
        "请说明文档类型，例如本科毕业论文、硕士论文、期刊论文或公文。",
    ]
    if has_description:
        return [
            "当前描述无法可靠匹配已有格式包，请补充学校/期刊、年份和文档类型。",
        ]
    return questions


def match_format_package(
    source_paths: Iterable[Path],
    formats_dir: Path = DEFAULT_FORMATS_DIR,
    description: str = "",
    metadata_threshold: float = 0.6,
) -> MatchResult:
    """Match a source file set against known format packages."""
    sources = fingerprint_sources(source_paths)
    combined_hash = combined_source_hash(sources)
    packages = scan_format_packages(formats_dir)

    for package in packages:
        if package.combined_source_hash == combined_hash:
            return MatchResult(
                status="matched",
                sources=sources,
                combined_source_hash=combined_hash,
                package=package,
                match_type="exact_hash",
                score=1.0,
            )

    query_parts = [description]
    query_parts.extend(Path(source.path).name for source in sources)
    query_text = " ".join(part for part in query_parts if part)

    scored = []
    for package in packages:
        score = metadata_score(package, query_text)
        if score > 0:
            scored.append(
                {
                    "id": package.id,
                    "name": package.name,
                    "path": str(package.package_dir),
                    "score": round(score, 4),
                    "match_type": "metadata",
                }
            )
    scored.sort(key=lambda item: item["score"], reverse=True)

    if scored and scored[0]["score"] >= metadata_threshold:
        package_id = scored[0]["id"]
        package = next(pkg for pkg in packages if pkg.id == package_id)
        return MatchResult(
            status="matched",
            sources=sources,
            combined_source_hash=combined_hash,
            package=package,
            match_type="metadata",
            score=float(scored[0]["score"]),
            candidates=scored[:5],
        )

    semantic = semantic_candidates(packages, query_text)
    candidates = scored[:5] + semantic[:5]
    has_description = bool(description.strip())
    return MatchResult(
        status="new_required",
        sources=sources,
        combined_source_hash=combined_hash,
        needs_clarification=clarification_questions(has_description),
        candidates=candidates,
    )


def build_manifest_template(
    package_id: str,
    name: str,
    source_paths: Iterable[Path],
    keywords: Iterable[str] = (),
) -> dict[str, Any]:
    """Build a manifest skeleton for a new format package."""
    sources = fingerprint_sources(source_paths)
    return {
        "id": package_id,
        "name": name,
        "version": "0.1.0",
        "created_at": "",
        "source_hashes": [
            {
                "path": Path(source.path).name,
                "sha256": source.sha256,
                "size_bytes": source.size_bytes,
            }
            for source in sources
        ],
        "combined_source_hash": combined_source_hash(sources),
        "keywords": [str(keyword) for keyword in keywords if str(keyword).strip()],
        "format_spec": "format_spec.json",
        "formatter": "formatter.py",
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Check or prepare reusable Word format packages.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    check = subparsers.add_parser("check", help="check whether a format source set matches an existing package")
    check.add_argument("--formats-dir", type=Path, default=DEFAULT_FORMATS_DIR)
    check.add_argument("--description", default="", help="用户提供的学校、期刊、年份或格式说明")
    check.add_argument("--metadata-threshold", type=float, default=0.6)
    check.add_argument("sources", nargs="+", type=Path, help="格式要求文件")

    manifest = subparsers.add_parser("manifest-template", help="print a manifest template for a new package")
    manifest.add_argument("--id", required=True, help="格式包 id")
    manifest.add_argument("--name", required=True, help="格式包名称")
    manifest.add_argument("--keyword", action="append", default=[], help="可重复传入的关键词")
    manifest.add_argument("sources", nargs="+", type=Path, help="格式要求文件")

    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.command == "check":
        result = match_format_package(
            source_paths=args.sources,
            formats_dir=args.formats_dir,
            description=args.description,
            metadata_threshold=args.metadata_threshold,
        )
        print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2))
        return

    if args.command == "manifest-template":
        manifest = build_manifest_template(
            package_id=args.id,
            name=args.name,
            source_paths=args.sources,
            keywords=args.keyword,
        )
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
        return

    raise RegistryError(f"未知命令: {args.command}")


if __name__ == "__main__":
    main()
