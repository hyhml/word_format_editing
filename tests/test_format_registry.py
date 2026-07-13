from __future__ import annotations

import json
import shutil
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from format_registry import (
    build_manifest_template,
    combined_source_hash,
    fingerprint_sources,
    match_format_package,
)


class FormatRegistryTests(unittest.TestCase):
    def write_package(
        self,
        root: Path,
        source: Path,
        package_id: str = "demo",
        keywords: list[str] | None = None,
    ) -> Path:
        package_dir = root / "formats" / package_id
        package_dir.mkdir(parents=True)
        shutil.copyfile(source, package_dir / "format_spec.json")
        formatter = package_dir / "formatter.py"
        formatter.write_text("print('demo formatter')\n", encoding="utf-8")

        manifest = build_manifest_template(
            package_id=package_id,
            name="武汉科技大学本科毕业论文格式 2024",
            source_paths=[source],
            keywords=keywords or ["武汉科技大学", "本科毕业论文", "2024"],
        )
        manifest["created_at"] = "2026-07-14"
        (package_dir / "manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return package_dir

    def test_combined_hash_is_order_independent(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            first = root / "a.txt"
            second = root / "b.txt"
            first.write_text("alpha", encoding="utf-8")
            second.write_text("beta", encoding="utf-8")

            left = combined_source_hash(fingerprint_sources([first, second]))
            right = combined_source_hash(fingerprint_sources([second, first]))

            self.assertEqual(left, right)

    def test_exact_hash_match_reuses_existing_package(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "format.txt"
            source.write_text("正文小四号宋体，1.5倍行距", encoding="utf-8")
            package_dir = self.write_package(root, source)

            result = match_format_package([source], formats_dir=root / "formats")

            self.assertEqual(result.status, "matched")
            self.assertEqual(result.match_type, "exact_hash")
            self.assertEqual(result.package.package_dir, package_dir)

    def test_metadata_match_reuses_similar_package(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "known.txt"
            source.write_text("旧格式要求", encoding="utf-8")
            self.write_package(root, source)

            new_source = root / "new_notice.txt"
            new_source.write_text("补充通知：页边距保持不变", encoding="utf-8")
            result = match_format_package(
                [new_source],
                formats_dir=root / "formats",
                description="武汉科技大学 本科毕业论文 2024 格式要求",
            )

            self.assertEqual(result.status, "matched")
            self.assertEqual(result.match_type, "metadata")
            self.assertGreaterEqual(result.score, 0.6)

    def test_new_required_returns_clarification_questions(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "unknown.txt"
            source.write_text("未知格式要求", encoding="utf-8")

            result = match_format_package([source], formats_dir=root / "formats")

            self.assertEqual(result.status, "new_required")
            self.assertGreaterEqual(len(result.needs_clarification), 1)


if __name__ == "__main__":
    unittest.main()
