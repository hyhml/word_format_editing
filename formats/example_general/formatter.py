#!/usr/bin/env python3
"""Thin formatter wrapper for the example_general format package."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> None:
    package_dir = Path(__file__).resolve().parent
    repo_root = package_dir.parents[1]
    output_script = package_dir / "_generated_formatter.py"

    subprocess.run(
        [
            sys.executable,
            str(repo_root / "generate_formatter.py"),
            "--spec",
            str(package_dir / "format_spec.json"),
            "--output",
            str(output_script),
        ],
        check=True,
    )
    subprocess.run([sys.executable, str(output_script), *sys.argv[1:]], check=True)


if __name__ == "__main__":
    main()
