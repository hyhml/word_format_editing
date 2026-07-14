"""OpenXML patch registry.

Module 2 only provides the patch framework. Concrete advanced patches will be
implemented in later milestones after the common python-docx engine is stable.
"""

from __future__ import annotations

from typing import Any


KNOWN_PATCHES = {
    "header_footer",
    "three_line_table",
    "captions",
    "equation_numbering",
    "math_font",
    "references",
}


def requested_patches(spec: dict[str, Any]) -> list[str]:
    raw = spec.get("openxml_patches", [])
    if raw is True:
        return sorted(KNOWN_PATCHES)
    if not raw:
        return []
    if not isinstance(raw, list):
        return [str(raw)]
    return [str(item) for item in raw]


def apply_requested_patches(_docx_path, spec: dict[str, Any]) -> dict[str, list[str]]:
    skipped = []
    unknown = []
    for patch_name in requested_patches(spec):
        if patch_name in KNOWN_PATCHES:
            skipped.append(f"{patch_name}: not implemented in module 2")
        else:
            unknown.append(f"{patch_name}: unknown patch")
    return {"applied": [], "skipped": skipped, "unknown": unknown}
