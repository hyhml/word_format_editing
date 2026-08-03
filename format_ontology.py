"""Canonical targets and properties for format specification schema v2."""

from __future__ import annotations

from typing import Any


SCHEMA_VERSION = "2.0.0"

TARGETS = {
    "document.page",
    "document.section",
    "document.pagination",
    "document.header",
    "document.footer",
    "cover",
    "originality_statement",
    "authorization_statement",
    "title_page",
    "abstract.zh.heading",
    "abstract.zh.body",
    "abstract.en.heading",
    "abstract.en.body",
    "keywords.zh",
    "keywords.en",
    "table_of_contents.heading",
    "table_of_contents.level_1",
    "table_of_contents.level_2",
    "table_of_contents.level_3",
    "list_of_figures",
    "list_of_tables",
    "body.paragraph",
    "heading.level_1",
    "heading.level_2",
    "heading.level_3",
    "heading.level_4",
    "heading.level_5",
    "heading.level_6",
    "paragraph.quote",
    "paragraph.list",
    "paragraph.code",
    "figure",
    "figure.caption",
    "figure.note",
    "figure.source",
    "table",
    "table.caption",
    "table.header",
    "table.body",
    "table.note",
    "table.continuation",
    "equation",
    "equation.number",
    "equation.explanation",
    "references.heading",
    "references.entry",
    "references.category_heading",
    "acknowledgements.heading",
    "acknowledgements.body",
    "appendix.heading",
    "appendix.body",
    "author_biography",
    "research_achievements",
    "footnote",
    "endnote",
    "hyperlink",
    "cross_reference",
    "textbox",
    "long_quote",
}

PRESERVE_REASONS = {
    "not_specified",
    "not_recognized",
    "unresolved_value",
    "unresolved_conflict",
    "unresolved_dependency",
    "unsupported",
    "explicitly_preserve",
}

ACTIONS = {"set", "preserve", "remove"}

ALIGNMENTS = {"left", "center", "right", "justify", "distributed"}
ORIENTATIONS = {"portrait", "landscape"}
LINE_SPACING_TYPES = {"single", "multiple", "fixed", "at_least"}
NUMBER_FORMATS = {"arabic", "lower_roman", "upper_roman", "lower_letter", "upper_letter", "chinese"}
POSITIONS = {"above", "below", "left", "right", "center", "inside", "outside"}

# Type names are interpreted by format_spec_validator_v2.  Keeping this table
# data-only makes it reusable by prompts, the CLI, and future format engines.
PROPERTY_DEFINITIONS: dict[str, dict[str, Any]] = {
    "font.east_asia": {"type": "string"},
    "font.latin": {"type": "string"},
    "font.math": {"type": "string"},
    "font.size_pt": {"type": "number", "minimum": 1, "maximum": 200, "unit": "pt"},
    "font.bold": {"type": "boolean"},
    "font.italic": {"type": "boolean"},
    "font.underline": {"type": "boolean"},
    "font.color": {"type": "color"},
    "font.character_spacing_pt": {"type": "number", "minimum": -20, "maximum": 100, "unit": "pt"},
    "font.scale_percent": {"type": "number", "minimum": 1, "maximum": 600, "unit": "percent"},
    "paragraph.alignment": {"type": "enum", "values": ALIGNMENTS},
    "paragraph.line_spacing.type": {"type": "enum", "values": LINE_SPACING_TYPES},
    "paragraph.line_spacing.value": {"type": "number", "minimum": 0, "maximum": 200, "unit": "multiple_or_pt"},
    "paragraph.space_before_pt": {"type": "number", "minimum": 0, "maximum": 500, "unit": "pt"},
    "paragraph.space_after_pt": {"type": "number", "minimum": 0, "maximum": 500, "unit": "pt"},
    "paragraph.first_line_indent_cm": {"type": "number", "minimum": -20, "maximum": 20, "unit": "cm"},
    "paragraph.first_line_indent_chars": {"type": "number", "minimum": -20, "maximum": 20, "unit": "character"},
    "paragraph.left_indent_cm": {"type": "number", "minimum": -20, "maximum": 20, "unit": "cm"},
    "paragraph.right_indent_cm": {"type": "number", "minimum": -20, "maximum": 20, "unit": "cm"},
    "paragraph.keep_with_next": {"type": "boolean"},
    "paragraph.keep_together": {"type": "boolean"},
    "paragraph.page_break_before": {"type": "boolean"},
    "paragraph.widow_control": {"type": "boolean"},
    "page.paper_size": {"type": "enum", "values": {"A3", "A4", "A5", "LETTER", "LEGAL", "CUSTOM"}},
    "page.width_cm": {"type": "number", "minimum": 1, "maximum": 200, "unit": "cm"},
    "page.height_cm": {"type": "number", "minimum": 1, "maximum": 200, "unit": "cm"},
    "page.orientation": {"type": "enum", "values": ORIENTATIONS},
    "page.margin_top_cm": {"type": "number", "minimum": 0, "maximum": 30, "unit": "cm"},
    "page.margin_bottom_cm": {"type": "number", "minimum": 0, "maximum": 30, "unit": "cm"},
    "page.margin_left_cm": {"type": "number", "minimum": 0, "maximum": 30, "unit": "cm"},
    "page.margin_right_cm": {"type": "number", "minimum": 0, "maximum": 30, "unit": "cm"},
    "page.gutter_cm": {"type": "number", "minimum": 0, "maximum": 30, "unit": "cm"},
    "page.header_distance_cm": {"type": "number", "minimum": 0, "maximum": 30, "unit": "cm"},
    "page.footer_distance_cm": {"type": "number", "minimum": 0, "maximum": 30, "unit": "cm"},
    "page.different_first_page": {"type": "boolean"},
    "page.different_odd_even": {"type": "boolean"},
    "numbering.enabled": {"type": "boolean"},
    "numbering.pattern": {"type": "string"},
    "numbering.sequence": {"type": "enum", "values": {"continuous", "by_chapter", "by_section"}},
    "numbering.format": {"type": "enum", "values": NUMBER_FORMATS},
    "numbering.restart": {"type": "boolean"},
    "numbering.start": {"type": "integer", "minimum": 0, "maximum": 100000},
    "numbering.position": {"type": "enum", "values": POSITIONS},
    "numbering.separator": {"type": "string"},
    "caption.position": {"type": "enum", "values": POSITIONS},
    "table.style": {"type": "enum", "values": {"three_line", "grid", "custom"}},
    "table.alignment": {"type": "enum", "values": {"left", "center", "right"}},
    "table.width_cm": {"type": "number", "minimum": 0, "maximum": 200, "unit": "cm"},
    "table.autofit": {"type": "boolean"},
    "table.allow_row_break": {"type": "boolean"},
    "table.repeat_header": {"type": "boolean"},
    "table.vertical_alignment": {"type": "enum", "values": {"top", "center", "bottom"}},
    "table.border_top_pt": {"type": "number", "minimum": 0, "maximum": 20, "unit": "pt"},
    "table.border_bottom_pt": {"type": "number", "minimum": 0, "maximum": 20, "unit": "pt"},
    "table.border_header_pt": {"type": "number", "minimum": 0, "maximum": 20, "unit": "pt"},
    "object.alignment": {"type": "enum", "values": {"left", "center", "right"}},
    "object.width_cm": {"type": "number", "minimum": 0, "maximum": 200, "unit": "cm"},
    "object.height_cm": {"type": "number", "minimum": 0, "maximum": 200, "unit": "cm"},
    "object.wrap": {"type": "enum", "values": {"inline", "square", "tight", "top_bottom", "behind", "in_front"}},
    "references.style": {"type": "string"},
    "references.hanging_indent_cm": {"type": "number", "minimum": 0, "maximum": 20, "unit": "cm"},
    "section.start": {"type": "enum", "values": {"continuous", "new_page", "odd_page", "even_page"}},
    "text.content": {"type": "string"},
}

PROPERTY_NAMES = set(PROPERTY_DEFINITIONS)


def ontology_summary() -> dict[str, Any]:
    """Return JSON-serializable ontology data for AI prompts and reports."""
    definitions = {}
    for name, definition in PROPERTY_DEFINITIONS.items():
        item = dict(definition)
        if isinstance(item.get("values"), set):
            item["values"] = sorted(item["values"])
        definitions[name] = item
    return {
        "schema_version": SCHEMA_VERSION,
        "targets": sorted(TARGETS),
        "properties": definitions,
        "actions": sorted(ACTIONS),
        "preserve_reasons": sorted(PRESERVE_REASONS),
    }
