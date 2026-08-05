# Compilation contract

## Candidate file

Write one JSON object:

```json
{
  "candidates": [
    {
      "target": "body.paragraph",
      "property": "font.size_pt",
      "value": 12,
      "unit": "pt",
      "when": {
        "field": "document.degree_type",
        "operator": "equals",
        "value": "master"
      },
      "evidence_ids": ["source_01_page_0008_line_0003"],
      "confidence": 0.98
    }
  ],
  "block_classifications": [
    {
      "evidence_id": "source_01_page_0008_line_0003",
      "classification": "requirement",
      "notes": "Explicit body font-size rule"
    }
  ],
  "unresolved_items": [
    {
      "evidence_id": "source_01_page_0008_line_0003",
      "text": "无法用当前属性词表完整表达的原文片段",
      "reason": "unsupported_property",
      "notes": "保留原格式并在报告中提示"
    }
  ]
}
```

Use only targets and properties included in `ai_request.json`. Omit `unit` only when the property definition has no unit. Confidence must be between 0 and 1.

`when` is optional. Use it only when the source explicitly states applicability, and only with the condition fields, operators, targets, and normalized values supplied in `ai_request.json`. Candidates for the same target/property under different conditions compile into one `conditional` action. If several cases match at execution time, the mandatory result is `preserve`; never rely on case order.

For numbering and content layouts, use `content.template` with structured segments instead of prose:

```json
{
  "segments": [
    {"kind": "field", "field": "chapter_number"},
    {"kind": "spacer", "count": 2, "unit": "character"},
    {"kind": "field", "field": "title"},
    {"kind": "leader", "character": "…"},
    {"kind": "field", "field": "page_number"}
  ]
}
```

Use only template fields exposed by the ontology. Punctuation is a `literal` segment. Use a `tab` segment with `left`, `center`, or `right` alignment for header/TOC tab stops. Do not put natural-language layout instructions into template values.

When the source explicitly permits one of several layouts, use a `choice` segment whose `options` are complete nested `{ "segments": [...] }` templates. Do not use `choice` merely to hide uncertainty; ambiguous source wording remains unresolved.

Some template fields declare runtime inputs in the ontology. For example, `degree_name` is resolved from `document.degree_type` in the paper being formatted. Keep this field dynamic instead of guessing “硕士” or “博士”; if the later paper-analysis stage cannot establish the input, the executor must preserve the original text.

## Block classification

- `requirement`: Contains an executable or conditional formatting requirement. Emit at least one candidate using this block.
- `explanation`: Defines scope or terminology but does not itself set a format value.
- `example`: Shows sample content or appearance and is not an authoritative written rule.
- `irrelevant`: Administrative or submission content unrelated to Word formatting.
- `unresolved`: May contain a formatting requirement but cannot be normalized reliably.

Classify every block. Never call a block `explanation` merely because its rule is difficult.

If one block contains both normalized candidates and unsupported details, classify it as `requirement`, emit the supported candidates, and add every unsupported fragment to `unresolved_items`. Allowed reasons are `ambiguous`, `missing_dependency`, `unsupported_property`, and `source_incomplete`.

Only format-related fragments belong in `unresolved_items`. Content-quality statements, required personal signatures, administrative submission instructions, printing, and physical binding are outside this compiler's scope. They may be mentioned in classification notes but do not make the format specification incomplete.

## Normalization

Convert values before emitting candidates:

- Chinese font sizes to points, such as 小四 → `12 pt`.
- Millimetres and inches to centimetres.
- Fixed and minimum line spacing to points.
- Multiple line spacing to `multiple`.
- Alignment, numbering, orientation, and positions to ontology enums.

Expand relative and inherited rules into concrete candidates. If a dependency is unavailable, classify the requirement as `unresolved`; do not estimate it.

Represent explicit conditional applicability with `when`, chapter placement with `section.position` or a `section.relative_position` plus `section.relative_to`, and mixed in-paragraph formatting with the dedicated child targets whose selectors use `text_span` scope.

## Evidence and conflicts

Every candidate must cite at least one source block. Use multiple IDs when a rule spans blocks. Do not cite neighboring text that does not support the value.

When sources disagree, emit both evidenced candidates. The compiler will create one `preserve` action and put the alternatives in the recognition report. Only collapse them when the source explicitly establishes precedence, such as a later supplement overriding the general rule.

## Repair loop

Use `validation.repair_request` and `ai_candidate_errors` as the exclusive repair scope. Do not rewrite valid candidates during repair. Maximum: three total compile attempts. After that, mark affected blocks `unresolved` so the final policy preserves the original Word formatting.
