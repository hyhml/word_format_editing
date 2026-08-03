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
  ]
}
```

Use only targets and properties included in `ai_request.json`. Omit `unit` only when the property definition has no unit. Confidence must be between 0 and 1.

## Block classification

- `requirement`: Contains an executable or conditional formatting requirement. Emit at least one candidate using this block.
- `explanation`: Defines scope or terminology but does not itself set a format value.
- `example`: Shows sample content or appearance and is not an authoritative written rule.
- `irrelevant`: Administrative or submission content unrelated to Word formatting.
- `unresolved`: May contain a formatting requirement but cannot be normalized reliably.

Classify every block. Never call a block `explanation` merely because its rule is difficult.

## Normalization

Convert values before emitting candidates:

- Chinese font sizes to points, such as 小四 → `12 pt`.
- Millimetres and inches to centimetres.
- Fixed and minimum line spacing to points.
- Multiple line spacing to `multiple`.
- Alignment, numbering, orientation, and positions to ontology enums.

Expand relative and inherited rules into concrete candidates. If a dependency is unavailable, classify the requirement as `unresolved`; do not estimate it.

## Evidence and conflicts

Every candidate must cite at least one source block. Use multiple IDs when a rule spans blocks. Do not cite neighboring text that does not support the value.

When sources disagree, emit both evidenced candidates. The compiler will create one `preserve` action and put the alternatives in the recognition report. Only collapse them when the source explicitly establishes precedence, such as a later supplement overriding the general rule.

## Repair loop

Use `validation.repair_request` and `ai_candidate_errors` as the exclusive repair scope. Do not rewrite valid candidates during repair. Maximum: three total compile attempts. After that, mark affected blocks `unresolved` so the final policy preserves the original Word formatting.
