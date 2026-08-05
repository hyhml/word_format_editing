# Resolution contract

The request produced by `build_runtime_context_request` contains:

- `document_fingerprint`: SHA-256 identity of the target DOCX.
- `required_fields`: the exact fields the format spec needs.
- `allowed_values`: fixed output enums.
- `evidence_blocks`: the only paper text that may support a decision.

Return one object:

```json
{
  "schema_version": "1.0.0",
  "document_fingerprint": "copy from request",
  "values": {
    "document.degree_type": {
      "status": "resolved",
      "value": "master",
      "confidence": 0.98,
      "evidence_ids": ["paper_paragraph_0001"],
      "reason": "扉页明确写有硕士学位论文"
    }
  }
}
```

Allowed `document.degree_type` values:

- `doctoral`: the paper explicitly identifies a doctoral degree or dissertation.
- `master`: the paper explicitly identifies an academic master's degree or thesis.
- `professional_master`: the paper explicitly identifies a professional master's program, such as engineering, MBA, or MPA.

Allowed `document.project_type` values:

- `thesis`: the paper explicitly identifies itself as a 本科毕业论文.
- `design`: the paper explicitly identifies itself as a 本科毕业设计.

Only return fields listed in `required_fields`. A request may require either field or both.

When evidence is insufficient or conflicting, return:

```json
{
  "status": "unknown",
  "value": null,
  "confidence": 0.0,
  "evidence_ids": [],
  "reason": "没有足够证据确定学位类型"
}
```

Return exactly every key in `required_fields` and no others. Copy the request fingerprint exactly. Cite only supplied evidence IDs. A resolved value needs at least one evidence ID. Keep `reason` concise and factual.
