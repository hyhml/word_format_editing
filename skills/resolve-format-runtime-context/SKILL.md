---
name: resolve-format-runtime-context
description: Resolve only the runtime variables that a canonical format_spec needs from the target paper DOCX, such as master versus doctoral degree type for a running-header template. Use after format requirements have been compiled when format_spec contains template runtime inputs and a target Word file is available. Ask AI for a bounded, evidence-backed classification; do not review, rewrite, summarize, or judge paper content, and do not format the document.
---

# Resolve Format Runtime Context

Supply narrowly scoped context values to a later format executor. Do not create a separate CLI, edit the DOCX, or analyze content beyond the fields required by `format_spec`.

## Locate the library

Work from the repository root containing `runtime_context.py`, `format_ontology.py`, and `schemas/runtime_context.schema.json`. Stop if these files are unavailable.

Read [references/resolution-contract.md](references/resolution-contract.md) before asking AI to classify evidence.

## Workflow

1. Load the canonical `format_spec.json` and call `required_runtime_inputs(spec)`. If it returns no AI-resolved fields, stop successfully without reading the paper.
2. Call `build_runtime_context_request(spec, paper_path)`. Give the AI only the returned request, including its bounded `evidence_blocks`; do not send unrelated body text or use the filename as evidence.
3. Ask the AI to return exactly the JSON object in the resolution contract. It must resolve every requested field or explicitly return `unknown`.
4. Call `validate_runtime_context_response(request, response)`. Repair only listed structural, scope, fingerprint, or evidence-ID errors. Maximum: three AI attempts.
5. Call `normalize_runtime_context(request, response)`. Values below the confidence threshold automatically become `unknown`.
6. Call `template_context(...)` only with the normalized result. Missing fields remain absent so the later executor preserves the original text.
7. Hand the normalized context and its evidence to the future format executor. Do not invoke the legacy formatter while Schema v2.1 execution remains unsupported.

## Decision rules

- Use only evidence blocks supplied by the request.
- Prefer explicit cover, title-page, degree-category, and application-degree statements.
- Treat a filename as metadata, never as sole evidence.
- Return `unknown` for conflicting types, indirect mentions, absent evidence, or uncertainty.
- Never infer degree type from research topic, department, writing style, author biography, or cited works.
- Never assess abstract quality, argument quality, originality, correctness, or writing quality.
- Cache results only by the exact `document_fingerprint` and field name.
- Preserve original formatting whenever validation fails, context is unknown, or confidence is too low.
