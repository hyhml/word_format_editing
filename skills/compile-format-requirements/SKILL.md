---
name: compile-format-requirements
description: Compile thesis, dissertation, journal, or institutional formatting requirements from TXT, Markdown, JSON, legacy DOC, DOCX, or PDF into the repository's canonical format_spec schema v2. Use when Codex must identify, normalize, reconcile, or audit Word formatting rules before any document formatting is performed, especially when requirements contain tables, exceptions, inheritance, relative values, conflicts, or unspecified fields that must preserve the original format.
---

# Compile Format Requirements

Produce one deterministic `format_spec.json` plus recognition reports. Do not format a paper or generate formatter code while using this skill.

## Locate the compiler

Work from the repository root containing `format_compiler.py`, `format_spec_validator_v2.py`, and `schemas/format_spec.schema.json`. Stop if these files are unavailable.

Read [references/compilation-contract.md](references/compilation-contract.md) before creating AI candidates.

## Workflow

1. Generate the structured AI request from every user-provided requirement file:

   ```bash
   python3 format_compiler.py ai-request \
     --source path/to/requirements.pdf \
     --source path/to/supplement.docx \
     --output work/ai_request.json
   ```

   For legacy `.doc`, require `antiword`. Treat its text as usable evidence but explicitly report that original styles and table structure are unavailable.

2. Read the complete request. Classify every source block and write `ai_candidates.json` using only the supplied target and property vocabulary.

3. For every explicit requirement, emit concrete normalized candidates with real `evidence_ids`. Expand inheritance, conditions, and exceptions. Resolve relative values only when their dependencies are evidenced. Never insert common thesis defaults.

4. Classify every source block as `requirement`, `explanation`, `example`, `irrelevant`, or `unresolved`. A `requirement` block must support at least one candidate; otherwise classify it as `unresolved`. When a block is only partially representable, also emit each unsupported fragment in `unresolved_items`.

5. Compile and merge deterministic and AI recognition:

   ```bash
   python3 format_compiler.py compile \
     --source path/to/requirements.pdf \
     --source path/to/supplement.docx \
     --ai-candidates work/ai_candidates.json \
     --name "规范名称" \
     --description "用户描述" \
     --output-dir work/result
   ```

6. Inspect all of the following in `recognition_report.json`:

   - `validation.errors`
   - `validation.repair_request`
   - `ai_candidate_errors`
   - `conflicts`
   - `coverage.unmapped_blocks`
   - `coverage.unresolved_blocks`
   - `unresolved_items`

7. If validation or AI candidate errors exist, return only those structured errors and their cited blocks to the AI reasoning step. Repair only the failing candidates, then compile again. Perform at most three total compile attempts.

8. After three failed attempts, remove invalid unsupported candidates, classify the affected blocks as `unresolved`, and compile once to preserve those properties. If the overall Schema remains invalid, stop with `blocked`; do not publish or hand off the spec.

9. Validate the final artifact independently:

   ```bash
   python3 format_spec_validator_v2.py \
     work/result/format_spec.json \
     --report work/result/schema_validation.json
   ```

10. Deliver `format_spec.json`, `recognition_report.json`, and `recognition_report.md`. Explain conflicts and unresolved blocks. Do not invoke the legacy formatter: it intentionally rejects Schema v2.

## Integrity rules

- Treat source evidence as authoritative; treat model knowledge as non-evidence.
- Never invent evidence IDs, targets, properties, units, or enum values.
- Keep one final action for each target/property pair.
- Let unresolved conflicts compile to `preserve`; never silently choose a value.
- Use `preserve` semantics for unspecified or unresolved formatting; never substitute 宋体、小四、A4, or other defaults without evidence.
- Keep format-requirement extraction regex internal. Put regex into final selectors only when it locates a paper object.
- Do not leave natural-language instructions, inheritance, or relative calculations for the future formatter.
