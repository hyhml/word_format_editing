from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document
from jsonschema import Draft202012Validator

from runtime_context import (
    build_runtime_context_request,
    load_runtime_context_schema,
    normalize_runtime_context,
    required_runtime_inputs,
    template_context,
    validate_runtime_context_response,
)


class RuntimeContextTests(unittest.TestCase):
    def dynamic_header_spec(self) -> dict:
        return {
            "document": {"properties": {}},
            "targets": {
                "body.header.odd": {
                    "properties": {
                        "content.template": {
                            "action": "set",
                            "value": {
                                "segments": [
                                    {"kind": "field", "field": "institution_name"},
                                    {"kind": "field", "field": "degree_name"},
                                    {"kind": "literal", "text": "学位论文"},
                                    {"kind": "field", "field": "page_number"},
                                ]
                            },
                            "evidence_ids": ["source_01_paragraph_0001"],
                        }
                    }
                }
            },
        }

    def make_paper(self, path: Path, *paragraphs: str) -> None:
        document = Document()
        for text in paragraphs:
            document.add_paragraph(text)
        document.save(path)

    def test_runtime_context_schema_is_valid(self) -> None:
        Draft202012Validator.check_schema(load_runtime_context_schema())

    def test_only_ai_resolved_template_fields_are_requested(self) -> None:
        self.assertEqual(required_runtime_inputs(self.dynamic_header_spec()), ["document.degree_type"])

        static_spec = {
            "document": {"properties": {}},
            "targets": {
                "body.header.odd": {
                    "properties": {
                        "content.template": {
                            "action": "set",
                            "value": {"segments": [{"kind": "field", "field": "page_number"}]},
                            "evidence_ids": ["source_01_paragraph_0001"],
                        }
                    }
                }
            },
        }
        self.assertEqual(required_runtime_inputs(static_spec), [])

    def test_request_extracts_bounded_degree_evidence_from_docx(self) -> None:
        with TemporaryDirectory() as tmp:
            paper = Path(tmp) / "paper.docx"
            self.make_paper(paper, "华东理工大学", "硕士学位论文", "正文内容与格式判断无关")

            request = build_runtime_context_request(self.dynamic_header_spec(), paper)

            self.assertEqual(request["required_fields"], ["document.degree_type"])
            self.assertRegex(request["document_fingerprint"], r"^[0-9a-f]{64}$")
            self.assertTrue(any(block["text"] == "硕士学位论文" for block in request["evidence_blocks"]))
            self.assertNotIn("正文内容与格式判断无关", {block["text"] for block in request["evidence_blocks"]})

    def test_valid_ai_resolution_produces_template_context(self) -> None:
        with TemporaryDirectory() as tmp:
            paper = Path(tmp) / "paper.docx"
            self.make_paper(paper, "硕士学位论文")
            request = build_runtime_context_request(self.dynamic_header_spec(), paper)
            evidence_id = request["evidence_blocks"][0]["id"]
            response = {
                "schema_version": "1.0.0",
                "document_fingerprint": request["document_fingerprint"],
                "values": {
                    "document.degree_type": {
                        "status": "resolved",
                        "value": "master",
                        "confidence": 0.98,
                        "evidence_ids": [evidence_id],
                        "reason": "扉页明确写有硕士学位论文",
                    }
                },
            }

            validation = validate_runtime_context_response(request, response)
            normalized = normalize_runtime_context(request, response)

            self.assertEqual(validation["status"], "success")
            self.assertEqual(normalized["status"], "success")
            self.assertEqual(template_context(normalized, "华东理工大学"), {"degree_name": "硕士", "institution_name": "华东理工大学"})

    def test_low_confidence_becomes_unknown_and_preserves_text(self) -> None:
        with TemporaryDirectory() as tmp:
            paper = Path(tmp) / "paper.docx"
            self.make_paper(paper, "疑似硕士论文")
            request = build_runtime_context_request(self.dynamic_header_spec(), paper)
            evidence_id = request["evidence_blocks"][0]["id"]
            response = {
                "schema_version": "1.0.0",
                "document_fingerprint": request["document_fingerprint"],
                "values": {
                    "document.degree_type": {
                        "status": "resolved",
                        "value": "master",
                        "confidence": 0.6,
                        "evidence_ids": [evidence_id],
                        "reason": "只有间接表述",
                    }
                },
            }

            normalized = normalize_runtime_context(request, response)

            self.assertEqual(normalized["status"], "warn")
            self.assertEqual(normalized["values"]["document.degree_type"]["status"], "unknown")
            self.assertNotIn("degree_name", template_context(normalized))

    def test_wrong_document_or_invented_evidence_is_blocked(self) -> None:
        with TemporaryDirectory() as tmp:
            paper = Path(tmp) / "paper.docx"
            self.make_paper(paper, "博士学位论文")
            request = build_runtime_context_request(self.dynamic_header_spec(), paper)
            response = {
                "schema_version": "1.0.0",
                "document_fingerprint": "0" * 64,
                "values": {
                    "document.degree_type": {
                        "status": "resolved",
                        "value": "doctoral",
                        "confidence": 0.99,
                        "evidence_ids": ["invented"],
                        "reason": "博士论文",
                    }
                },
            }

            validation = validate_runtime_context_response(request, response)
            error_types = {item["error_type"] for item in validation["errors"]}

            self.assertEqual(validation["status"], "failed")
            self.assertIn("document_mismatch", error_types)
            self.assertIn("unknown_evidence", error_types)

    def test_malformed_ai_response_is_blocked_without_crashing(self) -> None:
        request = {
            "document_fingerprint": "1" * 64,
            "required_fields": ["document.degree_type"],
            "evidence_blocks": [],
        }

        validation = validate_runtime_context_response(request, ["not", "an", "object"])
        normalized = normalize_runtime_context(request, ["not", "an", "object"])

        self.assertEqual(validation["status"], "failed")
        self.assertEqual(normalized["status"], "blocked")
        self.assertTrue(validation["errors"])


if __name__ == "__main__":
    unittest.main()
