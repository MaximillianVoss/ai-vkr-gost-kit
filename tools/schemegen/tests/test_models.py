from __future__ import annotations

from pathlib import Path
import unittest

from schemegen.models import ValidationError, load_document_file


ROOT = Path(__file__).resolve().parent.parent


class ModelTests(unittest.TestCase):
    def test_examples_are_valid(self) -> None:
        for example_name in ("flowchart_only.json", "coursework_document.json", "test_progression.json"):
            document = load_document_file(ROOT / "examples" / example_name)
            self.assertGreaterEqual(len(document.diagrams), 1)

    def test_idef0_child_requires_three_to_six_boxes(self) -> None:
        invalid = {
            "schema_version": "1.0",
            "title": "Broken IDEF0",
            "diagrams": [
                {
                    "id": "a1",
                    "title": "A1",
                    "type": "idef0",
                    "nodes": [
                        {"id": "n1", "label": "Шаг 1", "kind": "function", "row": 0, "column": 0},
                        {"id": "n2", "label": "Шаг 2", "kind": "function", "row": 0, "column": 1}
                    ],
                    "edges": []
                }
            ]
        }

        with self.assertRaises(ValidationError):
            from schemegen.models import load_document

            load_document(invalid)

    def test_idef0_child_boundary_arrow_requires_icom_or_tunnel(self) -> None:
        invalid = {
            "schema_version": "1.0",
            "title": "Broken IDEF0",
            "diagrams": [
                {
                    "id": "a1",
                    "title": "A1",
                    "type": "idef0",
                    "nodes": [
                        {"id": "n1", "label": "Принять", "kind": "function", "row": 0, "column": 0},
                        {"id": "n2", "label": "Проверить", "kind": "function", "row": 0, "column": 1},
                        {"id": "n3", "label": "Выдать", "kind": "function", "row": 0, "column": 2}
                    ],
                    "edges": [
                        {"id": "e1", "source": None, "target": "n1", "role": "input", "label": "Заявка"}
                    ]
                }
            ]
        }

        with self.assertRaises(ValidationError):
            from schemegen.models import load_document

            load_document(invalid)

    def test_invalid_decomposition_reference_raises(self) -> None:
        invalid = {
            "schema_version": "1.0",
            "title": "Broken",
            "diagrams": [
                {
                    "id": "main",
                    "title": "Main",
                    "type": "flowchart",
                    "nodes": [
                        {
                            "id": "n1",
                            "label": "Start",
                            "kind": "start",
                            "row": 0,
                            "column": 0,
                            "decomposes_to": "missing"
                        }
                    ],
                    "edges": []
                }
            ]
        }

        with self.assertRaises(ValidationError):
            from schemegen.models import load_document

            load_document(invalid)

    def test_empty_diagram_raises(self) -> None:
        invalid = {
            "schema_version": "1.0",
            "title": "Broken",
            "diagrams": [
                {
                    "id": "main",
                    "title": "Main",
                    "type": "flowchart",
                    "nodes": [],
                    "edges": []
                }
            ]
        }

        with self.assertRaises(ValidationError):
            from schemegen.models import load_document

            load_document(invalid)
