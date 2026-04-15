from __future__ import annotations

import json


STYLE_SCHEMA = {
    "type": "object",
    "properties": {
        "fill": {"type": "string"},
        "stroke": {"type": "string"},
        "stroke_width": {"type": "number", "exclusiveMinimum": 0},
        "text_color": {"type": "string"},
        "font_size": {"type": "number", "exclusiveMinimum": 0},
        "font_weight": {"type": "string"},
        "dasharray": {"type": "string"},
        "label_fill": {"type": "string"},
        "label_stroke": {"type": "string"},
        "corner_radius": {"type": "number", "minimum": 0},
        "opacity": {"type": "number", "minimum": 0, "maximum": 1},
        "arrow_scale": {"type": "number", "exclusiveMinimum": 0},
    },
    "additionalProperties": False,
}

DOCUMENT_SCHEMA = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "title": "SchemeGen Document",
    "type": "object",
    "required": ["schema_version", "title", "diagrams"],
    "properties": {
        "schema_version": {"type": "string"},
        "title": {"type": "string"},
        "diagrams": {
            "type": "array",
            "minItems": 1,
            "items": {
                "type": "object",
                "required": ["id", "title", "type", "nodes", "edges"],
                "properties": {
                    "id": {"type": "string"},
                    "title": {"type": "string"},
                    "description": {"type": "string"},
                    "type": {"enum": ["flowchart", "idef0", "idef1x", "idef3", "uml_class"]},
                    "style": {
                        "type": "object",
                        "properties": {
                            "background_fill": {"type": "string"},
                            "frame_fill": {"type": "string"},
                            "frame_stroke": {"type": "string"},
                            "node": STYLE_SCHEMA,
                            "edge": STYLE_SCHEMA,
                        },
                        "additionalProperties": False,
                    },
                    "frame": {
                        "type": "object",
                        "properties": {
                            "enabled": {"type": "boolean"},
                            "paper_fill": {"type": "string"},
                            "used_at": {"type": "string"},
                            "author": {"type": "string"},
                            "project": {"type": "string"},
                            "date": {"type": "string"},
                            "revision": {"type": "string"},
                            "status": {"type": "string"},
                            "reader": {"type": "string"},
                            "context": {"type": "string"},
                            "notes": {"type": "string"},
                            "page": {"type": "string"},
                            "node_ref": {"type": "string"},
                        },
                        "additionalProperties": False,
                    },
                    "nodes": {
                        "type": "array",
                        "minItems": 1,
                        "items": {
                            "type": "object",
                            "required": ["id", "label", "kind", "row", "column"],
                            "properties": {
                                "id": {"type": "string"},
                                "label": {"type": "string"},
                                "kind": {"type": "string"},
                                "row": {"type": "integer"},
                                "column": {"type": "integer"},
                                "x": {"type": "integer"},
                                "y": {"type": "integer"},
                                "width": {"type": "integer", "minimum": 1},
                                "height": {"type": "integer", "minimum": 1},
                                "code": {"type": "string"},
                                "decomposes_to": {"type": "string"},
                                "stereotype": {"type": "string"},
                                "attributes": {
                                    "type": "array",
                                    "items": {"type": "string"},
                                },
                                "operations": {
                                    "type": "array",
                                    "items": {"type": "string"},
                                },
                                "columns": {
                                    "type": "array",
                                    "items": {
                                        "type": "object",
                                        "required": ["name"],
                                        "properties": {
                                            "name": {"type": "string"},
                                            "data_type": {"type": "string"},
                                            "key": {"enum": ["pk", "fk", "pk_fk", "unique"]},
                                            "nullable": {"type": "boolean"},
                                            "default": {"type": "string"},
                                        },
                                        "additionalProperties": False,
                                    },
                                },
                                "style": STYLE_SCHEMA,
                            },
                            "additionalProperties": False,
                        },
                    },
                    "edges": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "required": ["id"],
                            "properties": {
                                "id": {"type": "string"},
                                "source": {"type": ["string", "null"]},
                                "target": {"type": ["string", "null"]},
                                "label": {"type": "string"},
                                "role": {
                                    "type": "string",
                                    "enum": ["input", "control", "output", "mechanism"],
                                },
                                "kind": {
                                    "type": "string",
                                },
                                "icom": {"type": "string"},
                                "source_side": {
                                    "type": "string",
                                    "enum": ["top", "right", "bottom", "left"],
                                },
                                "target_side": {
                                    "type": "string",
                                    "enum": ["top", "right", "bottom", "left"],
                                },
                                "tunnel_source": {"type": "boolean"},
                                "tunnel_target": {"type": "boolean"},
                                "source_label": {"type": "string"},
                                "target_label": {"type": "string"},
                                "style": STYLE_SCHEMA,
                            },
                            "additionalProperties": False,
                        },
                    },
                },
                "additionalProperties": False,
            },
        },
    },
    "additionalProperties": False,
}


def format_schema() -> str:
    return json.dumps(DOCUMENT_SCHEMA, ensure_ascii=False, indent=2)
