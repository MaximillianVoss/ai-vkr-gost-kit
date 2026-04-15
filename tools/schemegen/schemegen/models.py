from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import json
import re


FLOWCHART_KIND_ALIASES = {
    "start": "terminator",
    "end": "terminator",
    "operation": "process",
    "io": "data",
    "on_page_connector": "connector",
    "off_page_reference": "offpage",
    "off_page": "offpage",
}
FLOWCHART_KINDS = {
    "terminator",
    "process",
    "decision",
    "data",
    "document",
    "predefined_process",
    "connector",
    "offpage",
}
IDEF0_KIND_ALIASES = {
    "activity": "function",
}
IDEF0_KINDS = {"function"}
IDEF1X_KIND_ALIASES = {
    "table": "entity",
    "dependent": "dependent_entity",
}
IDEF1X_KINDS = {"entity", "dependent_entity"}
IDEF3_KIND_ALIASES = {
    "junction": "junction_x",
    "uob_box": "uob",
}
IDEF3_KINDS = {"uob", "junction_x", "junction_and", "junction_or"}
UML_CLASS_KIND_ALIASES = {
    "abstract": "abstract_class",
}
UML_CLASS_KINDS = {"class", "abstract_class", "interface", "enum"}
ALLOWED_DIAGRAM_TYPES = {"flowchart", "idef0", "idef1x", "idef3", "uml_class"}
ALLOWED_SIDES = {"top", "right", "bottom", "left"}
ALLOWED_IDEF0_ROLES = {"input", "control", "output", "mechanism"}
ALLOWED_UML_EDGE_KINDS = {
    "association",
    "inheritance",
    "aggregation",
    "composition",
    "dependency",
    "realization",
}
ALLOWED_IDEF1X_EDGE_KINDS = {"identifying", "non_identifying"}
ALLOWED_COLUMN_KEYS = {"pk", "fk", "pk_fk", "unique"}


class ValidationError(ValueError):
    """Raised when the JSON document does not match the expected format."""


@dataclass(slots=True)
class Style:
    fill: str | None = None
    stroke: str | None = None
    stroke_width: float | None = None
    text_color: str | None = None
    font_size: float | None = None
    font_weight: str | None = None
    dasharray: str | None = None
    label_fill: str | None = None
    label_stroke: str | None = None
    corner_radius: float | None = None
    opacity: float | None = None
    arrow_scale: float | None = None


@dataclass(slots=True)
class DiagramStyle:
    background_fill: str | None = None
    node: Style = field(default_factory=Style)
    edge: Style = field(default_factory=Style)
    frame_fill: str | None = None
    frame_stroke: str | None = None


@dataclass(slots=True)
class Column:
    name: str
    data_type: str | None = None
    key: str | None = None
    nullable: bool | None = None
    default: str | None = None


@dataclass(slots=True)
class Node:
    id: str
    label: str
    kind: str
    row: int
    column: int
    x: int | None = None
    y: int | None = None
    width: int | None = None
    height: int | None = None
    code: str | None = None
    decomposes_to: str | None = None
    stereotype: str | None = None
    attributes: list[str] = field(default_factory=list)
    operations: list[str] = field(default_factory=list)
    columns: list[Column] = field(default_factory=list)
    style: Style | None = None


@dataclass(slots=True)
class Edge:
    id: str
    source: str | None
    target: str | None
    label: str | None = None
    role: str | None = None
    kind: str | None = None
    icom: str | None = None
    source_side: str | None = None
    target_side: str | None = None
    tunnel_source: bool = False
    tunnel_target: bool = False
    source_label: str | None = None
    target_label: str | None = None
    style: Style | None = None


@dataclass(slots=True)
class Idef0Frame:
    enabled: bool = True
    paper_fill: str | None = None
    used_at: str | None = None
    author: str | None = None
    project: str | None = None
    date: str | None = None
    revision: str | None = None
    status: str | None = None
    reader: str | None = None
    context: str | None = None
    notes: str | None = None
    page: str | None = None
    node_ref: str | None = None


@dataclass(slots=True)
class Diagram:
    id: str
    title: str
    type: str
    nodes: list[Node]
    edges: list[Edge]
    description: str | None = None
    frame: Idef0Frame | None = None
    style: DiagramStyle | None = None

    @property
    def node_map(self) -> dict[str, Node]:
        return {node.id: node for node in self.nodes}


@dataclass(slots=True)
class Document:
    schema_version: str
    title: str
    diagrams: list[Diagram]

    @property
    def diagram_map(self) -> dict[str, Diagram]:
        return {diagram.id: diagram for diagram in self.diagrams}


def _expect_type(data: object, expected_type: type, path: str) -> object:
    if not isinstance(data, expected_type):
        expected_name = expected_type.__name__
        actual_name = type(data).__name__
        raise ValidationError(f"{path}: expected {expected_name}, got {actual_name}")
    return data


def _get_string(data: dict[str, object], key: str, path: str, required: bool = True) -> str | None:
    value = data.get(key)
    if value is None:
        if required:
            raise ValidationError(f"{path}.{key}: missing value")
        return None
    if not isinstance(value, str) or not value.strip():
        raise ValidationError(f"{path}.{key}: expected non-empty string")
    return value.strip()


def _get_int(data: dict[str, object], key: str, path: str, required: bool = True) -> int | None:
    value = data.get(key)
    if value is None:
        if required:
            raise ValidationError(f"{path}.{key}: missing value")
        return None
    if not isinstance(value, int):
        raise ValidationError(f"{path}.{key}: expected integer")
    return value


def _get_number(data: dict[str, object], key: str, path: str, required: bool = True) -> float | None:
    value = data.get(key)
    if value is None:
        if required:
            raise ValidationError(f"{path}.{key}: missing value")
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValidationError(f"{path}.{key}: expected number")
    return float(value)


def _get_bool(data: dict[str, object], key: str, path: str, required: bool = True) -> bool | None:
    value = data.get(key)
    if value is None:
        if required:
            raise ValidationError(f"{path}.{key}: missing value")
        return None
    if not isinstance(value, bool):
        raise ValidationError(f"{path}.{key}: expected boolean")
    return value


def _get_string_list(data: dict[str, object], key: str, path: str) -> list[str]:
    raw_value = data.get(key)
    if raw_value is None:
        return []
    raw_value = _expect_type(raw_value, list, f"{path}.{key}")
    assert isinstance(raw_value, list)
    values: list[str] = []
    for index, item in enumerate(raw_value):
        if not isinstance(item, str) or not item.strip():
            raise ValidationError(f"{path}.{key}[{index}]: expected non-empty string")
        values.append(item.strip())
    return values


def _normalize_kind(diagram_type: str, raw_kind: str) -> str:
    if diagram_type == "flowchart":
        normalized = FLOWCHART_KIND_ALIASES.get(raw_kind, raw_kind)
        if normalized not in FLOWCHART_KINDS:
            raise ValidationError(
                f"Unsupported flowchart node kind '{raw_kind}'. "
                f"Allowed: {', '.join(sorted(FLOWCHART_KINDS | set(FLOWCHART_KIND_ALIASES)))}"
            )
        return normalized

    if diagram_type == "idef0":
        normalized = IDEF0_KIND_ALIASES.get(raw_kind, raw_kind)
        if normalized not in IDEF0_KINDS:
            raise ValidationError(
                f"Unsupported IDEF0 node kind '{raw_kind}'. "
                f"Allowed: {', '.join(sorted(IDEF0_KINDS | set(IDEF0_KIND_ALIASES)))}"
            )
        return normalized

    if diagram_type == "idef1x":
        normalized = IDEF1X_KIND_ALIASES.get(raw_kind, raw_kind)
        if normalized not in IDEF1X_KINDS:
            raise ValidationError(
                f"Unsupported IDEF1X node kind '{raw_kind}'. "
                f"Allowed: {', '.join(sorted(IDEF1X_KINDS | set(IDEF1X_KIND_ALIASES)))}"
            )
        return normalized

    if diagram_type == "uml_class":
        normalized = UML_CLASS_KIND_ALIASES.get(raw_kind, raw_kind)
        if normalized not in UML_CLASS_KINDS:
            raise ValidationError(
                f"Unsupported UML class node kind '{raw_kind}'. "
                f"Allowed: {', '.join(sorted(UML_CLASS_KINDS | set(UML_CLASS_KIND_ALIASES)))}"
            )
        return normalized

    normalized = IDEF3_KIND_ALIASES.get(raw_kind, raw_kind)
    if normalized not in IDEF3_KINDS:
        raise ValidationError(
            f"Unsupported IDEF3 node kind '{raw_kind}'. "
            f"Allowed: {', '.join(sorted(IDEF3_KINDS | set(IDEF3_KIND_ALIASES)))}"
        )
    return normalized


def _normalize_side(value: str | None, path: str) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or value not in ALLOWED_SIDES:
        raise ValidationError(f"{path}: side must be one of {', '.join(sorted(ALLOWED_SIDES))}")
    return value


def _normalize_role(value: str | None, diagram_type: str, path: str) -> str | None:
    if value is None:
        return None
    if diagram_type != "idef0":
        raise ValidationError(f"{path}: role is only supported for idef0 diagrams")
    if not isinstance(value, str) or value not in ALLOWED_IDEF0_ROLES:
        raise ValidationError(
            f"{path}: role must be one of {', '.join(sorted(ALLOWED_IDEF0_ROLES))}"
        )
    return value


def _normalize_edge_kind(value: str | None, diagram_type: str, path: str) -> str | None:
    if value is None:
        if diagram_type == "uml_class":
            return "association"
        if diagram_type == "idef1x":
            return "non_identifying"
        return None
    if not isinstance(value, str) or not value.strip():
        raise ValidationError(f"{path}: expected non-empty string")
    value = value.strip()
    if diagram_type == "uml_class":
        if value not in ALLOWED_UML_EDGE_KINDS:
            raise ValidationError(
                f"{path}: UML edge kind must be one of {', '.join(sorted(ALLOWED_UML_EDGE_KINDS))}"
            )
        return value
    if diagram_type == "idef1x":
        if value not in ALLOWED_IDEF1X_EDGE_KINDS:
            raise ValidationError(
                f"{path}: IDEF1X edge kind must be one of {', '.join(sorted(ALLOWED_IDEF1X_EDGE_KINDS))}"
            )
        return value
    raise ValidationError(f"{path}: edge kind is only supported for uml_class/idef1x diagrams")


def _normalize_icom(value: str | None, diagram_type: str, path: str) -> str | None:
    if value is None:
        return None
    if diagram_type != "idef0":
        raise ValidationError(f"{path}: ICOM is only supported for idef0 diagrams")
    if not isinstance(value, str) or not re.fullmatch(r"[ICOM]\d+", value.strip()):
        raise ValidationError(f"{path}: ICOM must match I1/C1/O1/M1 style codes")
    return value.strip()


def _normalize_tunnel(value: object, diagram_type: str, path: str) -> bool:
    if value is None:
        return False
    if diagram_type != "idef0":
        raise ValidationError(f"{path}: tunnel markers are only supported for idef0 diagrams")
    if not isinstance(value, bool):
        raise ValidationError(f"{path}: expected boolean")
    return value


def _load_style(raw_style: object, path: str) -> Style | None:
    if raw_style is None:
        return None
    raw_style = _expect_type(raw_style, dict, path)
    assert isinstance(raw_style, dict)
    return Style(
        fill=_get_string(raw_style, "fill", path, required=False),
        stroke=_get_string(raw_style, "stroke", path, required=False),
        stroke_width=_get_number(raw_style, "stroke_width", path, required=False),
        text_color=_get_string(raw_style, "text_color", path, required=False),
        font_size=_get_number(raw_style, "font_size", path, required=False),
        font_weight=_get_string(raw_style, "font_weight", path, required=False),
        dasharray=_get_string(raw_style, "dasharray", path, required=False),
        label_fill=_get_string(raw_style, "label_fill", path, required=False),
        label_stroke=_get_string(raw_style, "label_stroke", path, required=False),
        corner_radius=_get_number(raw_style, "corner_radius", path, required=False),
        opacity=_get_number(raw_style, "opacity", path, required=False),
        arrow_scale=_get_number(raw_style, "arrow_scale", path, required=False),
    )


def _load_diagram_style(raw_style: object, path: str) -> DiagramStyle | None:
    if raw_style is None:
        return None
    raw_style = _expect_type(raw_style, dict, path)
    assert isinstance(raw_style, dict)
    return DiagramStyle(
        background_fill=_get_string(raw_style, "background_fill", path, required=False),
        node=_load_style(raw_style.get("node"), f"{path}.node") or Style(),
        edge=_load_style(raw_style.get("edge"), f"{path}.edge") or Style(),
        frame_fill=_get_string(raw_style, "frame_fill", path, required=False),
        frame_stroke=_get_string(raw_style, "frame_stroke", path, required=False),
    )


def _load_columns(raw_columns: object, path: str) -> list[Column]:
    if raw_columns is None:
        return []
    raw_columns = _expect_type(raw_columns, list, path)
    assert isinstance(raw_columns, list)
    columns: list[Column] = []
    for index, item in enumerate(raw_columns):
        item = _expect_type(item, dict, f"{path}[{index}]")
        assert isinstance(item, dict)
        key = _get_string(item, "key", f"{path}[{index}]", required=False)
        if key is not None and key not in ALLOWED_COLUMN_KEYS:
            raise ValidationError(
                f"{path}[{index}].key: must be one of {', '.join(sorted(ALLOWED_COLUMN_KEYS))}"
            )
        columns.append(
            Column(
                name=_get_string(item, "name", f"{path}[{index}]"),
                data_type=_get_string(item, "data_type", f"{path}[{index}]", required=False),
                key=key,
                nullable=_get_bool(item, "nullable", f"{path}[{index}]", required=False),
                default=_get_string(item, "default", f"{path}[{index}]", required=False),
            )
        )
    return columns


def _load_node(raw_node: object, diagram_type: str, path: str) -> Node:
    raw_node = _expect_type(raw_node, dict, path)
    assert isinstance(raw_node, dict)

    raw_kind = _get_string(raw_node, "kind", path)
    assert raw_kind is not None

    return Node(
        id=_get_string(raw_node, "id", path),
        label=_get_string(raw_node, "label", path),
        kind=_normalize_kind(diagram_type, raw_kind),
        row=_get_int(raw_node, "row", path),
        column=_get_int(raw_node, "column", path),
        x=_get_int(raw_node, "x", path, required=False),
        y=_get_int(raw_node, "y", path, required=False),
        width=_get_int(raw_node, "width", path, required=False),
        height=_get_int(raw_node, "height", path, required=False),
        code=_get_string(raw_node, "code", path, required=False),
        decomposes_to=_get_string(raw_node, "decomposes_to", path, required=False),
        stereotype=_get_string(raw_node, "stereotype", path, required=False),
        attributes=_get_string_list(raw_node, "attributes", path),
        operations=_get_string_list(raw_node, "operations", path),
        columns=_load_columns(raw_node.get("columns"), f"{path}.columns"),
        style=_load_style(raw_node.get("style"), f"{path}.style"),
    )


def _load_edge(raw_edge: object, diagram_type: str, path: str) -> Edge:
    raw_edge = _expect_type(raw_edge, dict, path)
    assert isinstance(raw_edge, dict)

    source = raw_edge.get("source")
    target = raw_edge.get("target")

    if source is not None and (not isinstance(source, str) or not source.strip()):
        raise ValidationError(f"{path}.source: expected non-empty string or null")
    if target is not None and (not isinstance(target, str) or not target.strip()):
        raise ValidationError(f"{path}.target: expected non-empty string or null")

    return Edge(
        id=_get_string(raw_edge, "id", path),
        source=source.strip() if isinstance(source, str) else None,
        target=target.strip() if isinstance(target, str) else None,
        label=_get_string(raw_edge, "label", path, required=False),
        role=_normalize_role(raw_edge.get("role"), diagram_type, f"{path}.role"),
        kind=_normalize_edge_kind(raw_edge.get("kind"), diagram_type, f"{path}.kind"),
        icom=_normalize_icom(raw_edge.get("icom"), diagram_type, f"{path}.icom"),
        source_side=_normalize_side(raw_edge.get("source_side"), f"{path}.source_side"),
        target_side=_normalize_side(raw_edge.get("target_side"), f"{path}.target_side"),
        tunnel_source=_normalize_tunnel(raw_edge.get("tunnel_source"), diagram_type, f"{path}.tunnel_source"),
        tunnel_target=_normalize_tunnel(raw_edge.get("tunnel_target"), diagram_type, f"{path}.tunnel_target"),
        source_label=_get_string(raw_edge, "source_label", path, required=False),
        target_label=_get_string(raw_edge, "target_label", path, required=False),
        style=_load_style(raw_edge.get("style"), f"{path}.style"),
    )


def _load_frame(raw_frame: object, diagram_type: str, path: str) -> Idef0Frame | None:
    if raw_frame is None:
        return Idef0Frame() if diagram_type in {"idef0", "idef3"} else None

    if diagram_type not in {"idef0", "idef3"}:
        raise ValidationError(f"{path}: frame metadata is only supported for idef0/idef3 diagrams")

    raw_frame = _expect_type(raw_frame, dict, path)
    assert isinstance(raw_frame, dict)

    return Idef0Frame(
        enabled=_get_bool(raw_frame, "enabled", path, required=False)
        if "enabled" in raw_frame
        else True,
        paper_fill=_get_string(raw_frame, "paper_fill", path, required=False),
        used_at=_get_string(raw_frame, "used_at", path, required=False),
        author=_get_string(raw_frame, "author", path, required=False),
        project=_get_string(raw_frame, "project", path, required=False),
        date=_get_string(raw_frame, "date", path, required=False),
        revision=_get_string(raw_frame, "revision", path, required=False),
        status=_get_string(raw_frame, "status", path, required=False),
        reader=_get_string(raw_frame, "reader", path, required=False),
        context=_get_string(raw_frame, "context", path, required=False),
        notes=_get_string(raw_frame, "notes", path, required=False),
        page=_get_string(raw_frame, "page", path, required=False),
        node_ref=_get_string(raw_frame, "node_ref", path, required=False),
    )


def _load_diagram(raw_diagram: object, path: str) -> Diagram:
    raw_diagram = _expect_type(raw_diagram, dict, path)
    assert isinstance(raw_diagram, dict)

    diagram_type = _get_string(raw_diagram, "type", path)
    assert diagram_type is not None
    if diagram_type not in ALLOWED_DIAGRAM_TYPES:
        raise ValidationError(
            f"{path}.type: unsupported diagram type '{diagram_type}'. "
            f"Allowed: {', '.join(sorted(ALLOWED_DIAGRAM_TYPES))}"
        )

    raw_nodes = _expect_type(raw_diagram.get("nodes"), list, f"{path}.nodes")
    raw_edges = _expect_type(raw_diagram.get("edges"), list, f"{path}.edges")
    assert isinstance(raw_nodes, list)
    assert isinstance(raw_edges, list)

    nodes = [_load_node(item, diagram_type, f"{path}.nodes[{index}]") for index, item in enumerate(raw_nodes)]
    edges = [_load_edge(item, diagram_type, f"{path}.edges[{index}]") for index, item in enumerate(raw_edges)]

    diagram = Diagram(
        id=_get_string(raw_diagram, "id", path),
        title=_get_string(raw_diagram, "title", path),
        type=diagram_type,
        nodes=nodes,
        edges=edges,
        description=_get_string(raw_diagram, "description", path, required=False),
        frame=_load_frame(raw_diagram.get("frame"), diagram_type, f"{path}.frame"),
        style=_load_diagram_style(raw_diagram.get("style"), f"{path}.style"),
    )
    _validate_diagram(diagram, path)
    return diagram


def _validate_style(style: Style, path: str) -> None:
    if style.stroke_width is not None and style.stroke_width <= 0:
        raise ValidationError(f"{path}.stroke_width: must be > 0")
    if style.font_size is not None and style.font_size <= 0:
        raise ValidationError(f"{path}.font_size: must be > 0")
    if style.corner_radius is not None and style.corner_radius < 0:
        raise ValidationError(f"{path}.corner_radius: must be >= 0")
    if style.arrow_scale is not None and style.arrow_scale <= 0:
        raise ValidationError(f"{path}.arrow_scale: must be > 0")
    if style.opacity is not None and not (0 <= style.opacity <= 1):
        raise ValidationError(f"{path}.opacity: must be between 0 and 1")


def _validate_diagram(diagram: Diagram, path: str) -> None:
    if not diagram.nodes:
        raise ValidationError(f"{path}.nodes: at least one node is required")

    if diagram.type == "idef0":
        node_count = len(diagram.nodes)
        if node_count != 1 and not (3 <= node_count <= 6):
            raise ValidationError(
                f"{path}.nodes: IDEF0 diagrams must have 1 context box or between 3 and 6 child boxes"
            )

    if diagram.style is not None:
        _validate_style(diagram.style.node, f"{path}.style.node")
        _validate_style(diagram.style.edge, f"{path}.style.edge")

    seen_nodes: set[str] = set()
    for node in diagram.nodes:
        if node.id in seen_nodes:
            raise ValidationError(f"{path}.nodes: duplicate node id '{node.id}'")
        seen_nodes.add(node.id)
        if node.width is not None and node.width <= 0:
            raise ValidationError(f"{path}.nodes[{node.id}].width: must be > 0")
        if node.height is not None and node.height <= 0:
            raise ValidationError(f"{path}.nodes[{node.id}].height: must be > 0")
        if node.style is not None:
            _validate_style(node.style, f"{path}.nodes[{node.id}].style")

    seen_edges: set[str] = set()
    for edge in diagram.edges:
        if edge.id in seen_edges:
            raise ValidationError(f"{path}.edges: duplicate edge id '{edge.id}'")
        seen_edges.add(edge.id)

        if diagram.type in {"flowchart", "uml_class", "idef1x"}:
            if edge.source is None or edge.target is None:
                raise ValidationError(
                    f"{path}.edges[{edge.id}]: {diagram.type} edges require both source and target"
                )
        else:
            if edge.source is None and edge.target is None:
                raise ValidationError(f"{path}.edges[{edge.id}]: edge cannot have both ends null")

        if edge.source is not None and edge.source not in seen_nodes:
            raise ValidationError(f"{path}.edges[{edge.id}]: unknown source node '{edge.source}'")
        if edge.target is not None and edge.target not in seen_nodes:
            raise ValidationError(f"{path}.edges[{edge.id}]: unknown target node '{edge.target}'")
        if edge.style is not None:
            _validate_style(edge.style, f"{path}.edges[{edge.id}].style")

        if diagram.type == "idef0":
            is_boundary = edge.source is None or edge.target is None
            if edge.icom is not None and not is_boundary:
                raise ValidationError(f"{path}.edges[{edge.id}]: ICOM codes are only valid on boundary arrows")
            if len(diagram.nodes) > 1 and is_boundary:
                if edge.source is None and edge.icom is None and not edge.tunnel_source:
                    raise ValidationError(
                        f"{path}.edges[{edge.id}]: child IDEF0 boundary input/control/mechanism arrows need icom or tunnel_source"
                    )
                if edge.target is None and edge.icom is None and not edge.tunnel_target:
                    raise ValidationError(
                        f"{path}.edges[{edge.id}]: child IDEF0 boundary output arrows need icom or tunnel_target"
                    )


def load_document(data: dict[str, object]) -> Document:
    if not isinstance(data, dict):
        raise ValidationError("Document root must be an object")

    schema_version = _get_string(data, "schema_version", "document")
    title = _get_string(data, "title", "document")
    raw_diagrams = _expect_type(data.get("diagrams"), list, "document.diagrams")
    assert isinstance(raw_diagrams, list)

    diagrams = [_load_diagram(item, f"document.diagrams[{index}]") for index, item in enumerate(raw_diagrams)]
    document = Document(schema_version=schema_version, title=title, diagrams=diagrams)
    _validate_document(document)
    return document


def _validate_document(document: Document) -> None:
    if not document.diagrams:
        raise ValidationError("document.diagrams: at least one diagram is required")

    seen_diagrams: set[str] = set()
    for diagram in document.diagrams:
        if diagram.id in seen_diagrams:
            raise ValidationError(f"document.diagrams: duplicate diagram id '{diagram.id}'")
        seen_diagrams.add(diagram.id)

    diagram_ids = seen_diagrams
    for diagram in document.diagrams:
        for node in diagram.nodes:
            if node.decomposes_to is not None and node.decomposes_to not in diagram_ids:
                raise ValidationError(
                    f"document.diagrams[{diagram.id}].nodes[{node.id}].decomposes_to: "
                    f"unknown diagram id '{node.decomposes_to}'"
                )


def load_document_file(path: str | Path) -> Document:
    document_path = Path(path)
    data = json.loads(document_path.read_text(encoding="utf-8"))
    return load_document(data)
