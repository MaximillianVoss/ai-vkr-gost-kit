from __future__ import annotations

from dataclasses import dataclass
from heapq import heappop, heappush
from html import escape
import textwrap

from .models import Column, Diagram, Document, Edge, Idef0Frame, Node, Style


FONT_FAMILY = "Segoe UI, Arial, sans-serif"
TEXT_COLOR = "#1f2937"
STROKE_COLOR = "#0f172a"
FILL_COLOR = "#ffffff"
ACCENT_COLOR = "#2563eb"
IDEF_FILL = "#f8fafc"
IDEF_PAGE_FILL = "#ffffff"
FRAME_COLOR = "#64748b"

FLOW_GRID_X = 300
FLOW_GRID_Y = 176
FLOW_ORIGIN_X = 220
FLOW_ORIGIN_Y = 240

IDEF_GRID_X = 280
IDEF_GRID_Y = 200
IDEF_ORIGIN_X = 360
IDEF_ORIGIN_Y = 260
IDEF_FRAME_MARGIN = 18
IDEF_HEADER_HEIGHT = 86
IDEF_FOOTER_HEIGHT = 40
IDEF_LAYOUT_CENTER_X = 620
IDEF_LAYOUT_CENTER_Y = 330
IDEF_CHILD_BASE_X = 260
IDEF_CHILD_BASE_Y = 240
IDEF_CHILD_HORIZONTAL_GAP = 96
IDEF_CHILD_VERTICAL_GAP = 112
IDEF_STAIR_STEP = 44
ROUTE_MARGIN = 24
BEND_PENALTY = 28.0
PREFERRED_AXIS_WEIGHT = 0.18

FLOWCHART_DEFAULT_SIZES = {
    "terminator": (220, 80),
    "process": (230, 100),
    "decision": (230, 130),
    "data": (230, 100),
    "document": (230, 105),
    "predefined_process": (230, 100),
    "connector": (44, 44),
    "offpage": (164, 92),
}

IDEF0_DEFAULT_SIZE = (280, 120)
IDEF1X_DEFAULT_SIZE = (250, 132)
IDEF3_DEFAULT_SIZES = {
    "uob": (170, 92),
    "junction_x": (34, 34),
    "junction_and": (34, 34),
    "junction_or": (34, 34),
}
UML_CLASS_DEFAULT_SIZE = (240, 128)


@dataclass(slots=True)
class Box:
    x: float
    y: float
    width: float
    height: float

    @property
    def center_x(self) -> float:
        return self.x + self.width / 2

    @property
    def center_y(self) -> float:
        return self.y + self.height / 2


@dataclass(slots=True)
class Bounds:
    left: float
    top: float
    right: float
    bottom: float

    @property
    def width(self) -> float:
        return self.right - self.left

    @property
    def height(self) -> float:
        return self.bottom - self.top


    def contains_point(self, point: tuple[float, float]) -> bool:
        return self.left <= point[0] <= self.right and self.top <= point[1] <= self.bottom


@dataclass(slots=True)
class RoutePreferences:
    preferred_xs: tuple[float, ...] = ()
    preferred_ys: tuple[float, ...] = ()
    bend_penalty: float = BEND_PENALTY


@dataclass(slots=True)
class IdefBoundaryGroup:
    direction: str
    side: str
    label: str | None
    icom: str | None
    edges: tuple[Edge, ...]


def _merge_style(base: Style | None, override: Style | None) -> Style:
    base = base or Style()
    override = override or Style()
    return Style(
        fill=override.fill if override.fill is not None else base.fill,
        stroke=override.stroke if override.stroke is not None else base.stroke,
        stroke_width=override.stroke_width if override.stroke_width is not None else base.stroke_width,
        text_color=override.text_color if override.text_color is not None else base.text_color,
        font_size=override.font_size if override.font_size is not None else base.font_size,
        font_weight=override.font_weight if override.font_weight is not None else base.font_weight,
        dasharray=override.dasharray if override.dasharray is not None else base.dasharray,
        label_fill=override.label_fill if override.label_fill is not None else base.label_fill,
        label_stroke=override.label_stroke if override.label_stroke is not None else base.label_stroke,
        corner_radius=override.corner_radius if override.corner_radius is not None else base.corner_radius,
        opacity=override.opacity if override.opacity is not None else base.opacity,
        arrow_scale=override.arrow_scale if override.arrow_scale is not None else base.arrow_scale,
    )


def _diagram_background(diagram: Diagram, fallback: str) -> str:
    if diagram.style and diagram.style.background_fill:
        return diagram.style.background_fill
    if diagram.frame and diagram.frame.paper_fill:
        return diagram.frame.paper_fill
    return fallback


def _node_style(diagram: Diagram, node: Node, *, fill: str, stroke: str, stroke_width: float, font_size: float) -> Style:
    base = Style(
        fill=fill,
        stroke=stroke,
        stroke_width=stroke_width,
        text_color=TEXT_COLOR,
        font_size=font_size,
        font_weight="400",
        corner_radius=0,
        opacity=1.0,
    )
    if diagram.style:
        base = _merge_style(base, diagram.style.node)
    return _merge_style(base, node.style)


def _edge_style(diagram: Diagram, edge: Edge, *, stroke: str, stroke_width: float, text_color: str) -> Style:
    base = Style(
        stroke=stroke,
        stroke_width=stroke_width,
        text_color=text_color,
        font_size=11,
        font_weight="600",
        label_fill="#ffffff",
        label_stroke="#e2e8f0",
        opacity=1.0,
        arrow_scale=1.0,
    )
    if diagram.style:
        base = _merge_style(base, diagram.style.edge)
    return _merge_style(base, edge.style)


def render_svg(document: Document, diagram_id: str | None = None) -> str:
    if diagram_id is None:
        diagram = document.diagrams[0]
    else:
        diagram = next((item for item in document.diagrams if item.id == diagram_id), None)
        if diagram is None:
            raise ValueError(f"Diagram '{diagram_id}' not found")
    return render_diagram_svg(diagram, document.title)


def render_diagram_svg(diagram: Diagram, document_title: str) -> str:
    if diagram.type == "idef0":
        return _render_idef0_diagram_svg(diagram, document_title)
    if diagram.type == "idef1x":
        return _render_idef1x_diagram_svg(diagram, document_title)
    if diagram.type == "idef3":
        return _render_idef3_diagram_svg(diagram, document_title)
    if diagram.type == "uml_class":
        return _render_uml_class_diagram_svg(diagram, document_title)
    return _render_flowchart_svg(diagram, document_title)


def _render_flowchart_svg(diagram: Diagram, document_title: str) -> str:
    boxes = _layout_nodes(
        diagram,
        lane_origin_x=FLOW_ORIGIN_X,
        lane_origin_y=FLOW_ORIGIN_Y,
        grid_x=FLOW_GRID_X,
        grid_y=FLOW_GRID_Y,
    )
    width, height = _flow_canvas_size(boxes)
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}">',
        f'<rect x="0" y="0" width="{width}" height="{height}" fill="{_diagram_background(diagram, "#ffffff")}" />',
        _simple_title(document_title, diagram.title, diagram.description, width),
    ]

    for edge in diagram.edges:
        points = _flowchart_edge_points(edge, boxes)
        parts.extend(
            _render_edge(
                points,
                edge.label,
                _edge_style(diagram, edge, stroke=ACCENT_COLOR, stroke_width=1.7, text_color=ACCENT_COLOR),
                end_marker="arrow",
                source_label=edge.source_label,
                target_label=edge.target_label,
            )
        )

    for node in diagram.nodes:
        parts.extend(_render_flowchart_node(diagram, node, boxes[node.id]))

    parts.append("</svg>")
    return "\n".join(parts)


def _render_idef0_diagram_svg(diagram: Diagram, document_title: str) -> str:
    boxes = _layout_nodes(
        diagram,
        lane_origin_x=IDEF_ORIGIN_X,
        lane_origin_y=IDEF_ORIGIN_Y,
        grid_x=IDEF_GRID_X,
        grid_y=IDEF_GRID_Y,
    )
    frame = _resolve_idef_frame(diagram, document_title)
    page_bounds, content_bounds = _idef_page_bounds(boxes)
    _recenter_single_idef_box(diagram, boxes, content_bounds)
    page_bounds, content_bounds = _idef_page_bounds(boxes)
    width = int(page_bounds.right)
    height = int(page_bounds.bottom)
    edge_sides, anchor_map = _build_idef_anchor_map(diagram, boxes)
    grouped_boundary_edges, grouped_edge_ids = _build_idef_boundary_groups(diagram, edge_sides)
    frame_fill = diagram.style.frame_fill if diagram.style and diagram.style.frame_fill else "none"
    frame_stroke = diagram.style.frame_stroke if diagram.style and diagram.style.frame_stroke else FRAME_COLOR

    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}">',
        f'<rect x="0" y="0" width="{width}" height="{height}" fill="{_diagram_background(diagram, IDEF_PAGE_FILL)}" />',
        f'<rect x="{IDEF_FRAME_MARGIN}" y="{IDEF_FRAME_MARGIN}" '
        f'width="{page_bounds.width - IDEF_FRAME_MARGIN * 2}" height="{page_bounds.height - IDEF_FRAME_MARGIN * 2}" '
        f'fill="{frame_fill}" stroke="{frame_stroke}" stroke-width="1.4" />',
    ]

    parts.extend(_render_idef_frame(frame, diagram, page_bounds))

    for group in grouped_boundary_edges:
        parts.extend(
            _render_idef_boundary_group(
                diagram,
                group,
                boxes=boxes,
                content_bounds=content_bounds,
                edge_sides=edge_sides,
                anchor_map=anchor_map,
            )
        )

    for edge in diagram.edges:
        if edge.id in grouped_edge_ids:
            continue
        edge_style = _edge_style(diagram, edge, stroke="#111827", stroke_width=1.7, text_color="#111827")
        points = _idef0_edge_points(
            edge,
            boxes=boxes,
            content_bounds=content_bounds,
            edge_sides=edge_sides,
            anchor_map=anchor_map,
        )
        parts.extend(
            _render_edge(
                points,
                edge.label,
                edge_style,
                end_marker="arrow",
                source_label=edge.source_label,
                target_label=edge.target_label,
            )
        )
        parts.extend(_render_idef_boundary_annotations(edge, points, edge_style))

    for node in diagram.nodes:
        parts.extend(_render_idef0_node(diagram, node, boxes[node.id]))

    parts.append("</svg>")
    return "\n".join(parts)


def _render_idef3_diagram_svg(diagram: Diagram, document_title: str) -> str:
    boxes = _layout_nodes(
        diagram,
        lane_origin_x=IDEF_ORIGIN_X,
        lane_origin_y=IDEF_ORIGIN_Y,
        grid_x=IDEF_GRID_X,
        grid_y=IDEF_GRID_Y,
    )
    frame = _resolve_idef_frame(diagram, document_title)
    page_bounds, content_bounds = _idef_page_bounds(boxes)
    width = int(page_bounds.right)
    height = int(page_bounds.bottom)
    obstacles = _expanded_obstacles(boxes, margin=14)
    frame_fill = diagram.style.frame_fill if diagram.style and diagram.style.frame_fill else "none"
    frame_stroke = diagram.style.frame_stroke if diagram.style and diagram.style.frame_stroke else FRAME_COLOR

    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}">',
        f'<rect x="0" y="0" width="{width}" height="{height}" fill="{_diagram_background(diagram, IDEF_PAGE_FILL)}" />',
        f'<rect x="{IDEF_FRAME_MARGIN}" y="{IDEF_FRAME_MARGIN}" '
        f'width="{page_bounds.width - IDEF_FRAME_MARGIN * 2}" height="{page_bounds.height - IDEF_FRAME_MARGIN * 2}" '
        f'fill="{frame_fill}" stroke="{frame_stroke}" stroke-width="1.4" />',
    ]

    parts.extend(_render_idef_frame(frame, diagram, page_bounds))
    parts.append(
        f'<rect x="{content_bounds.left}" y="{content_bounds.top}" width="{content_bounds.width}" '
        f'height="{content_bounds.height}" rx="12" ry="12" fill="none" stroke="#9ca3af" stroke-width="1" />'
    )

    for edge in diagram.edges:
        if edge.source is None and edge.target is not None:
            target_side = edge.target_side or "left"
            end = _anchor_point(boxes[edge.target], target_side)
            start = _frame_boundary_point(content_bounds, target_side, end)
            points = _route_with_obstacles(
                start,
                end,
                target_side,
                target_side,
                obstacles,
                end_exclude=edge.target,
                content_bounds=content_bounds,
            )
        elif edge.target is None and edge.source is not None:
            source_side = edge.source_side or "right"
            start = _anchor_point(boxes[edge.source], source_side)
            end = _frame_boundary_point(content_bounds, source_side, start)
            points = _route_with_obstacles(
                start,
                end,
                source_side,
                source_side,
                obstacles,
                start_exclude=edge.source,
                content_bounds=content_bounds,
            )
        else:
            assert edge.source is not None and edge.target is not None
            source_box = boxes[edge.source]
            target_box = boxes[edge.target]
            source_side, target_side = _resolve_relative_sides(
                source_box,
                target_box,
                edge.source_side,
                edge.target_side,
            )
            start = _anchor_point(source_box, source_side)
            end = _anchor_point(target_box, target_side)
            points = _route_with_obstacles(
                start,
                end,
                source_side,
                target_side,
                obstacles,
                start_exclude=edge.source,
                end_exclude=edge.target,
                content_bounds=content_bounds,
            )
        parts.extend(
            _render_edge(
                points,
                edge.label,
                _edge_style(diagram, edge, stroke="#111827", stroke_width=1.7, text_color="#111827"),
                end_marker="arrow",
                source_label=edge.source_label,
                target_label=edge.target_label,
            )
        )

    for node in diagram.nodes:
        parts.extend(_render_idef3_node(diagram, node, boxes[node.id]))

    parts.append("</svg>")
    return "\n".join(parts)


def _render_uml_class_diagram_svg(diagram: Diagram, document_title: str) -> str:
    boxes = _layout_nodes(
        diagram,
        lane_origin_x=FLOW_ORIGIN_X,
        lane_origin_y=FLOW_ORIGIN_Y,
        grid_x=FLOW_GRID_X,
        grid_y=FLOW_GRID_Y,
    )
    width, height = _flow_canvas_size(boxes)
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}">',
        f'<rect x="0" y="0" width="{width}" height="{height}" fill="{_diagram_background(diagram, "#ffffff")}" />',
        _simple_title(document_title, diagram.title, diagram.description, width),
    ]
    for edge in diagram.edges:
        points = _uml_edge_points(edge, boxes)
        start_marker, end_marker = _uml_markers(edge)
        edge_style = _edge_style(diagram, edge, stroke=STROKE_COLOR, stroke_width=1.6, text_color=STROKE_COLOR)
        if edge.kind in {"realization", "dependency"} and not edge_style.dasharray:
            edge_style = _merge_style(edge_style, Style(dasharray="7 4"))
        parts.extend(
            _render_edge(
                points,
                edge.label,
                edge_style,
                start_marker=start_marker,
                end_marker=end_marker,
                source_label=edge.source_label,
                target_label=edge.target_label,
            )
        )
    for node in diagram.nodes:
        parts.extend(_render_uml_class_node(diagram, node, boxes[node.id]))
    parts.append("</svg>")
    return "\n".join(parts)


def _render_idef1x_diagram_svg(diagram: Diagram, document_title: str) -> str:
    boxes = _layout_nodes(
        diagram,
        lane_origin_x=FLOW_ORIGIN_X,
        lane_origin_y=FLOW_ORIGIN_Y,
        grid_x=FLOW_GRID_X,
        grid_y=FLOW_GRID_Y,
    )
    width, height = _flow_canvas_size(boxes)
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
        f'viewBox="0 0 {width} {height}">',
        f'<rect x="0" y="0" width="{width}" height="{height}" fill="{_diagram_background(diagram, "#ffffff")}" />',
        _simple_title(document_title, diagram.title, diagram.description, width),
    ]
    for edge in diagram.edges:
        points = _idef1x_edge_points(edge, boxes)
        edge_style = _edge_style(diagram, edge, stroke=STROKE_COLOR, stroke_width=1.5, text_color=STROKE_COLOR)
        if edge.kind == "non_identifying" and not edge_style.dasharray:
            edge_style = _merge_style(edge_style, Style(dasharray="8 4"))
        parts.extend(
            _render_edge(
                points,
                edge.label,
                edge_style,
                start_marker="none",
                end_marker="none",
                source_label=edge.source_label,
                target_label=edge.target_label,
            )
        )
    for node in diagram.nodes:
        parts.extend(_render_idef1x_node(diagram, node, boxes[node.id]))
    parts.append("</svg>")
    return "\n".join(parts)


def _marker_def(marker_id: str, color: str) -> str:
    return (
        f'<marker id="{marker_id}" markerWidth="10" markerHeight="10" refX="9" refY="3" '
        'orient="auto" markerUnits="strokeWidth">'
        f'<path d="M0,0 L0,6 L9,3 z" fill="{color}" />'
        "</marker>"
    )


def _simple_title(document_title: str, title: str, description: str | None, width: int) -> str:
    parts = [
        f'<text x="{width / 2:.1f}" y="42" text-anchor="middle" font-size="26" font-weight="700" '
        f'font-family="{FONT_FAMILY}" fill="{TEXT_COLOR}">{escape(document_title)}</text>',
        f'<text x="{width / 2:.1f}" y="74" text-anchor="middle" font-size="18" font-weight="600" '
        f'font-family="{FONT_FAMILY}" fill="{TEXT_COLOR}">{escape(title)}</text>',
    ]
    if description:
        parts.append(
            f'<text x="{width / 2:.1f}" y="100" text-anchor="middle" font-size="13" '
            f'font-family="{FONT_FAMILY}" fill="#475569">{escape(description)}</text>'
        )
    return "\n".join(parts)


def _layout_nodes(
    diagram: Diagram,
    lane_origin_x: int,
    lane_origin_y: int,
    grid_x: int,
    grid_y: int,
) -> dict[str, Box]:
    if diagram.type == "flowchart":
        return _layout_grid_nodes(
            diagram,
            lane_origin_x=lane_origin_x,
            lane_origin_y=lane_origin_y,
            grid_x=grid_x,
            grid_y=grid_y,
        )
    if diagram.type == "idef0":
        return _layout_idef0_nodes(diagram)
    return _layout_grid_nodes(
        diagram,
        lane_origin_x=lane_origin_x,
        lane_origin_y=lane_origin_y,
        grid_x=grid_x,
        grid_y=grid_y,
    )


def _layout_grid_nodes(
    diagram: Diagram,
    lane_origin_x: int,
    lane_origin_y: int,
    grid_x: int,
    grid_y: int,
) -> dict[str, Box]:
    boxes: dict[str, Box] = {}
    for node in diagram.nodes:
        default_width, default_height = _default_node_size(diagram, node)
        width = node.width or default_width
        height = node.height or default_height
        center_x = node.x if node.x is not None else lane_origin_x + node.column * grid_x
        center_y = node.y if node.y is not None else lane_origin_y + node.row * grid_y
        boxes[node.id] = Box(
            x=center_x - width / 2,
            y=center_y - height / 2,
            width=width,
            height=height,
        )
    return boxes


def _layout_idef0_nodes(diagram: Diagram) -> dict[str, Box]:
    boxes: dict[str, Box] = {}
    auto_nodes: list[tuple[Node, float, float, float, float]] = []
    auto_sizes: dict[str, tuple[float, float]] = {}

    sorted_columns = sorted({node.column for node in diagram.nodes})
    sorted_rows = sorted({node.row for node in diagram.nodes})
    column_index = {value: index for index, value in enumerate(sorted_columns)}
    row_index = {value: index for index, value in enumerate(sorted_rows)}

    ordered_nodes = sorted(diagram.nodes, key=lambda item: (item.row, item.column, item.id))
    row_members: dict[int, list[Node]] = {}
    for node in ordered_nodes:
        row_members.setdefault(node.row, []).append(node)

    column_widths = [0.0 for _ in sorted_columns]
    row_heights = [0.0 for _ in sorted_rows]
    for node in ordered_nodes:
        default_width, default_height = _default_node_size(diagram, node)
        width = float(node.width or default_width)
        height = float(node.height or default_height)
        auto_sizes[node.id] = (width, height)
        if node.x is not None or node.y is not None:
            continue
        col = column_index[node.column]
        row = row_index[node.row]
        column_widths[col] = max(column_widths[col], width)
        row_heights[row] = max(row_heights[row], height)

    column_centers: list[float] = []
    for index, width in enumerate(column_widths):
        if index == 0:
            column_centers.append(float(IDEF_CHILD_BASE_X))
            continue
        previous_center = column_centers[index - 1]
        previous_width = column_widths[index - 1]
        distance = previous_width / 2 + width / 2 + IDEF_CHILD_HORIZONTAL_GAP
        column_centers.append(previous_center + distance)

    row_centers: list[float] = []
    for index, height in enumerate(row_heights):
        if index == 0:
            row_centers.append(float(IDEF_CHILD_BASE_Y))
            continue
        previous_center = row_centers[index - 1]
        previous_height = row_heights[index - 1]
        distance = previous_height / 2 + height / 2 + IDEF_CHILD_VERTICAL_GAP
        row_centers.append(previous_center + distance)

    for node in ordered_nodes:
        width, height = auto_sizes[node.id]
        if node.x is not None or node.y is not None:
            center_x = float(node.x if node.x is not None else IDEF_LAYOUT_CENTER_X)
            center_y = float(node.y if node.y is not None else IDEF_LAYOUT_CENTER_Y)
        else:
            col = column_index[node.column]
            row = row_index[node.row]
            row_order = row_members[node.row].index(node)
            center_x = column_centers[col]
            center_y = row_centers[row] + row_order * IDEF_STAIR_STEP
            auto_nodes.append((node, center_x, center_y, width, height))
            continue

        boxes[node.id] = Box(
            x=center_x - width / 2,
            y=center_y - height / 2,
            width=width,
            height=height,
        )

    if len(diagram.nodes) == 1:
        node = diagram.nodes[0]
        default_width, default_height = _default_node_size(diagram, node)
        width = node.width or default_width
        height = node.height or default_height
        if node.id not in boxes:
            boxes[node.id] = Box(
                x=IDEF_LAYOUT_CENTER_X - width / 2,
                y=IDEF_LAYOUT_CENTER_Y - height / 2,
                width=width,
                height=height,
            )
        return boxes

    if auto_nodes:
        centers = [item[1] for item in auto_nodes]
        auto_group_center = sum(centers) / len(centers)
        shift_x = IDEF_LAYOUT_CENTER_X - auto_group_center
        for node, center_x, center_y, width, height in auto_nodes:
            boxes[node.id] = Box(
                x=center_x + shift_x - width / 2,
                y=center_y - height / 2,
                width=width,
                height=height,
            )

    return boxes


def _default_node_size(diagram: Diagram, node: Node) -> tuple[int, int]:
    if diagram.type == "flowchart":
        return FLOWCHART_DEFAULT_SIZES[node.kind]
    if diagram.type == "idef0":
        return IDEF0_DEFAULT_SIZE
    if diagram.type == "idef1x":
        return _idef1x_node_size(node)
    if diagram.type == "uml_class":
        return _uml_class_node_size(node)
    return IDEF3_DEFAULT_SIZES[node.kind]


def _uml_class_node_size(node: Node) -> tuple[int, int]:
    style = node.style or Style()
    font_size = style.font_size or 13
    widest = max(
        [_approx_text_width(node.label, font_size + 1)]
        + [_approx_text_width(item, font_size) for item in node.attributes]
        + [_approx_text_width(item, font_size) for item in node.operations]
        + [_approx_text_width(node.stereotype, font_size - 1) if node.stereotype else 0.0]
    )
    width = int(max(UML_CLASS_DEFAULT_SIZE[0], widest + 42))
    header_lines = 1 + (1 if node.stereotype or node.kind in {"interface", "enum"} else 0)
    section_lines = max(1, len(node.attributes)) + max(1, len(node.operations))
    height = int(max(UML_CLASS_DEFAULT_SIZE[1], 26 + header_lines * 18 + section_lines * 17))
    return width, height


def _idef1x_node_size(node: Node) -> tuple[int, int]:
    font_size = (node.style.font_size if node.style and node.style.font_size else 12)
    widest = max(
        [_approx_text_width(node.label, font_size + 1)]
        + [_approx_text_width(_format_column_line(column), font_size) for column in node.columns]
    )
    width = int(max(IDEF1X_DEFAULT_SIZE[0], widest + 38))
    lines = max(1, len(node.columns))
    height = int(max(IDEF1X_DEFAULT_SIZE[1], 38 + lines * 18))
    return width, height


def _flow_canvas_size(boxes: dict[str, Box]) -> tuple[int, int]:
    max_right = max(box.x + box.width for box in boxes.values())
    max_bottom = max(box.y + box.height for box in boxes.values())
    width = int(max(900, max_right + 150))
    height = int(max(700, max_bottom + 120))
    return width, height


def _idef_page_bounds(boxes: dict[str, Box]) -> tuple[Bounds, Bounds]:
    min_x = min(box.x for box in boxes.values())
    max_x = max(box.x + box.width for box in boxes.values())
    max_y = max(box.y + box.height for box in boxes.values())

    content_left = max(70.0, min_x - 180.0)
    content_top = float(IDEF_HEADER_HEIGHT + 34)
    content_right = max_x + 180.0
    content_bottom = max_y + 110.0
    page_width = max(960.0, content_right + IDEF_FRAME_MARGIN)
    page_height = max(700.0, content_bottom + IDEF_FOOTER_HEIGHT + 20.0)

    return (
        Bounds(left=0.0, top=0.0, right=page_width, bottom=page_height),
        Bounds(
            left=content_left,
            top=content_top,
            right=page_width - IDEF_FRAME_MARGIN,
            bottom=page_height - IDEF_FOOTER_HEIGHT - IDEF_FRAME_MARGIN,
        ),
    )


def _recenter_single_idef_box(diagram: Diagram, boxes: dict[str, Box], content_bounds: Bounds | None = None) -> None:
    if diagram.type != "idef0" or len(diagram.nodes) != 1:
        return
    node = diagram.nodes[0]
    if node.x is not None or node.y is not None:
        return
    box = boxes[node.id]
    if content_bounds is None:
        target_center_x = float(IDEF_LAYOUT_CENTER_X)
        target_center_y = float(IDEF_LAYOUT_CENTER_Y)
    else:
        target_center_x = content_bounds.left + content_bounds.width / 2
        target_center_y = content_bounds.top + content_bounds.height / 2
    boxes[node.id] = Box(
        x=target_center_x - box.width / 2,
        y=target_center_y - box.height / 2,
        width=box.width,
        height=box.height,
    )


def _render_flowchart_node(diagram: Diagram, node: Node, box: Box) -> list[str]:
    style = _node_style(diagram, node, fill=FILL_COLOR, stroke=STROKE_COLOR, stroke_width=2.5, font_size=16)
    parts: list[str] = []
    if node.kind == "terminator":
        parts.append(
            f'<rect x="{box.x}" y="{box.y}" width="{box.width}" height="{box.height}" '
            f'rx="{style.corner_radius if style.corner_radius is not None else box.height / 2}" '
            f'ry="{style.corner_radius if style.corner_radius is not None else box.height / 2}" '
            f'fill="{style.fill}" stroke="{style.stroke}" stroke-width="{style.stroke_width}"{_opacity_attr(style)}{_dash_attr(style)} />'
        )
        text_width = box.width - 36
    elif node.kind == "process":
        parts.append(
            f'<rect x="{box.x}" y="{box.y}" width="{box.width}" height="{box.height}" '
            f'rx="{style.corner_radius or 0}" ry="{style.corner_radius or 0}" '
            f'fill="{style.fill}" stroke="{style.stroke}" stroke-width="{style.stroke_width}"{_opacity_attr(style)}{_dash_attr(style)} />'
        )
        text_width = box.width - 28
    elif node.kind == "decision":
        points = [
            (box.center_x, box.y),
            (box.x + box.width, box.center_y),
            (box.center_x, box.y + box.height),
            (box.x, box.center_y),
        ]
        parts.append(
            f'<polygon points="{_points_to_svg(points)}" fill="{style.fill}" '
            f'stroke="{style.stroke}" stroke-width="{style.stroke_width}"{_opacity_attr(style)}{_dash_attr(style)} />'
        )
        text_width = box.width - 84
    elif node.kind == "data":
        skew = 26
        points = [
            (box.x + skew, box.y),
            (box.x + box.width, box.y),
            (box.x + box.width - skew, box.y + box.height),
            (box.x, box.y + box.height),
        ]
        parts.append(
            f'<polygon points="{_points_to_svg(points)}" fill="{style.fill}" '
            f'stroke="{style.stroke}" stroke-width="{style.stroke_width}"{_opacity_attr(style)}{_dash_attr(style)} />'
        )
        text_width = box.width - 42
    elif node.kind == "document":
        wave_y = box.y + box.height - 14
        path = (
            f"M {box.x} {box.y} "
            f"L {box.x + box.width} {box.y} "
            f"L {box.x + box.width} {wave_y} "
            f"Q {box.x + box.width * 0.75} {box.y + box.height + 8} {box.x + box.width * 0.5} {wave_y} "
            f"Q {box.x + box.width * 0.25} {box.y + box.height - 32} {box.x} {wave_y} Z"
        )
        parts.append(
            f'<path d="{path}" fill="{style.fill}" stroke="{style.stroke}" stroke-width="{style.stroke_width}"{_opacity_attr(style)}{_dash_attr(style)} />'
        )
        text_width = box.width - 32
    elif node.kind == "connector":
        radius = min(box.width, box.height) / 2
        parts.append(
            f'<circle cx="{box.center_x:.1f}" cy="{box.center_y:.1f}" r="{radius:.1f}" '
            f'fill="{style.fill}" stroke="{style.stroke}" stroke-width="{style.stroke_width}"{_opacity_attr(style)}{_dash_attr(style)} />'
        )
        text_width = box.width - 8
    elif node.kind == "offpage":
        inset = 18
        points = [
            (box.x, box.y),
            (box.x + box.width, box.y),
            (box.x + box.width, box.y + box.height - inset),
            (box.center_x, box.y + box.height),
            (box.x, box.y + box.height - inset),
        ]
        parts.append(
            f'<polygon points="{_points_to_svg(points)}" fill="{style.fill}" '
            f'stroke="{style.stroke}" stroke-width="{style.stroke_width}"{_opacity_attr(style)}{_dash_attr(style)} />'
        )
        text_width = box.width - 26
    else:
        parts.append(
            f'<rect x="{box.x}" y="{box.y}" width="{box.width}" height="{box.height}" '
            f'rx="{style.corner_radius or 0}" ry="{style.corner_radius or 0}" '
            f'fill="{style.fill}" stroke="{style.stroke}" stroke-width="{style.stroke_width}"{_opacity_attr(style)}{_dash_attr(style)} />'
        )
        left_inner = box.x + 14
        right_inner = box.x + box.width - 14
        parts.append(
            f'<line x1="{left_inner}" y1="{box.y}" x2="{left_inner}" y2="{box.y + box.height}" '
            f'stroke="{style.stroke}" stroke-width="{max(1.0, style.stroke_width or 2)}" />'
        )
        parts.append(
            f'<line x1="{right_inner}" y1="{box.y}" x2="{right_inner}" y2="{box.y + box.height}" '
            f'stroke="{style.stroke}" stroke-width="{max(1.0, style.stroke_width or 2)}" />'
        )
        text_width = box.width - 40

    parts.extend(
        _render_centered_text(
            node.label,
            box,
            font_size=int(style.font_size or 16),
            max_width=text_width,
            color=style.text_color or TEXT_COLOR,
            font_weight=style.font_weight or "400",
        )
    )
    if node.decomposes_to:
        parts.append(
            f'<text x="{box.x + box.width - 10}" y="{box.y + box.height - 10}" text-anchor="end" '
            f'font-size="11" font-family="{FONT_FAMILY}" fill="#64748b">детализация → {escape(node.decomposes_to)}</text>'
        )
    return parts


def _render_idef0_node(diagram: Diagram, node: Node, box: Box) -> list[str]:
    style = _node_style(diagram, node, fill=IDEF_FILL, stroke=STROKE_COLOR, stroke_width=1.8, font_size=15)
    parts = [
        f'<rect x="{box.x}" y="{box.y}" width="{box.width}" height="{box.height}" '
        f'rx="{style.corner_radius or 0}" ry="{style.corner_radius or 0}" '
        f'fill="{style.fill}" stroke="{style.stroke}" stroke-width="{style.stroke_width}"{_opacity_attr(style)}{_dash_attr(style)} />'
    ]
    parts.extend(
        _render_centered_text(
            node.label,
            box,
            font_size=int(style.font_size or 15),
            max_width=box.width - 34,
            color=style.text_color or TEXT_COLOR,
            font_weight=style.font_weight or "400",
        )
    )
    if node.code:
        parts.append(
            f'<text x="{box.x + box.width - 10}" y="{box.y + box.height - 10}" text-anchor="end" '
            f'font-size="12" font-weight="700" font-family="{FONT_FAMILY}" fill="#475569">{escape(node.code)}</text>'
        )
    if node.decomposes_to:
        parts.append(
            f'<text x="{box.x + 10}" y="{box.y + box.height - 10}" text-anchor="start" '
            f'font-size="10.5" font-family="{FONT_FAMILY}" fill="#64748b">→ {escape(node.decomposes_to)}</text>'
        )
    return parts


def _render_idef3_node(diagram: Diagram, node: Node, box: Box) -> list[str]:
    style = _node_style(diagram, node, fill=FILL_COLOR, stroke=STROKE_COLOR, stroke_width=1.8, font_size=12)
    if node.kind == "uob":
        shadow_x = box.x + 6
        shadow_y = box.y + 6
        footer_height = 18
        footer_cell = 28
        parts = [
            f'<rect x="{shadow_x}" y="{shadow_y}" width="{box.width}" height="{box.height}" '
            'fill="none" stroke="#cbd5e1" stroke-width="1" stroke-dasharray="2 2" />',
            f'<rect x="{box.x}" y="{box.y}" width="{box.width}" height="{box.height}" '
            f'rx="{style.corner_radius or 0}" ry="{style.corner_radius or 0}" '
            f'fill="{style.fill}" stroke="{style.stroke}" stroke-width="{style.stroke_width}"{_opacity_attr(style)}{_dash_attr(style)} />',
            f'<line x1="{box.x}" y1="{box.y + box.height - footer_height}" '
            f'x2="{box.x + box.width}" y2="{box.y + box.height - footer_height}" stroke="{style.stroke}" stroke-width="{max(1.0, (style.stroke_width or 1.8) - 0.6)}" />',
            f'<line x1="{box.x + footer_cell}" y1="{box.y + box.height - footer_height}" '
            f'x2="{box.x + footer_cell}" y2="{box.y + box.height}" stroke="{style.stroke}" stroke-width="{max(1.0, (style.stroke_width or 1.8) - 0.6)}" />',
        ]
        parts.extend(
            _render_centered_text(
                node.label,
                Box(box.x, box.y, box.width, box.height - footer_height),
                font_size=int(style.font_size or 12),
                max_width=box.width - 20,
                color=style.text_color or TEXT_COLOR,
                font_weight=style.font_weight or "400",
            )
        )
        if node.code:
            parts.append(
                f'<text x="{box.x + footer_cell / 2:.1f}" y="{box.y + box.height - 5:.1f}" text-anchor="middle" '
                f'font-size="10" font-family="{FONT_FAMILY}" font-weight="700" fill="{TEXT_COLOR}">{escape(node.code)}</text>'
            )
        return parts

    symbol = {
        "junction_x": "X",
        "junction_and": "&",
        "junction_or": "O",
    }[node.kind]
    parts = [
        f'<rect x="{box.x}" y="{box.y}" width="{box.width}" height="{box.height}" '
        f'rx="{style.corner_radius or 0}" ry="{style.corner_radius or 0}" '
        f'fill="{style.fill}" stroke="{style.stroke}" stroke-width="{style.stroke_width or 1.6}"{_opacity_attr(style)}{_dash_attr(style)} />',
        f'<text x="{box.center_x:.1f}" y="{box.center_y + 4:.1f}" text-anchor="middle" font-size="18" '
        f'font-family="{FONT_FAMILY}" font-weight="700" fill="{style.text_color or TEXT_COLOR}">{symbol}</text>',
    ]
    if node.label:
        parts.append(
            f'<text x="{box.center_x:.1f}" y="{box.y + box.height + 16:.1f}" text-anchor="middle" font-size="11" '
            f'font-family="{FONT_FAMILY}" font-weight="700" fill="{style.text_color or TEXT_COLOR}">{escape(node.label)}</text>'
        )
    return parts


def _render_uml_class_node(diagram: Diagram, node: Node, box: Box) -> list[str]:
    style = _node_style(diagram, node, fill="#ffffff", stroke=STROKE_COLOR, stroke_width=1.7, font_size=13)
    stroke = style.stroke or STROKE_COLOR
    text_color = style.text_color or TEXT_COLOR
    font_size = style.font_size or 13
    parts = [
        f'<rect x="{box.x}" y="{box.y}" width="{box.width}" height="{box.height}" '
        f'rx="{style.corner_radius or 4}" ry="{style.corner_radius or 4}" '
        f'fill="{style.fill}" stroke="{stroke}" stroke-width="{style.stroke_width}"{_opacity_attr(style)}{_dash_attr(style)} />'
    ]
    stereotype = node.stereotype
    if node.kind == "interface" and stereotype is None:
        stereotype = "interface"
    if node.kind == "enum" and stereotype is None:
        stereotype = "enumeration"

    header_lines: list[str] = []
    if stereotype:
        header_lines.append(f"≪{stereotype}≫")
    header_lines.append(node.label)
    header_height = 18 * len(header_lines) + 12
    attributes_top = box.y + header_height
    operations_top = box.y + box.height - max(28.0, len(node.operations) * 16 + 12)
    operations_top = max(attributes_top + 24.0, operations_top)

    parts.append(
        f'<line x1="{box.x}" y1="{attributes_top:.1f}" x2="{box.x + box.width}" y2="{attributes_top:.1f}" '
        f'stroke="{stroke}" stroke-width="{max(1.0, (style.stroke_width or 1.7) - 0.5)}" />'
    )
    parts.append(
        f'<line x1="{box.x}" y1="{operations_top:.1f}" x2="{box.x + box.width}" y2="{operations_top:.1f}" '
        f'stroke="{stroke}" stroke-width="{max(1.0, (style.stroke_width or 1.7) - 0.5)}" />'
    )

    for index, line in enumerate(header_lines):
        y = box.y + 20 + index * 16
        font_style = "italic" if node.kind == "abstract_class" and index == len(header_lines) - 1 else "normal"
        parts.append(
            f'<text x="{box.center_x:.1f}" y="{y:.1f}" text-anchor="middle" font-size="{font_size if index else max(11, font_size - 1)}" '
            f'font-family="{FONT_FAMILY}" font-style="{font_style}" font-weight="700" fill="{text_color}">{escape(line)}</text>'
        )

    parts.extend(
        _render_left_text_block(
            node.attributes or [" "],
            box.x + 10,
            attributes_top + 18,
            font_size=font_size,
            color=text_color,
        )
    )
    parts.extend(
        _render_left_text_block(
            node.operations or [" "],
            box.x + 10,
            operations_top + 18,
            font_size=font_size,
            color=text_color,
        )
    )
    return parts


def _render_idef1x_node(diagram: Diagram, node: Node, box: Box) -> list[str]:
    style = _node_style(diagram, node, fill="#ffffff", stroke=STROKE_COLOR, stroke_width=1.6, font_size=12)
    radius = style.corner_radius if style.corner_radius is not None else (10 if node.kind == "dependent_entity" else 2)
    parts = [
        f'<text x="{box.center_x:.1f}" y="{box.y - 8:.1f}" text-anchor="middle" font-size="{(style.font_size or 12) + 1}" '
        f'font-family="{FONT_FAMILY}" font-weight="700" fill="{style.text_color or TEXT_COLOR}">{escape(node.label)}</text>',
        f'<rect x="{box.x}" y="{box.y}" width="{box.width}" height="{box.height}" '
        f'rx="{radius}" ry="{radius}" fill="{style.fill}" stroke="{style.stroke}" stroke-width="{style.stroke_width}"{_opacity_attr(style)}{_dash_attr(style)} />',
    ]
    header_y = box.y + 24
    parts.append(
        f'<line x1="{box.x}" y1="{header_y:.1f}" x2="{box.x + box.width}" y2="{header_y:.1f}" '
        f'stroke="{style.stroke}" stroke-width="{max(1.0, (style.stroke_width or 1.6) - 0.4)}" />'
    )
    lines = [_format_column_line(column) for column in node.columns] or ["(атрибуты не заданы)"]
    parts.extend(
        _render_left_text_block(
            lines,
            box.x + 10,
            header_y + 18,
            font_size=style.font_size or 12,
            color=style.text_color or TEXT_COLOR,
        )
    )
    return parts


def _render_edge(
    points: list[tuple[float, float]],
    label: str | None,
    style: Style,
    *,
    start_marker: str = "none",
    end_marker: str = "none",
    source_label: str | None = None,
    target_label: str | None = None,
) -> list[str]:
    stroke = style.stroke or STROKE_COLOR
    parts = [
        f'<path d="{_polyline_path(points)}" fill="none" stroke="{stroke}" stroke-width="{style.stroke_width or 1.7}"'
        f'{_dash_attr(style)}{_opacity_attr(style)} />'
    ]
    if label:
        parts.extend(_render_edge_label(label, points, style))
    if source_label and len(points) >= 2:
        parts.extend(_render_endpoint_label(source_label, points[0], points[1], style, is_start=True))
    if target_label and len(points) >= 2:
        parts.extend(_render_endpoint_label(target_label, points[-1], points[-2], style, is_start=False))
    if start_marker != "none" and len(points) >= 2:
        parts.extend(_render_marker_shape(points[0], points[1], stroke, style, start_marker, is_start=True))
    if end_marker != "none" and len(points) >= 2:
        parts.extend(_render_marker_shape(points[-1], points[-2], stroke, style, end_marker, is_start=False))
    return parts


def _render_edge_label(label: str, points: list[tuple[float, float]], style: Style) -> list[str]:
    start, end = _best_label_segment(points)
    orientation = "horizontal" if abs(end[0] - start[0]) >= abs(end[1] - start[1]) else "vertical"
    font_size = int(style.font_size or 11)
    lines = _wrap_text(label, max_width=146, font_size=font_size)
    max_line_width = max(_approx_text_width(line, font_size) for line in lines)
    padding_x = 8
    padding_y = 5
    line_height = font_size * 1.25
    box_width = max_line_width + padding_x * 2
    box_height = len(lines) * line_height + padding_y * 2
    mid_x = (start[0] + end[0]) / 2
    mid_y = (start[1] + end[1]) / 2

    if orientation == "horizontal":
        box_x = mid_x - box_width / 2
        box_y = mid_y - box_height - 10
    else:
        box_x = mid_x + 10
        box_y = mid_y - box_height / 2

    label_fill = style.label_fill or "#ffffff"
    label_stroke = style.label_stroke or "#e2e8f0"
    label_color = style.text_color or style.stroke or TEXT_COLOR
    label_weight = style.font_weight or "600"
    parts = [
        f'<rect x="{box_x:.1f}" y="{box_y:.1f}" width="{box_width:.1f}" height="{box_height:.1f}" '
        f'rx="6" ry="6" fill="{label_fill}" fill-opacity="0.96" '
        f'stroke="{label_stroke}" stroke-width="0.8" />'
    ]
    text_x = box_x + box_width / 2
    first_line_y = box_y + padding_y + font_size
    for index, line in enumerate(lines):
        y = first_line_y + index * line_height
        parts.append(
            f'<text x="{text_x:.1f}" y="{y:.1f}" text-anchor="middle" font-size="{font_size}" '
            f'font-family="{FONT_FAMILY}" font-weight="{label_weight}" fill="{label_color}">{escape(line)}</text>'
        )
    return parts


def _opacity_attr(style: Style) -> str:
    if style.opacity is None or style.opacity >= 1:
        return ""
    return f' opacity="{style.opacity:.3f}"'


def _dash_attr(style: Style) -> str:
    if not style.dasharray:
        return ""
    return f' stroke-dasharray="{style.dasharray}"'


def _render_left_text_block(lines: list[str], x: float, start_y: float, *, font_size: float, color: str) -> list[str]:
    parts: list[str] = []
    line_height = font_size * 1.2
    for index, line in enumerate(lines):
        parts.append(
            f'<text x="{x:.1f}" y="{start_y + index * line_height:.1f}" text-anchor="start" font-size="{font_size}" '
            f'font-family="{FONT_FAMILY}" fill="{color}">{escape(line)}</text>'
        )
    return parts


def _render_endpoint_label(
    label: str,
    anchor: tuple[float, float],
    neighbor: tuple[float, float],
    style: Style,
    *,
    is_start: bool,
) -> list[str]:
    dx = neighbor[0] - anchor[0]
    dy = neighbor[1] - anchor[1]
    offset_x = -10 if dx > 0 else 10 if dx < 0 else 12
    offset_y = -10 if dy > 0 else 10 if dy < 0 else -8
    if is_start:
        offset_x *= -1
        offset_y *= -1
    x = anchor[0] + offset_x
    y = anchor[1] + offset_y
    anchor_mode = "end" if offset_x < 0 else "start"
    return [
        f'<text x="{x:.1f}" y="{y:.1f}" text-anchor="{anchor_mode}" font-size="{max(10, int((style.font_size or 11) - 1))}" '
        f'font-family="{FONT_FAMILY}" font-weight="600" fill="{style.text_color or style.stroke or TEXT_COLOR}">{escape(label)}</text>'
    ]


def _render_marker_shape(
    anchor: tuple[float, float],
    neighbor: tuple[float, float],
    stroke: str,
    style: Style,
    marker: str,
    *,
    is_start: bool,
) -> list[str]:
    if marker == "none":
        return []
    size = 10.0 * (style.arrow_scale or 1.0)
    ux, uy = _unit_vector(anchor, neighbor)
    if not is_start:
        ux, uy = -ux, -uy
    px, py = -uy, ux

    if marker in {"arrow", "triangle"}:
        base_x = anchor[0] + ux * size
        base_y = anchor[1] + uy * size
        points = [
            anchor,
            (base_x + px * size * 0.55, base_y + py * size * 0.55),
            (base_x - px * size * 0.55, base_y - py * size * 0.55),
        ]
        return [
            f'<polygon points="{_points_to_svg(points)}" fill="{stroke}" stroke="{stroke}" stroke-width="1" />'
        ]

    if marker == "triangle_hollow":
        base_x = anchor[0] + ux * size * 1.15
        base_y = anchor[1] + uy * size * 1.15
        points = [
            anchor,
            (base_x + px * size * 0.62, base_y + py * size * 0.62),
            (base_x - px * size * 0.62, base_y - py * size * 0.62),
        ]
        return [
            f'<polygon points="{_points_to_svg(points)}" fill="#ffffff" stroke="{stroke}" stroke-width="1.4" />'
        ]

    inner_x = anchor[0] + ux * size * 0.85
    inner_y = anchor[1] + uy * size * 0.85
    outer_x = anchor[0] + ux * size * 1.7
    outer_y = anchor[1] + uy * size * 1.7
    points = [
        anchor,
        (inner_x + px * size * 0.55, inner_y + py * size * 0.55),
        (outer_x, outer_y),
        (inner_x - px * size * 0.55, inner_y - py * size * 0.55),
    ]
    fill = stroke if marker == "diamond_filled" else "#ffffff"
    return [
        f'<polygon points="{_points_to_svg(points)}" fill="{fill}" stroke="{stroke}" stroke-width="1.3" />'
    ]


def _render_idef_boundary_annotations(edge: Edge, points: list[tuple[float, float]], style: Style) -> list[str]:
    if len(points) < 2:
        return []
    parts: list[str] = []
    if edge.icom:
        if edge.source is None:
            parts.extend(_render_icom_label(edge.icom, points[0], points[1], style))
        elif edge.target is None:
            parts.extend(_render_icom_label(edge.icom, points[-1], points[-2], style))
    if edge.tunnel_source:
        parts.extend(_render_tunnel_marker(points[0], points[1], style))
    if edge.tunnel_target:
        parts.extend(_render_tunnel_marker(points[-1], points[-2], style))
    return parts


def _render_icom_label(
    label: str,
    anchor: tuple[float, float],
    neighbor: tuple[float, float],
    style: Style,
) -> list[str]:
    ux, uy = _unit_vector(neighbor, anchor)
    x = anchor[0] + ux * 18
    y = anchor[1] + uy * 18 + 4
    if abs(ux) < 0.2:
        anchor_mode = "middle"
    else:
        anchor_mode = "start" if ux > 0 else "end"
    return [
        f'<text x="{x:.1f}" y="{y:.1f}" text-anchor="{anchor_mode}" font-size="{max(10, int((style.font_size or 11) - 0.5))}" '
        f'font-family="{FONT_FAMILY}" font-weight="700" fill="{style.text_color or style.stroke or TEXT_COLOR}">{escape(label)}</text>'
    ]


def _render_tunnel_marker(
    anchor: tuple[float, float],
    neighbor: tuple[float, float],
    style: Style,
) -> list[str]:
    dx = neighbor[0] - anchor[0]
    dy = neighbor[1] - anchor[1]
    angle = 0 if abs(dx) >= abs(dy) else 90
    color = style.stroke or STROKE_COLOR
    return [
        f'<text x="{anchor[0] - 6:.1f}" y="{anchor[1] + 4:.1f}" text-anchor="middle" '
        f'font-size="11" font-family="{FONT_FAMILY}" fill="{color}" transform="rotate({angle} {anchor[0] - 6:.1f} {anchor[1] + 4:.1f})">(</text>',
        f'<text x="{anchor[0] + 6:.1f}" y="{anchor[1] + 4:.1f}" text-anchor="middle" '
        f'font-size="11" font-family="{FONT_FAMILY}" fill="{color}" transform="rotate({angle} {anchor[0] + 6:.1f} {anchor[1] + 4:.1f})">)</text>',
    ]


def _unit_vector(anchor: tuple[float, float], neighbor: tuple[float, float]) -> tuple[float, float]:
    dx = neighbor[0] - anchor[0]
    dy = neighbor[1] - anchor[1]
    length = max(1.0, (dx * dx + dy * dy) ** 0.5)
    return dx / length, dy / length


def _format_column_line(column: Column) -> str:
    prefix = ""
    if column.key == "pk":
        prefix = "PK "
    elif column.key == "fk":
        prefix = "FK "
    elif column.key == "pk_fk":
        prefix = "PK/FK "
    elif column.key == "unique":
        prefix = "UQ "
    parts = [f"{prefix}{column.name}"]
    if column.data_type:
        parts.append(f": {column.data_type}")
    if column.nullable is False:
        parts.append(" NOT NULL")
    if column.default:
        parts.append(f" = {column.default}")
    return "".join(parts)


def _flowchart_edge_points(edge: Edge, boxes: dict[str, Box]) -> list[tuple[float, float]]:
    assert edge.source is not None
    assert edge.target is not None
    source_box = boxes[edge.source]
    target_box = boxes[edge.target]
    source_side, target_side = _resolve_flowchart_sides(edge, boxes)
    start = _anchor_point(source_box, source_side)
    end = _anchor_point(target_box, target_side)
    loopback_path = _flowchart_loopback_path(
        edge,
        boxes,
        start=start,
        end=end,
        source_side=source_side,
        target_side=target_side,
    )
    if loopback_path is not None:
        return loopback_path
    return _route_with_obstacles(
        start,
        end,
        source_side,
        target_side,
        _expanded_obstacles(boxes, margin=14),
        start_exclude=edge.source,
        end_exclude=edge.target,
    )


def _resolve_flowchart_sides(edge: Edge, boxes: dict[str, Box]) -> tuple[str, str]:
    assert edge.source is not None
    assert edge.target is not None
    source_box = boxes[edge.source]
    target_box = boxes[edge.target]
    if edge.source_side and edge.target_side:
        return edge.source_side, edge.target_side

    dx = target_box.center_x - source_box.center_x
    dy = target_box.center_y - source_box.center_y

    # For loopbacks in a top-down algorithm, route the return branch around the left rail by default.
    if dy < -max(source_box.height, target_box.height) * 0.45 and abs(dx) <= max(source_box.width, target_box.width) * 0.5:
        loop_side = edge.source_side or edge.target_side or "left"
        return edge.source_side or loop_side, edge.target_side or loop_side

    if dx < -max(source_box.width, target_box.width) * 0.45 and abs(dy) <= max(source_box.height, target_box.height) * 0.5:
        loop_side = edge.source_side or edge.target_side or "top"
        return edge.source_side or loop_side, edge.target_side or loop_side

    return _resolve_relative_sides(source_box, target_box, edge.source_side, edge.target_side)


def _flowchart_loopback_path(
    edge: Edge,
    boxes: dict[str, Box],
    *,
    start: tuple[float, float],
    end: tuple[float, float],
    source_side: str,
    target_side: str,
) -> list[tuple[float, float]] | None:
    assert edge.source is not None
    assert edge.target is not None
    source_box = boxes[edge.source]
    target_box = boxes[edge.target]
    obstacles = _expanded_obstacles(boxes, margin=14)
    start_exit = _shift_point(start, source_side, ROUTE_MARGIN)
    end_entry = _shift_point(end, target_side, ROUTE_MARGIN)

    if end[1] < start[1]:
        rail_x = _vertical_flow_loop_rail(boxes, source_box, target_box, source_side, target_side)
        candidate = _simplify_path([start, start_exit, (rail_x, start_exit[1]), (rail_x, end_entry[1]), end_entry, end])
        if _path_is_clear(candidate, obstacles, edge.source, edge.target):
            return candidate

    if end[0] < start[0] and abs(end[1] - start[1]) <= max(source_box.height, target_box.height) * 0.65:
        rail_y = _horizontal_flow_loop_rail(boxes, source_box, target_box, source_side, target_side)
        candidate = _simplify_path([start, start_exit, (start_exit[0], rail_y), (end_entry[0], rail_y), end_entry, end])
        if _path_is_clear(candidate, obstacles, edge.source, edge.target):
            return candidate

    return None


def _vertical_flow_loop_rail(
    boxes: dict[str, Box],
    source_box: Box,
    target_box: Box,
    source_side: str,
    target_side: str,
) -> float:
    band_top = min(source_box.y, target_box.y)
    band_bottom = max(source_box.y + source_box.height, target_box.y + target_box.height)
    relevant = [
        box
        for box in boxes.values()
        if not (box.y + box.height < band_top or box.y > band_bottom)
    ]
    left_rail = min(box.x for box in relevant) - ROUTE_MARGIN * 2
    right_rail = max(box.x + box.width for box in relevant) + ROUTE_MARGIN * 2
    if source_side == "right" or target_side == "right":
        return right_rail
    return left_rail


def _horizontal_flow_loop_rail(
    boxes: dict[str, Box],
    source_box: Box,
    target_box: Box,
    source_side: str,
    target_side: str,
) -> float:
    band_left = min(source_box.x, target_box.x)
    band_right = max(source_box.x + source_box.width, target_box.x + target_box.width)
    relevant = [
        box
        for box in boxes.values()
        if not (box.x + box.width < band_left or box.x > band_right)
    ]
    top_rail = min(box.y for box in relevant) - ROUTE_MARGIN * 2
    bottom_rail = max(box.y + box.height for box in relevant) + ROUTE_MARGIN * 2
    if source_side == "bottom" or target_side == "bottom":
        return bottom_rail
    return top_rail


def _uml_edge_points(edge: Edge, boxes: dict[str, Box]) -> list[tuple[float, float]]:
    assert edge.source is not None
    assert edge.target is not None
    source_box = boxes[edge.source]
    target_box = boxes[edge.target]
    source_side, target_side = _resolve_relative_sides(source_box, target_box, edge.source_side, edge.target_side)
    start = _anchor_point(source_box, source_side)
    end = _anchor_point(target_box, target_side)
    return _route_with_obstacles(
        start,
        end,
        source_side,
        target_side,
        _expanded_obstacles(boxes, margin=16),
        start_exclude=edge.source,
        end_exclude=edge.target,
    )


def _idef1x_edge_points(edge: Edge, boxes: dict[str, Box]) -> list[tuple[float, float]]:
    assert edge.source is not None
    assert edge.target is not None
    source_box = boxes[edge.source]
    target_box = boxes[edge.target]
    source_side, target_side = _resolve_relative_sides(source_box, target_box, edge.source_side, edge.target_side)
    start = _anchor_point(source_box, source_side)
    end = _anchor_point(target_box, target_side)
    return _route_with_obstacles(
        start,
        end,
        source_side,
        target_side,
        _expanded_obstacles(boxes, margin=16),
        start_exclude=edge.source,
        end_exclude=edge.target,
    )


def _uml_markers(edge: Edge) -> tuple[str, str]:
    if edge.kind == "inheritance":
        return "none", "triangle_hollow"
    if edge.kind == "realization":
        return "none", "triangle_hollow"
    if edge.kind == "aggregation":
        return "diamond_hollow", "none"
    if edge.kind == "composition":
        return "diamond_filled", "none"
    if edge.kind == "dependency":
        return "none", "arrow"
    return "none", "none"


def _build_idef_anchor_map(
    diagram: Diagram,
    boxes: dict[str, Box],
) -> tuple[dict[str, tuple[str | None, str | None]], dict[tuple[str, str], tuple[float, float]]]:
    edge_sides: dict[str, tuple[str | None, str | None]] = {}
    requests: dict[tuple[str, str], list[tuple[str, str]]] = {}

    for edge in diagram.edges:
        source_side, target_side = _resolve_idef_sides(edge, boxes)
        edge_sides[edge.id] = (source_side, target_side)
        if edge.source is not None and source_side is not None:
            requests.setdefault((edge.source, source_side), []).append((edge.id, "source"))
        if edge.target is not None and target_side is not None:
            requests.setdefault((edge.target, target_side), []).append((edge.id, "target"))

    anchors: dict[tuple[str, str], tuple[float, float]] = {}
    for (node_id, side), items in requests.items():
        box = boxes[node_id]
        for index, (edge_id, endpoint_kind) in enumerate(items, start=1):
            anchors[(edge_id, endpoint_kind)] = _slot_anchor(box, side, index, len(items))

    return edge_sides, anchors


def _build_idef_boundary_groups(
    diagram: Diagram,
    edge_sides: dict[str, tuple[str | None, str | None]],
) -> tuple[list[IdefBoundaryGroup], set[str]]:
    if diagram.type != "idef0" or len(diagram.nodes) <= 1:
        return [], set()

    buckets: dict[tuple[str, str, str | None, str | None, str | None, tuple[object, ...]], list[Edge]] = {}
    for edge in diagram.edges:
        direction: str | None = None
        side: str | None = None
        if edge.source is None and edge.target is not None and not edge.tunnel_source:
            direction = "in"
            side = edge_sides[edge.id][1]
        elif edge.target is None and edge.source is not None and not edge.tunnel_target:
            direction = "out"
            side = edge_sides[edge.id][0]

        if direction is None or side is None or (edge.label is None and edge.icom is None):
            continue

        key = (
            direction,
            side,
            edge.role,
            edge.label,
            edge.icom,
            _style_fingerprint(edge.style),
        )
        buckets.setdefault(key, []).append(edge)

    groups: list[IdefBoundaryGroup] = []
    grouped_edge_ids: set[str] = set()
    for (direction, side, _, label, icom, _), edges in buckets.items():
        if len(edges) < 2:
            continue
        groups.append(
            IdefBoundaryGroup(
                direction=direction,
                side=side,
                label=label,
                icom=icom,
                edges=tuple(edges),
            )
        )
        grouped_edge_ids.update(edge.id for edge in edges)

    return groups, grouped_edge_ids


def _style_fingerprint(style: Style | None) -> tuple[object, ...]:
    if style is None:
        return ()
    return (
        style.fill,
        style.stroke,
        style.stroke_width,
        style.text_color,
        style.font_size,
        style.font_weight,
        style.dasharray,
        style.label_fill,
        style.label_stroke,
        style.corner_radius,
        style.opacity,
        style.arrow_scale,
    )


def _render_idef_boundary_group(
    diagram: Diagram,
    group: IdefBoundaryGroup,
    *,
    boxes: dict[str, Box],
    content_bounds: Bounds,
    edge_sides: dict[str, tuple[str | None, str | None]],
    anchor_map: dict[tuple[str, str], tuple[float, float]],
) -> list[str]:
    representative = group.edges[0]
    style = _edge_style(diagram, representative, stroke="#111827", stroke_width=1.7, text_color="#111827")
    obstacles = _expanded_obstacles(boxes, margin=12)
    anchors: list[tuple[Edge, tuple[float, float], str]] = []
    for edge in group.edges:
        source_side, target_side = edge_sides[edge.id]
        if group.direction == "in":
            anchors.append((edge, anchor_map[(edge.id, "target")], target_side or group.side))
        else:
            anchors.append((edge, anchor_map[(edge.id, "source")], source_side or group.side))

    branch_axis = [anchor[1] for _, anchor, _ in anchors] if group.side in {"left", "right"} else [anchor[0] for _, anchor, _ in anchors]
    boundary_axis = _median_value(branch_axis)
    lead_path, backbone_path = _idef_group_trunk_paths(group.side, boundary_axis, anchors, boxes, content_bounds)

    parts: list[str] = []
    if backbone_path is not None:
        parts.extend(_render_edge(backbone_path, None, style))
    lead_marker = "none" if group.direction == "in" else "arrow"
    parts.extend(_render_edge(lead_path, representative.label, style, end_marker=lead_marker))
    parts.extend(_render_idef_boundary_annotations(representative, lead_path, style))

    for edge, anchor, side in anchors:
        branch_style = _edge_style(diagram, edge, stroke="#111827", stroke_width=1.7, text_color="#111827")
        branch_start = _group_branch_start(group.side, anchor, backbone_path, lead_path)
        branch_path = _idef_group_branch_path(
            group.direction,
            group.side,
            branch_start,
            anchor,
            side,
            edge=edge,
            obstacles=obstacles,
            content_bounds=content_bounds,
        )
        branch_marker = "arrow" if group.direction == "in" else "none"
        parts.extend(
            _render_edge(
                branch_path,
                None,
                branch_style,
                end_marker=branch_marker,
                source_label=edge.source_label,
                target_label=edge.target_label,
            )
        )

    return parts


def _idef_group_trunk_paths(
    side: str,
    boundary_axis: float,
    anchors: list[tuple[Edge, tuple[float, float], str]],
    boxes: dict[str, Box],
    content_bounds: Bounds,
) -> tuple[list[tuple[float, float]], list[tuple[float, float]] | None]:
    if side == "left":
        rail = max(content_bounds.left + ROUTE_MARGIN, min(box.x for box in boxes.values()) - ROUTE_MARGIN * 2)
        min_y = min(anchor[1] for _, anchor, _ in anchors)
        max_y = max(anchor[1] for _, anchor, _ in anchors)
        lead = [(content_bounds.left, boundary_axis), (rail, boundary_axis)]
        backbone = None if _almost_equal(min_y, max_y) else [(rail, min_y), (rail, max_y)]
        return lead, backbone
    if side == "right":
        rail = min(content_bounds.right - ROUTE_MARGIN, max(box.x + box.width for box in boxes.values()) + ROUTE_MARGIN * 2)
        min_y = min(anchor[1] for _, anchor, _ in anchors)
        max_y = max(anchor[1] for _, anchor, _ in anchors)
        lead = [(rail, boundary_axis), (content_bounds.right, boundary_axis)]
        backbone = None if _almost_equal(min_y, max_y) else [(rail, min_y), (rail, max_y)]
        return lead, backbone
    if side == "top":
        rail = max(content_bounds.top + ROUTE_MARGIN, min(box.y for box in boxes.values()) - ROUTE_MARGIN * 2)
        min_x = min(anchor[0] for _, anchor, _ in anchors)
        max_x = max(anchor[0] for _, anchor, _ in anchors)
        lead = [(boundary_axis, content_bounds.top), (boundary_axis, rail)]
        backbone = None if _almost_equal(min_x, max_x) else [(min_x, rail), (max_x, rail)]
        return lead, backbone

    rail = min(content_bounds.bottom - ROUTE_MARGIN, max(box.y + box.height for box in boxes.values()) + ROUTE_MARGIN * 2)
    min_x = min(anchor[0] for _, anchor, _ in anchors)
    max_x = max(anchor[0] for _, anchor, _ in anchors)
    lead = [(boundary_axis, rail), (boundary_axis, content_bounds.bottom)]
    backbone = None if _almost_equal(min_x, max_x) else [(min_x, rail), (max_x, rail)]
    return lead, backbone


def _group_branch_start(
    side: str,
    anchor: tuple[float, float],
    backbone_path: list[tuple[float, float]] | None,
    lead_path: list[tuple[float, float]],
) -> tuple[float, float]:
    if side in {"left", "right"}:
        x = backbone_path[0][0] if backbone_path is not None else lead_path[-1][0]
        return x, anchor[1]
    y = backbone_path[0][1] if backbone_path is not None else lead_path[-1][1]
    return anchor[0], y


def _idef_group_branch_path(
    direction: str,
    side: str,
    branch_start: tuple[float, float],
    anchor: tuple[float, float],
    anchor_side: str,
    *,
    edge: Edge,
    obstacles: dict[str, Bounds],
    content_bounds: Bounds,
) -> list[tuple[float, float]]:
    if direction == "in":
        source_side = "right" if side == "left" else "left" if side == "right" else "bottom" if side == "top" else "top"
        return _route_with_obstacles(
            branch_start,
            anchor,
            source_side,
            anchor_side,
            obstacles,
            end_exclude=edge.target,
            content_bounds=content_bounds,
            preferences=_group_branch_preferences(side, branch_start),
        )

    target_side = "left" if side == "right" else "right" if side == "left" else "top" if side == "bottom" else "bottom"
    return _route_with_obstacles(
        anchor,
        branch_start,
        anchor_side,
        target_side,
        obstacles,
        start_exclude=edge.source,
        content_bounds=content_bounds,
        preferences=_group_branch_preferences(side, branch_start),
    )


def _group_branch_preferences(side: str, branch_start: tuple[float, float]) -> RoutePreferences:
    if side in {"left", "right"}:
        return RoutePreferences(preferred_ys=(branch_start[1],))
    return RoutePreferences(preferred_xs=(branch_start[0],))


def _median_value(values: list[float]) -> float:
    ordered = sorted(values)
    middle = len(ordered) // 2
    if len(ordered) % 2 == 1:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) / 2


def _resolve_idef_sides(edge: Edge, boxes: dict[str, Box]) -> tuple[str | None, str | None]:
    if edge.source is None and edge.target is not None:
        return None, edge.target_side or _role_to_side(edge.role, target=True) or "left"
    if edge.target is None and edge.source is not None:
        return edge.source_side or _role_to_side(edge.role, target=False) or "right", None

    assert edge.source is not None
    assert edge.target is not None
    source_box = boxes[edge.source]
    target_box = boxes[edge.target]
    source_side = edge.source_side or _role_to_side(edge.role, target=False)
    target_side = edge.target_side or _role_to_side(edge.role, target=True)
    return _resolve_relative_sides(source_box, target_box, source_side, target_side)


def _idef0_edge_points(
    edge: Edge,
    boxes: dict[str, Box],
    content_bounds: Bounds,
    edge_sides: dict[str, tuple[str | None, str | None]],
    anchor_map: dict[tuple[str, str], tuple[float, float]],
) -> list[tuple[float, float]]:
    source_side, target_side = edge_sides[edge.id]
    preferences = _idef_route_preferences(edge, source_side, target_side, content_bounds)

    if edge.source is None and edge.target is not None:
        end = anchor_map[(edge.id, "target")]
        start = _frame_boundary_point(content_bounds, target_side or "left", end)
        return _route_with_obstacles(
            start,
            end,
            target_side or "left",
            target_side or "left",
            _expanded_obstacles(boxes, margin=12),
            end_exclude=edge.target,
            content_bounds=content_bounds,
            preferences=preferences,
        )

    if edge.target is None and edge.source is not None:
        start = anchor_map[(edge.id, "source")]
        end = _frame_boundary_point(content_bounds, source_side or "right", start)
        return _route_with_obstacles(
            start,
            end,
            source_side or "right",
            source_side or "right",
            _expanded_obstacles(boxes, margin=12),
            start_exclude=edge.source,
            content_bounds=content_bounds,
            preferences=preferences,
        )

    assert edge.source is not None
    assert edge.target is not None
    start = anchor_map[(edge.id, "source")]
    end = anchor_map[(edge.id, "target")]
    assert source_side is not None
    assert target_side is not None
    return _route_with_obstacles(
        start,
        end,
        source_side,
        target_side,
        _expanded_obstacles(boxes, margin=12),
        start_exclude=edge.source,
        end_exclude=edge.target,
        content_bounds=content_bounds,
        preferences=preferences,
    )


def _idef_route_preferences(
    edge: Edge,
    source_side: str | None,
    target_side: str | None,
    content_bounds: Bounds,
) -> RoutePreferences:
    inset = 18.0
    preferred_xs: tuple[float, ...] = ()
    preferred_ys: tuple[float, ...] = ()

    if edge.role == "control" or (source_side == "top" and target_side == "top"):
        preferred_ys = (content_bounds.top + inset,)
    elif edge.role == "mechanism" or (source_side == "bottom" and target_side == "bottom"):
        preferred_ys = (content_bounds.bottom - inset,)
    elif edge.role == "input" or (source_side == "left" and target_side == "left"):
        preferred_xs = (content_bounds.left + inset,)
    elif edge.role == "output" or (source_side == "right" and target_side == "right"):
        preferred_xs = (content_bounds.right - inset,)

    return RoutePreferences(preferred_xs=preferred_xs, preferred_ys=preferred_ys)


def _route_with_obstacles(
    start: tuple[float, float],
    end: tuple[float, float],
    source_side: str,
    target_side: str,
    obstacles: dict[str, Bounds],
    start_exclude: str | None = None,
    end_exclude: str | None = None,
    content_bounds: Bounds | None = None,
    preferences: RoutePreferences | None = None,
) -> list[tuple[float, float]]:
    preferences = preferences or RoutePreferences()
    if (
        (_almost_equal(start[0], end[0]) or _almost_equal(start[1], end[1]))
        and _path_is_clear([start, end], obstacles, start_exclude, end_exclude)
    ):
        return _finalize_route([start, end], source_side, target_side, obstacles, start_exclude, end_exclude)

    exit_distance = ROUTE_MARGIN
    start_exit = _shift_point(start, source_side, exit_distance)
    end_entry = _shift_point(end, target_side, exit_distance)

    if _path_is_clear([start, start_exit], obstacles, start_exclude, end_exclude) and _path_is_clear(
        [end_entry, end], obstacles, start_exclude, end_exclude
    ):
        core_path = _graph_route(
            start_exit,
            end_entry,
            obstacles=obstacles,
            start_exclude=start_exclude,
            end_exclude=end_exclude,
            content_bounds=content_bounds,
            preferences=preferences,
        )
        if core_path is not None:
            return _finalize_route([start, *core_path, end], source_side, target_side, obstacles, start_exclude, end_exclude)

    return _finalize_route(
        _orthogonal_route(start, end, source_side, target_side, loop_offset=52),
        source_side,
        target_side,
        obstacles,
        start_exclude,
        end_exclude,
    )


def _graph_route(
    start: tuple[float, float],
    end: tuple[float, float],
    obstacles: dict[str, Bounds],
    start_exclude: str | None,
    end_exclude: str | None,
    content_bounds: Bounds | None,
    preferences: RoutePreferences,
) -> list[tuple[float, float]] | None:
    axes_x, axes_y = _routing_axes(start, end, obstacles, content_bounds, preferences)
    graph = _build_routing_graph(
        axes_x,
        axes_y,
        obstacles=obstacles,
        start_exclude=start_exclude,
        end_exclude=end_exclude,
        content_bounds=content_bounds,
    )
    if start not in graph or end not in graph:
        return None

    queue: list[tuple[float, int, tuple[float, float], str | None]] = []
    distances: dict[tuple[tuple[float, float], str | None], float] = {}
    previous: dict[tuple[tuple[float, float], str | None], tuple[tuple[float, float], str | None] | None] = {}
    counter = 0
    start_state = (start, None)
    distances[start_state] = 0.0
    previous[start_state] = None
    heappush(queue, (0.0, counter, start, None))

    best_end_state: tuple[tuple[float, float], str | None] | None = None
    while queue:
        cost, _, point, direction = heappop(queue)
        state = (point, direction)
        if cost > distances.get(state, float("inf")):
            continue
        if point == end:
            best_end_state = state
            break

        for neighbor, next_direction, segment_length in graph[point]:
            bend_cost = preferences.bend_penalty if direction and direction != next_direction else 0.0
            axis_cost = _axis_penalty(point, neighbor, next_direction, preferences)
            next_state = (neighbor, next_direction)
            next_cost = cost + segment_length + bend_cost + axis_cost
            if next_cost >= distances.get(next_state, float("inf")):
                continue
            distances[next_state] = next_cost
            previous[next_state] = state
            counter += 1
            heappush(queue, (next_cost, counter, neighbor, next_direction))

    if best_end_state is None:
        return None

    path: list[tuple[float, float]] = []
    cursor: tuple[tuple[float, float], str | None] | None = best_end_state
    while cursor is not None:
        path.append(cursor[0])
        cursor = previous.get(cursor)
    path.reverse()
    return _simplify_path(path)


def _routing_axes(
    start: tuple[float, float],
    end: tuple[float, float],
    obstacles: dict[str, Bounds],
    content_bounds: Bounds | None,
    preferences: RoutePreferences,
) -> tuple[list[float], list[float]]:
    lanes = _route_lanes(obstacles, content_bounds)
    xs = {start[0], end[0], (start[0] + end[0]) / 2, lanes["left"], lanes["right"], *preferences.preferred_xs}
    ys = {start[1], end[1], (start[1] + end[1]) / 2, lanes["top"], lanes["bottom"], *preferences.preferred_ys}

    if content_bounds is not None:
        xs.update({content_bounds.left, content_bounds.right})
        ys.update({content_bounds.top, content_bounds.bottom})

    for obstacle in obstacles.values():
        xs.update({obstacle.left, obstacle.right})
        ys.update({obstacle.top, obstacle.bottom})

    return sorted(xs), sorted(ys)


def _build_routing_graph(
    axes_x: list[float],
    axes_y: list[float],
    obstacles: dict[str, Bounds],
    start_exclude: str | None,
    end_exclude: str | None,
    content_bounds: Bounds | None,
) -> dict[tuple[float, float], list[tuple[tuple[float, float], str, float]]]:
    points: dict[tuple[float, float], list[tuple[tuple[float, float], str, float]]] = {}
    by_y: dict[float, list[tuple[float, float]]] = {}
    by_x: dict[float, list[tuple[float, float]]] = {}

    for x in axes_x:
        for y in axes_y:
            point = (x, y)
            if not _point_is_available(point, obstacles, start_exclude, end_exclude, content_bounds):
                continue
            points[point] = []
            by_y.setdefault(y, []).append(point)
            by_x.setdefault(x, []).append(point)

    for y, row_points in by_y.items():
        ordered = sorted(row_points, key=lambda item: item[0])
        for first, second in zip(ordered, ordered[1:]):
            if _segment_is_clear(first, second, obstacles, start_exclude, end_exclude):
                distance = abs(second[0] - first[0])
                points[first].append((second, "h", distance))
                points[second].append((first, "h", distance))

    for x, column_points in by_x.items():
        ordered = sorted(column_points, key=lambda item: item[1])
        for first, second in zip(ordered, ordered[1:]):
            if _segment_is_clear(first, second, obstacles, start_exclude, end_exclude):
                distance = abs(second[1] - first[1])
                points[first].append((second, "v", distance))
                points[second].append((first, "v", distance))

    return points


def _point_is_available(
    point: tuple[float, float],
    obstacles: dict[str, Bounds],
    start_exclude: str | None,
    end_exclude: str | None,
    content_bounds: Bounds | None,
) -> bool:
    if content_bounds is not None and not content_bounds.contains_point(point):
        return False
    for obstacle_id, obstacle in obstacles.items():
        if obstacle_id in {start_exclude, end_exclude}:
            continue
        if obstacle.left < point[0] < obstacle.right and obstacle.top < point[1] < obstacle.bottom:
            return False
    return True


def _segment_is_clear(
    start: tuple[float, float],
    end: tuple[float, float],
    obstacles: dict[str, Bounds],
    start_exclude: str | None,
    end_exclude: str | None,
) -> bool:
    return _path_is_clear([start, end], obstacles, start_exclude, end_exclude)


def _axis_penalty(
    start: tuple[float, float],
    end: tuple[float, float],
    direction: str,
    preferences: RoutePreferences,
) -> float:
    if direction == "h" and preferences.preferred_ys:
        y = start[1]
        return min(abs(y - preferred_y) for preferred_y in preferences.preferred_ys) * PREFERRED_AXIS_WEIGHT
    if direction == "v" and preferences.preferred_xs:
        x = start[0]
        return min(abs(x - preferred_x) for preferred_x in preferences.preferred_xs) * PREFERRED_AXIS_WEIGHT
    return 0.0


def _orthogonal_route(
    start: tuple[float, float],
    end: tuple[float, float],
    source_side: str,
    target_side: str,
    loop_offset: float,
) -> list[tuple[float, float]]:
    if _almost_equal(start[0], end[0]) or _almost_equal(start[1], end[1]):
        return [start, end]

    if source_side == target_side:
        if source_side == "right":
            bend_x = max(start[0], end[0]) + loop_offset
            return [start, (bend_x, start[1]), (bend_x, end[1]), end]
        if source_side == "left":
            bend_x = min(start[0], end[0]) - loop_offset
            return [start, (bend_x, start[1]), (bend_x, end[1]), end]
        if source_side == "top":
            bend_y = min(start[1], end[1]) - loop_offset
            return [start, (start[0], bend_y), (end[0], bend_y), end]
        bend_y = max(start[1], end[1]) + loop_offset
        return [start, (start[0], bend_y), (end[0], bend_y), end]

    source_vertical = source_side in {"top", "bottom"}
    target_vertical = target_side in {"top", "bottom"}

    if source_vertical and target_vertical:
        middle_y = (start[1] + end[1]) / 2
        return [start, (start[0], middle_y), (end[0], middle_y), end]
    if not source_vertical and not target_vertical:
        middle_x = (start[0] + end[0]) / 2
        return [start, (middle_x, start[1]), (middle_x, end[1]), end]
    if source_vertical:
        return [start, (start[0], end[1]), end]
    return [start, (end[0], start[1]), end]


def _expanded_obstacles(boxes: dict[str, Box], margin: float) -> dict[str, Bounds]:
    return {
        node_id: Bounds(
            left=box.x - margin,
            top=box.y - margin,
            right=box.x + box.width + margin,
            bottom=box.y + box.height + margin,
        )
        for node_id, box in boxes.items()
    }


def _route_lanes(obstacles: dict[str, Bounds], content_bounds: Bounds | None) -> dict[str, float]:
    min_left = min(bound.left for bound in obstacles.values())
    max_right = max(bound.right for bound in obstacles.values())
    min_top = min(bound.top for bound in obstacles.values())
    max_bottom = max(bound.bottom for bound in obstacles.values())

    left = min_left - ROUTE_MARGIN * 2
    right = max_right + ROUTE_MARGIN * 2
    top = min_top - ROUTE_MARGIN * 2
    bottom = max_bottom + ROUTE_MARGIN * 2

    if content_bounds is not None:
        left = max(content_bounds.left, left)
        right = min(content_bounds.right, right)
        top = max(content_bounds.top, top)
        bottom = min(content_bounds.bottom, bottom)

    return {"left": left, "right": right, "top": top, "bottom": bottom}


def _shift_point(point: tuple[float, float], side: str, distance: float) -> tuple[float, float]:
    if side == "top":
        return (point[0], point[1] - distance)
    if side == "right":
        return (point[0] + distance, point[1])
    if side == "bottom":
        return (point[0], point[1] + distance)
    return (point[0] - distance, point[1])


def _path_is_clear(
    points: list[tuple[float, float]],
    obstacles: dict[str, Bounds],
    start_exclude: str | None,
    end_exclude: str | None,
) -> bool:
    for segment_start, segment_end in zip(points, points[1:]):
        for obstacle_id, obstacle in obstacles.items():
            if obstacle_id in {start_exclude, end_exclude}:
                continue
            if _segment_intersects_bounds(segment_start, segment_end, obstacle):
                return False
    return True


def _finalize_route(
    points: list[tuple[float, float]],
    source_side: str,
    target_side: str,
    obstacles: dict[str, Bounds],
    start_exclude: str | None,
    end_exclude: str | None,
) -> list[tuple[float, float]]:
    base_path = _simplify_path(points)
    normalized = _normalize_route_start(points, source_side)
    normalized = _normalize_route_end(normalized, target_side)
    normalized = _simplify_path(normalized, source_side=source_side, target_side=target_side)
    try:
        _validate_route_shape(normalized)
        _validate_route_terminals(normalized, source_side, target_side)
        _validate_no_redundant_points(normalized, source_side=source_side, target_side=target_side)
        if not _path_is_clear(normalized, obstacles, start_exclude, end_exclude):
            raise ValueError("Route intersects an obstacle after terminal normalization")
        return normalized
    except ValueError:
        _validate_route_shape(base_path)
        _validate_no_redundant_points(base_path)
        return base_path


def _normalize_route_start(
    points: list[tuple[float, float]],
    side: str,
    distance: float = ROUTE_MARGIN,
) -> list[tuple[float, float]]:
    if len(points) < 2 or _segment_matches_side(points[0], points[1], side):
        return points

    start = points[0]
    nxt = points[1]
    start_exit = _shift_point(start, side, distance)
    normalized = [start]
    _append_unique_point(normalized, start_exit)
    if side in {"top", "bottom"}:
        _append_unique_point(normalized, (nxt[0], start_exit[1]))
    else:
        _append_unique_point(normalized, (start_exit[0], nxt[1]))
    for point in points[1:]:
        _append_unique_point(normalized, point)
    return normalized


def _normalize_route_end(
    points: list[tuple[float, float]],
    side: str,
    distance: float = ROUTE_MARGIN,
) -> list[tuple[float, float]]:
    if len(points) < 2 or _segment_enters_side(points[-2], points[-1], side):
        return points

    end = points[-1]
    prev = points[-2]
    end_entry = _shift_point(end, side, distance)
    normalized = list(points[:-1])
    if side in {"top", "bottom"}:
        _append_unique_point(normalized, (prev[0], end_entry[1]))
        _append_unique_point(normalized, end_entry)
    else:
        _append_unique_point(normalized, (end_entry[0], prev[1]))
        _append_unique_point(normalized, end_entry)
    _append_unique_point(normalized, end)
    return normalized


def _append_unique_point(points: list[tuple[float, float]], point: tuple[float, float]) -> None:
    if not points:
        points.append(point)
        return
    last = points[-1]
    if _almost_equal(last[0], point[0]) and _almost_equal(last[1], point[1]):
        return
    points.append(point)


def _validate_route_shape(points: list[tuple[float, float]]) -> None:
    if len(points) < 2:
        raise ValueError("Route must contain at least two points")
    for start, end in zip(points, points[1:]):
        if _almost_equal(start[0], end[0]) and _almost_equal(start[1], end[1]):
            raise ValueError("Route contains a zero-length segment")
        if not (_almost_equal(start[0], end[0]) or _almost_equal(start[1], end[1])):
            raise ValueError("Route contains a diagonal segment")


def _validate_route_terminals(points: list[tuple[float, float]], source_side: str, target_side: str) -> None:
    if not _segment_matches_side(points[0], points[1], source_side):
        raise ValueError(f"Route leaves source on the wrong side: expected {source_side}")
    if not _segment_enters_side(points[-2], points[-1], target_side):
        raise ValueError(f"Route enters target on the wrong side: expected {target_side}")


def _validate_no_redundant_points(
    points: list[tuple[float, float]],
    *,
    source_side: str | None = None,
    target_side: str | None = None,
) -> None:
    for index, (first, middle, last) in enumerate(zip(points, points[1:], points[2:]), start=1):
        if (_almost_equal(first[0], middle[0]) and _almost_equal(middle[0], last[0])) or (
            _almost_equal(first[1], middle[1]) and _almost_equal(middle[1], last[1])
        ):
            if index == 1 and source_side and _segment_matches_side(first, middle, source_side):
                continue
            if index == len(points) - 2 and target_side and _segment_enters_side(middle, last, target_side):
                continue
            raise ValueError("Route contains a redundant collinear point")


def _segment_intersects_bounds(
    start: tuple[float, float],
    end: tuple[float, float],
    bounds: Bounds,
) -> bool:
    if _almost_equal(start[0], end[0]):
        x = start[0]
        if not (bounds.left < x < bounds.right):
            return False
        low = min(start[1], end[1])
        high = max(start[1], end[1])
        return low < bounds.bottom and high > bounds.top

    if _almost_equal(start[1], end[1]):
        y = start[1]
        if not (bounds.top < y < bounds.bottom):
            return False
        low = min(start[0], end[0])
        high = max(start[0], end[0])
        return low < bounds.right and high > bounds.left

    return True


def _simplify_path(
    points: list[tuple[float, float]],
    *,
    source_side: str | None = None,
    target_side: str | None = None,
) -> list[tuple[float, float]]:
    deduped: list[tuple[float, float]] = []
    for point in points:
        if not deduped or not (_almost_equal(point[0], deduped[-1][0]) and _almost_equal(point[1], deduped[-1][1])):
            deduped.append(point)

    simplified = list(deduped)
    index = 1
    while index < len(simplified) - 1:
        a, b, c = simplified[index - 1], simplified[index], simplified[index + 1]
        collinear = (_almost_equal(a[0], b[0]) and _almost_equal(b[0], c[0])) or (
            _almost_equal(a[1], b[1]) and _almost_equal(b[1], c[1])
        )
        if not collinear:
            index += 1
            continue
        if index == 1 and source_side and not _segment_matches_side(a, c, source_side):
            index += 1
            continue
        if index == len(simplified) - 2 and target_side and not _segment_enters_side(a, c, target_side):
            index += 1
            continue
        simplified.pop(index)
        if index > 1:
            index -= 1
    return simplified


def _path_length(points: list[tuple[float, float]]) -> float:
    return sum(abs(end[0] - start[0]) + abs(end[1] - start[1]) for start, end in zip(points, points[1:]))


def _segment_matches_side(start: tuple[float, float], end: tuple[float, float], side: str) -> bool:
    if side == "top":
        return _almost_equal(start[0], end[0]) and end[1] < start[1]
    if side == "bottom":
        return _almost_equal(start[0], end[0]) and end[1] > start[1]
    if side == "left":
        return _almost_equal(start[1], end[1]) and end[0] < start[0]
    if side == "right":
        return _almost_equal(start[1], end[1]) and end[0] > start[0]
    raise ValueError(f"Unsupported route side '{side}'")


def _segment_enters_side(start: tuple[float, float], end: tuple[float, float], side: str) -> bool:
    opposite = {"top": "bottom", "bottom": "top", "left": "right", "right": "left"}
    return _segment_matches_side(start, end, opposite[side])


def _role_to_side(role: str | None, target: bool) -> str | None:
    if role is None:
        return None
    if role == "input":
        return "left" if target else "right"
    if role == "control":
        return "top" if target else "bottom"
    if role == "output":
        return "left" if target else "right"
    return "bottom" if target else "top"


def _resolve_relative_sides(
    source_box: Box,
    target_box: Box,
    source_side: str | None,
    target_side: str | None,
) -> tuple[str, str]:
    if source_side and target_side:
        return source_side, target_side

    dx = target_box.center_x - source_box.center_x
    dy = target_box.center_y - source_box.center_y
    if abs(dy) >= abs(dx):
        resolved_source = "bottom" if dy >= 0 else "top"
        resolved_target = "top" if dy >= 0 else "bottom"
    else:
        resolved_source = "right" if dx >= 0 else "left"
        resolved_target = "left" if dx >= 0 else "right"
    return source_side or resolved_source, target_side or resolved_target


def _anchor_point(box: Box, side: str) -> tuple[float, float]:
    if side == "top":
        return (box.center_x, box.y)
    if side == "right":
        return (box.x + box.width, box.center_y)
    if side == "bottom":
        return (box.center_x, box.y + box.height)
    return (box.x, box.center_y)


def _slot_anchor(box: Box, side: str, index: int, total: int) -> tuple[float, float]:
    padding = 22.0
    if side in {"top", "bottom"}:
        usable = max(20.0, box.width - padding * 2)
        x = box.x + padding + usable * index / (total + 1)
        y = box.y if side == "top" else box.y + box.height
        return (x, y)

    usable = max(20.0, box.height - padding * 2)
    y = box.y + padding + usable * index / (total + 1)
    x = box.x if side == "left" else box.x + box.width
    return (x, y)


def _frame_boundary_point(bounds: Bounds, side: str, anchor: tuple[float, float]) -> tuple[float, float]:
    if side == "left":
        return (bounds.left, anchor[1])
    if side == "right":
        return (bounds.right, anchor[1])
    if side == "top":
        return (anchor[0], bounds.top)
    return (anchor[0], bounds.bottom)


def _resolve_idef_frame(diagram: Diagram, document_title: str) -> Idef0Frame:
    frame = diagram.frame or Idef0Frame()
    first_code = next((node.code for node in diagram.nodes if node.code), diagram.id.upper())
    return Idef0Frame(
        enabled=frame.enabled,
        used_at=frame.used_at,
        author=frame.author or "Автогенерация",
        project=frame.project or document_title,
        date=frame.date,
        revision=frame.revision or "1",
        status=frame.status or "Рабочий проект",
        reader=frame.reader,
        context=frame.context or ("ВЕРХ" if first_code.upper() == "A0" else first_code.upper()),
        notes=frame.notes,
        page=frame.page or "1",
        node_ref=frame.node_ref or first_code,
    )


def _render_idef_frame(frame: Idef0Frame, diagram: Diagram, page_bounds: Bounds) -> list[str]:
    if not frame.enabled:
        return []

    frame_color = diagram.style.frame_stroke if diagram.style and diagram.style.frame_stroke else FRAME_COLOR
    inner_left = IDEF_FRAME_MARGIN
    inner_right = page_bounds.right - IDEF_FRAME_MARGIN
    header_top = IDEF_FRAME_MARGIN
    header_bottom = header_top + IDEF_HEADER_HEIGHT
    footer_bottom = page_bounds.bottom - IDEF_FRAME_MARGIN
    footer_top = footer_bottom - IDEF_FOOTER_HEIGHT
    width = inner_right - inner_left

    parts = [
        f'<line x1="{inner_left}" y1="{header_bottom}" x2="{inner_right}" y2="{header_bottom}" stroke="{frame_color}" stroke-width="1" />',
        f'<line x1="{inner_left}" y1="{footer_top}" x2="{inner_right}" y2="{footer_top}" stroke="{frame_color}" stroke-width="1" />',
    ]

    x1 = inner_left + width * 0.11
    x2 = inner_left + width * 0.30
    x3 = inner_left + width * 0.40
    x4 = inner_left + width * 0.58
    x5 = inner_left + width * 0.70
    x6 = inner_left + width * 0.80
    x7 = inner_right
    main_header_height = 48.0
    split_half = main_header_height / 2

    parts.extend(_render_frame_cell("ИСПОЛЬЗУЕТСЯ В", frame.used_at, inner_left, header_top, x1 - inner_left, main_header_height, frame_color=frame_color))
    parts.extend(_render_frame_cell("АВТОР", frame.author, x1, header_top, x2 - x1, split_half, frame_color=frame_color))
    parts.extend(_render_frame_cell("ПРОЕКТ", frame.project, x1, header_top + split_half, x2 - x1, split_half, frame_color=frame_color))
    parts.extend(_render_frame_cell("ДАТА", frame.date, x2, header_top, x3 - x2, split_half, frame_color=frame_color))
    parts.extend(_render_frame_cell("РЕВИЗИЯ", frame.revision, x2, header_top + split_half, x3 - x2, split_half, frame_color=frame_color))
    parts.extend(_render_status_block(frame.status, x3, header_top, x4 - x3, header_bottom - header_top, frame_color=frame_color))
    parts.extend(_render_frame_cell("ЧИТАТЕЛЬ", frame.reader, x4, header_top, x5 - x4, header_bottom - header_top, frame_color=frame_color))
    parts.extend(_render_frame_cell("ДАТА", frame.date, x5, header_top, x6 - x5, header_bottom - header_top, frame_color=frame_color))
    parts.extend(
        _render_frame_cell(
            "КОНТЕКСТ",
            frame.context,
            x6,
            header_top,
            x7 - x6,
            header_bottom - header_top,
            value_align="middle",
            value_font_size=16,
            value_weight="700",
            frame_color=frame_color,
        )
    )
    parts.extend(
        _render_frame_cell(
            "ЗАМЕЧАНИЯ",
            frame.notes,
            inner_left,
            header_top + main_header_height,
            x3 - inner_left,
            header_bottom - header_top - main_header_height,
            frame_color=frame_color,
        )
    )

    footer_cols = [inner_left, inner_left + width * 0.18, inner_left + width * 0.82, inner_right]
    parts.extend(_render_frame_cell("ВЕТКА", frame.node_ref, footer_cols[0], footer_top, footer_cols[1] - footer_cols[0], IDEF_FOOTER_HEIGHT, frame_color=frame_color))
    parts.extend(
        _render_frame_cell(
            "НАЗВАНИЕ",
            diagram.title,
            footer_cols[1],
            footer_top,
            footer_cols[2] - footer_cols[1],
            IDEF_FOOTER_HEIGHT,
            value_align="middle",
            value_font_size=12,
            frame_color=frame_color,
        )
    )
    parts.extend(
        _render_frame_cell(
            "НОМЕР",
            frame.page,
            footer_cols[2],
            footer_top,
            footer_cols[3] - footer_cols[2],
            IDEF_FOOTER_HEIGHT,
            frame_color=frame_color,
        )
    )

    return parts


def _render_frame_cell(
    label: str,
    value: str | None,
    x: float,
    y: float,
    width: float,
    height: float,
    value_align: str = "start",
    value_font_size: float = 8.8,
    value_weight: str = "400",
    frame_color: str = FRAME_COLOR,
) -> list[str]:
    parts = [
        f'<rect x="{x:.1f}" y="{y:.1f}" width="{width:.1f}" height="{height:.1f}" '
        f'fill="none" stroke="{frame_color}" stroke-width="0.8" />'
    ]
    if label:
        parts.append(
            f'<text x="{x + 4:.1f}" y="{y + 9.5:.1f}" text-anchor="start" font-size="7.2" '
            f'font-family="{FONT_FAMILY}" font-weight="700" fill="#475569">{escape(label)}</text>'
        )

    if not value:
        return parts

    value_lines = _wrap_text(value, max_width=max(36.0, width - 10), font_size=value_font_size)
    if value_align == "middle":
        line_height = value_font_size * 1.12
        block_height = len(value_lines[:3]) * line_height
        first_value_y = y + height / 2 - block_height / 2 + value_font_size * 0.78
        text_x = x + width / 2
        anchor = "middle"
    else:
        first_value_y = y + min(height - 6, 20)
        text_x = x + 6
        anchor = "start"

    for index, line in enumerate(value_lines[:3]):
        parts.append(
            f'<text x="{text_x:.1f}" y="{first_value_y + index * value_font_size * 1.15:.1f}" text-anchor="{anchor}" '
            f'font-size="{value_font_size}" font-family="{FONT_FAMILY}" font-weight="{value_weight}" fill="{TEXT_COLOR}">{escape(line)}</text>'
        )
    return parts


def _render_status_block(status: str | None, x: float, y: float, width: float, height: float, frame_color: str = FRAME_COLOR) -> list[str]:
    status_options = [
        "Рабочий проект",
        "Черновик",
        "Рекомендовано",
        "Публикация",
    ]
    selected = status or ""
    row_height = height / len(status_options)
    parts: list[str] = [
        f'<rect x="{x:.1f}" y="{y:.1f}" width="{width:.1f}" height="{height:.1f}" '
        f'fill="none" stroke="{frame_color}" stroke-width="0.8" />',
        f'<text x="{x + 4:.1f}" y="{y + 9.5:.1f}" text-anchor="start" font-size="7.2" '
        f'font-family="{FONT_FAMILY}" font-weight="700" fill="#475569">СТАТУС</text>',
    ]
    for index, option in enumerate(status_options):
        row_top = y + index * row_height
        if index:
            parts.append(
                f'<line x1="{x:.1f}" y1="{row_top:.1f}" x2="{x + width:.1f}" y2="{row_top:.1f}" '
                f'stroke="{frame_color}" stroke-width="0.6" />'
            )
        marker = "■" if option == selected else "□"
        parts.append(
            f'<text x="{x + 6:.1f}" y="{row_top + row_height / 2 + 3:.1f}" text-anchor="start" '
            f'font-size="8.3" font-family="{FONT_FAMILY}" fill="{TEXT_COLOR}">{marker} {escape(option)}</text>'
        )
    return parts


def _render_centered_text(
    label: str,
    box: Box,
    font_size: int,
    max_width: float,
    color: str = TEXT_COLOR,
    font_weight: str = "400",
) -> list[str]:
    lines = _wrap_text(label, max_width=max_width, font_size=font_size)
    line_height = font_size * 1.22
    start_y = box.center_y - ((len(lines) - 1) * line_height) / 2 + font_size * 0.35
    parts: list[str] = []
    for index, line in enumerate(lines):
        y = start_y + index * line_height
        parts.append(
            f'<text x="{box.center_x:.1f}" y="{y:.1f}" text-anchor="middle" font-size="{font_size}" '
            f'font-family="{FONT_FAMILY}" font-weight="{font_weight}" fill="{color}">{escape(line)}</text>'
        )
    return parts


def _best_label_segment(points: list[tuple[float, float]]) -> tuple[tuple[float, float], tuple[float, float]]:
    best = (points[0], points[-1])
    best_length = -1.0
    for start, end in zip(points, points[1:]):
        length = abs(end[0] - start[0]) + abs(end[1] - start[1])
        if length > best_length:
            best_length = length
            best = (start, end)
    return best


def _polyline_path(points: list[tuple[float, float]]) -> str:
    segments = [f"M {points[0][0]:.1f} {points[0][1]:.1f}"]
    for point_x, point_y in points[1:]:
        segments.append(f"L {point_x:.1f} {point_y:.1f}")
    return " ".join(segments)


def _points_to_svg(points: list[tuple[float, float]]) -> str:
    return " ".join(f"{x:.1f},{y:.1f}" for x, y in points)


def _wrap_text(label: str, max_width: float, font_size: int) -> list[str]:
    if not label:
        return [""]
    line_length = max(6, int(max_width / max(4.0, font_size * 0.58)))
    wrapped = textwrap.wrap(label, width=line_length, break_long_words=False)
    return wrapped or [label]


def _approx_text_width(label: str, font_size: int) -> float:
    return max(6.0, len(label) * font_size * 0.58)


def _almost_equal(left: float, right: float) -> bool:
    return abs(left - right) < 0.5
