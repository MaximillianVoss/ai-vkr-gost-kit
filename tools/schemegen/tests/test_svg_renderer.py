from __future__ import annotations

from pathlib import Path
import unittest

from schemegen.models import load_document, load_document_file
from schemegen.svg_renderer import (
    FLOW_GRID_X,
    FLOW_GRID_Y,
    FLOW_ORIGIN_X,
    FLOW_ORIGIN_Y,
    IDEF_GRID_X,
    IDEF_GRID_Y,
    IDEF_ORIGIN_X,
    IDEF_ORIGIN_Y,
    Bounds,
    _build_idef_anchor_map,
    _flowchart_edge_points,
    _idef0_edge_points,
    _idef_page_bounds,
    _layout_nodes,
    _polyline_path,
    _recenter_single_idef_box,
    _segment_intersects_bounds,
    render_svg,
)


ROOT = Path(__file__).resolve().parent.parent


class SvgRendererTests(unittest.TestCase):
    def test_render_flowchart_svg_contains_labels(self) -> None:
        document = load_document_file(ROOT / "examples" / "flowchart_only.json")
        svg = render_svg(document)
        self.assertIn("<svg", svg)
        self.assertIn("Алгоритм проверки входных данных", svg)
        self.assertIn("<polygon", svg)

    def test_render_flowchart_connector_and_offpage_symbols(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        svg = render_svg(document, "flow_04_connectors")
        self.assertIn("<circle", svg)
        self.assertIn("Лист 2 / B", svg)
        self.assertIn(">A</text>", svg)

    def test_render_specific_idef0_diagram(self) -> None:
        document = load_document_file(ROOT / "examples" / "coursework_document.json")
        svg = render_svg(document, "context_a0")
        self.assertIn("IDEF0 контекстная диаграмма", svg)
        self.assertIn("A0", svg)
        self.assertIn("АВТОР", svg)
        self.assertIn("КОНТЕКСТ", svg)
        self.assertIn("ВЕТКА", svg)

    def test_render_idef0_boundary_icom_and_tunnel_annotations(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        svg = render_svg(document, "idef0_05_decomposition")
        self.assertEqual(svg.count(">Исходные задачи</text>"), 1)
        self.assertEqual(svg.count(">ГОСТ и методичка</text>"), 1)
        self.assertEqual(svg.count(">I1</text>"), 1)
        self.assertEqual(svg.count(">C1</text>"), 1)
        self.assertIn(">(</text>", svg)
        self.assertIn(">)</text>", svg)

    def test_render_idef0_grouped_output_uses_single_boundary_label(self) -> None:
        document = load_document(
            {
                "schema_version": "1.0",
                "title": "Grouped output",
                "diagrams": [
                    {
                        "id": "a1",
                        "title": "A1",
                        "type": "idef0",
                        "nodes": [
                            {"id": "a11", "label": "Подготовить", "kind": "function", "row": 0, "column": 0},
                            {"id": "a12", "label": "Проверить", "kind": "function", "row": 0, "column": 1},
                            {"id": "a13", "label": "Выдать", "kind": "function", "row": 0, "column": 2},
                        ],
                        "edges": [
                            {"id": "e1", "source": "a11", "target": None, "role": "output", "icom": "O1", "label": "Пакет документов"},
                            {"id": "e2", "source": "a12", "target": None, "role": "output", "icom": "O1", "label": "Пакет документов"},
                        ],
                    }
                ],
            }
        )
        svg = render_svg(document)
        self.assertEqual(svg.count(">Пакет документов</text>"), 1)
        self.assertEqual(svg.count(">O1</text>"), 1)

    def test_render_nested_idef0_decomposition(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        svg = render_svg(document, "idef0_05a_solve_tasks_detail")
        self.assertIn("05a. IDEF0 декомпозиция A2", svg)
        self.assertIn(">1</text>", svg)
        self.assertIn("Проверить корректность и", svg)
        self.assertIn("формат", svg)

    def test_render_idef0_context_contains_purpose_and_viewpoint(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        svg = render_svg(document, "idef0_04_context")
        self.assertIn("PURPOSE", svg)
        self.assertIn("Автоматизировать подготовку решений", svg)
        self.assertIn("VIEWPOINT", svg)
        self.assertIn("Студент, оформляющий работу по ГОСТ", svg)

    def test_render_idef0_line_jumps_for_crossing_paths(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        svg = render_svg(document, "idef0_05c_line_jumps")
        self.assertIn('class="line-jump"', svg)
        self.assertIn('class="line-jump-mask"', svg)

    def test_render_idef0_external_edge_without_role_uses_default_side(self) -> None:
        document = load_document(
            {
                "schema_version": "1.0",
                "title": "Fallback",
                "diagrams": [
                    {
                        "id": "idef",
                        "title": "Fallback",
                        "type": "idef0",
                        "nodes": [
                            {
                                "id": "a0",
                                "label": "Функция",
                                "kind": "function",
                                "row": 0,
                                "column": 0
                            }
                        ],
                        "edges": [
                            {
                                "id": "e1",
                                "source": None,
                                "target": "a0",
                                "label": "Вход"
                            }
                        ]
                    }
                ]
            }
        )
        svg = render_svg(document)
        self.assertIn("Вход", svg)

    def test_straight_vertical_flow_edges_have_no_extra_bends(self) -> None:
        document = load_document_file(ROOT / "examples" / "flowchart_only.json")
        diagram = document.diagrams[0]
        boxes = _layout_nodes(
            diagram,
            lane_origin_x=FLOW_ORIGIN_X,
            lane_origin_y=FLOW_ORIGIN_Y,
            grid_x=FLOW_GRID_X,
            grid_y=FLOW_GRID_Y,
        )
        points = _flowchart_edge_points(diagram, diagram.edges[0], boxes)
        self.assertEqual(len(points), 2)
        self.assertAlmostEqual(points[0][0], points[1][0])

    def test_flowchart_layout_uses_actual_shape_sizes(self) -> None:
        document = load_document(
            {
                "schema_version": "1.0",
                "title": "Wide flow",
                "diagrams": [
                    {
                        "id": "wide",
                        "title": "Wide",
                        "type": "flowchart",
                        "nodes": [
                            {"id": "a", "label": "Широкий блок", "kind": "process", "row": 0, "column": 0, "width": 340},
                            {"id": "b", "label": "Еще один широкий блок", "kind": "process", "row": 0, "column": 1, "width": 340},
                            {"id": "c", "label": "Следующий шаг", "kind": "process", "row": 1, "column": 0, "height": 140},
                        ],
                        "edges": [],
                    }
                ],
            }
        )
        diagram = document.diagram_map["wide"]
        boxes = _layout_nodes(
            diagram,
            lane_origin_x=FLOW_ORIGIN_X,
            lane_origin_y=FLOW_ORIGIN_Y,
            grid_x=FLOW_GRID_X,
            grid_y=FLOW_GRID_Y,
        )
        self.assertGreaterEqual(boxes["b"].x - (boxes["a"].x + boxes["a"].width), 100.0)
        self.assertGreaterEqual(boxes["c"].y - (boxes["a"].y + boxes["a"].height), 80.0)

    def test_flowchart_decision_prefers_bottom_for_yes_and_side_for_no(self) -> None:
        document = load_document(
            {
                "schema_version": "1.0",
                "title": "Decision",
                "diagrams": [
                    {
                        "id": "decision",
                        "title": "Decision",
                        "type": "flowchart",
                        "nodes": [
                            {"id": "check", "label": "Готово?", "kind": "decision", "row": 0, "column": 0},
                            {"id": "next", "label": "Продолжить", "kind": "process", "row": 1, "column": 0},
                            {"id": "fix", "label": "Исправить", "kind": "process", "row": 0, "column": 1},
                        ],
                        "edges": [
                            {"id": "yes", "source": "check", "target": "next", "label": "Да"},
                            {"id": "no", "source": "check", "target": "fix", "label": "Нет"},
                        ],
                    }
                ],
            }
        )
        diagram = document.diagram_map["decision"]
        boxes = _layout_nodes(
            diagram,
            lane_origin_x=FLOW_ORIGIN_X,
            lane_origin_y=FLOW_ORIGIN_Y,
            grid_x=FLOW_GRID_X,
            grid_y=FLOW_GRID_Y,
        )
        yes_edge = next(item for item in diagram.edges if item.id == "yes")
        no_edge = next(item for item in diagram.edges if item.id == "no")
        yes_points = _flowchart_edge_points(diagram, yes_edge, boxes)
        no_points = _flowchart_edge_points(diagram, no_edge, boxes)
        self.assertAlmostEqual(yes_points[0][0], boxes["check"].center_x, delta=0.1)
        self.assertGreater(yes_points[-1][1], yes_points[0][1])
        self.assertGreater(no_points[0][0], boxes["check"].center_x)

    def test_flowchart_auto_loopback_uses_outer_left_rail(self) -> None:
        document = load_document(
            {
                "schema_version": "1.0",
                "title": "Loop",
                "diagrams": [
                    {
                        "id": "loop",
                        "title": "Loop",
                        "type": "flowchart",
                        "nodes": [
                            {"id": "start", "label": "Начало", "kind": "start", "row": 0, "column": 0},
                            {"id": "check", "label": "Есть элементы?", "kind": "decision", "row": 1, "column": 0},
                            {"id": "process", "label": "Обработать", "kind": "process", "row": 2, "column": 0},
                            {"id": "finish", "label": "Конец", "kind": "end", "row": 3, "column": 0},
                        ],
                        "edges": [
                            {"id": "e1", "source": "start", "target": "check"},
                            {"id": "e2", "source": "check", "target": "process", "label": "Да"},
                            {"id": "e3", "source": "process", "target": "check"},
                            {"id": "e4", "source": "check", "target": "finish", "label": "Нет"},
                        ],
                    }
                ],
            }
        )
        diagram = document.diagram_map["loop"]
        boxes = _layout_nodes(
            diagram,
            lane_origin_x=FLOW_ORIGIN_X,
            lane_origin_y=FLOW_ORIGIN_Y,
            grid_x=FLOW_GRID_X,
            grid_y=FLOW_GRID_Y,
        )
        edge = next(item for item in diagram.edges if item.id == "e3")
        points = _flowchart_edge_points(diagram, edge, boxes)
        left_bound = min(boxes["process"].x, boxes["check"].x)
        self.assertLess(min(point[0] for point in points[1:-1]), left_bound)
        for segment_start, segment_end in zip(points, points[1:]):
            for node_id, box in boxes.items():
                if node_id in {edge.source, edge.target}:
                    continue
                bounds = Bounds(box.x, box.y, box.x + box.width, box.y + box.height)
                self.assertFalse(_segment_intersects_bounds(segment_start, segment_end, bounds), node_id)

    def test_render_idef3_diagram_contains_junction(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        svg = render_svg(document, "idef3_06_behavior")
        self.assertIn("IDEF3 поведенческая диаграмма", svg)
        self.assertIn("J1", svg)
        self.assertIn("Начало анализа", svg)

    def test_render_uml_class_diagram_contains_members_and_style(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        svg = render_svg(document, "uml_07_classes")
        self.assertIn("OrderService", svg)
        self.assertIn("createOrder(dto: OrderDto): Order", svg)
        self.assertIn("IOrderRepository", svg)
        self.assertIn("#fef3c7", svg)
        self.assertIn("использует", svg)

    def test_render_idef1x_diagram_contains_columns_and_cardinality(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        svg = render_svg(document, "idef1x_08_data_model")
        self.assertIn("Customer", svg)
        self.assertIn("PK customer_id: uuid NOT NULL", svg)
        self.assertIn("0..N", svg)
        self.assertIn("stroke-dasharray=\"8 4\"", svg)

    def test_flowchart_loopback_route_avoids_other_boxes(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        diagram = document.diagram_map["flow_03_loop_subprocess"]
        edge = next(item for item in diagram.edges if item.id == "e10")
        boxes = _layout_nodes(
            diagram,
            lane_origin_x=FLOW_ORIGIN_X,
            lane_origin_y=FLOW_ORIGIN_Y,
            grid_x=FLOW_GRID_X,
            grid_y=FLOW_GRID_Y,
        )
        points = _flowchart_edge_points(diagram, edge, boxes)
        for segment_start, segment_end in zip(points, points[1:]):
            for node_id, box in boxes.items():
                if node_id in {edge.source, edge.target}:
                    continue
                bounds = Bounds(box.x, box.y, box.x + box.width, box.y + box.height)
                self.assertFalse(_segment_intersects_bounds(segment_start, segment_end, bounds), node_id)

    def test_flowchart_multi_cycle_example_keeps_return_paths_clear(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        diagram = document.diagram_map["flow_05_multi_cycle_review"]
        boxes = _layout_nodes(
            diagram,
            lane_origin_x=FLOW_ORIGIN_X,
            lane_origin_y=FLOW_ORIGIN_Y,
            grid_x=FLOW_GRID_X,
            grid_y=FLOW_GRID_Y,
        )
        for edge_id in {"e4", "e9", "e12"}:
            edge = next(item for item in diagram.edges if item.id == edge_id)
            points = _flowchart_edge_points(diagram, edge, boxes)
            for segment_start, segment_end in zip(points, points[1:]):
                for node_id, box in boxes.items():
                    if node_id in {edge.source, edge.target}:
                        continue
                    bounds = Bounds(box.x, box.y, box.x + box.width, box.y + box.height)
                    self.assertFalse(_segment_intersects_bounds(segment_start, segment_end, bounds), f"{edge_id}:{node_id}")

    def test_idef0_feedback_route_uses_top_corridor(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        diagram = document.diagram_map["idef0_05_decomposition"]
        edge = next(item for item in diagram.edges if item.id == "review_cycle")
        boxes = _layout_nodes(
            diagram,
            lane_origin_x=IDEF_ORIGIN_X,
            lane_origin_y=IDEF_ORIGIN_Y,
            grid_x=IDEF_GRID_X,
            grid_y=IDEF_GRID_Y,
        )
        page_bounds, content_bounds = _idef_page_bounds(diagram, boxes)
        edge_sides, anchor_map = _build_idef_anchor_map(diagram, boxes)
        points = _idef0_edge_points(edge, boxes, content_bounds, edge_sides, anchor_map)
        source_box = boxes[edge.source]
        target_box = boxes[edge.target]
        corridor_y = min(point[1] for point in points[1:-1])
        self.assertLess(corridor_y, min(source_box.y, target_box.y))

    def test_single_idef0_box_is_centered_in_content_area(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        diagram = document.diagram_map["idef0_04_context"]
        boxes = _layout_nodes(
            diagram,
            lane_origin_x=IDEF_ORIGIN_X,
            lane_origin_y=IDEF_ORIGIN_Y,
            grid_x=IDEF_GRID_X,
            grid_y=IDEF_GRID_Y,
        )
        _, content_bounds = _idef_page_bounds(diagram, boxes)
        _recenter_single_idef_box(diagram, boxes, content_bounds)
        _, centered_bounds = _idef_page_bounds(diagram, boxes)
        box = boxes["manage_tasks"]
        self.assertAlmostEqual(box.center_x, centered_bounds.left + centered_bounds.width / 2, delta=1.0)
        self.assertAlmostEqual(box.center_y, centered_bounds.top + centered_bounds.height / 2, delta=1.0)

    def test_idef0_child_boxes_have_horizontal_spacing(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        diagram = document.diagram_map["idef0_05_decomposition"]
        boxes = _layout_nodes(
            diagram,
            lane_origin_x=IDEF_ORIGIN_X,
            lane_origin_y=IDEF_ORIGIN_Y,
            grid_x=IDEF_GRID_X,
            grid_y=IDEF_GRID_Y,
        )
        ordered = sorted((boxes[node.id] for node in diagram.nodes), key=lambda box: box.x)
        for previous, current in zip(ordered, ordered[1:]):
            self.assertGreaterEqual(current.x - (previous.x + previous.width), 70.0)

    def test_idef0_child_boxes_follow_diagonal_layout(self) -> None:
        document = load_document_file(ROOT / "examples" / "test_progression.json")
        diagram = document.diagram_map["idef0_05_decomposition"]
        boxes = _layout_nodes(
            diagram,
            lane_origin_x=IDEF_ORIGIN_X,
            lane_origin_y=IDEF_ORIGIN_Y,
            grid_x=IDEF_GRID_X,
            grid_y=IDEF_GRID_Y,
        )
        ordered = sorted((boxes[node.id] for node in diagram.nodes), key=lambda box: box.x)
        centers_y = [box.center_y for box in ordered]
        self.assertEqual(centers_y, sorted(centers_y))
        self.assertGreater(centers_y[-1] - centers_y[0], 120.0)

    def test_idef_edge_path_uses_rounded_bends(self) -> None:
        path = _polyline_path([(0.0, 0.0), (40.0, 0.0), (40.0, 60.0)])
        self.assertIn("Q 40.0 0.0", path)
