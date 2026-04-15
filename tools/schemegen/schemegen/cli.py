from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from .models import ValidationError, load_document, load_document_file
from .schema import format_schema
from .svg_renderer import render_diagram_svg


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="schemegen",
        description="Generate coursework-friendly diagrams from JSON documents",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    validate_parser = subparsers.add_parser("validate", help="Validate a JSON document")
    validate_parser.add_argument("input", help="Path to the input JSON document")
    validate_parser.set_defaults(func=cmd_validate)

    render_parser = subparsers.add_parser("render", help="Render one diagram into SVG")
    render_parser.add_argument("input", nargs="?", help="Path to the input JSON document")
    render_parser.add_argument("-o", "--output", required=True, help="Output SVG file path")
    render_parser.add_argument("--diagram", help="Diagram id to render")
    render_parser.add_argument(
        "--stdin",
        action="store_true",
        help="Read the JSON document from standard input instead of a file",
    )
    render_parser.set_defaults(func=cmd_render)

    render_all_parser = subparsers.add_parser("render-all", help="Render all diagrams into SVG files")
    render_all_parser.add_argument("input", help="Path to the input JSON document")
    render_all_parser.add_argument("-o", "--output-dir", required=True, help="Output directory")
    render_all_parser.set_defaults(func=cmd_render_all)

    schema_parser = subparsers.add_parser("schema", help="Print the JSON schema")
    schema_parser.set_defaults(func=cmd_schema)

    examples_parser = subparsers.add_parser("examples", help="List bundled example files")
    examples_parser.set_defaults(func=cmd_examples)

    return parser


def cmd_validate(args: argparse.Namespace) -> int:
    document = load_document_file(args.input)
    print(
        f"OK: document '{document.title}' contains {len(document.diagrams)} diagram(s).",
        file=sys.stdout,
    )
    return 0


def _load_from_stdin() -> dict[str, object]:
    text = sys.stdin.read()
    if not text.strip():
        raise ValidationError("stdin is empty")
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValidationError("Document root must be an object")
    return data


def cmd_render(args: argparse.Namespace) -> int:
    if args.stdin:
        document = load_document(_load_from_stdin())
    elif args.input:
        document = load_document_file(args.input)
    else:
        raise ValidationError("render requires either an input file or --stdin")

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if args.diagram:
        diagram = next((item for item in document.diagrams if item.id == args.diagram), None)
        if diagram is None:
            raise ValidationError(f"Diagram '{args.diagram}' not found")
    else:
        diagram = document.diagrams[0]

    svg = render_diagram_svg(diagram, document.title)
    output_path.write_text(svg, encoding="utf-8")
    print(f"Rendered '{diagram.id}' to '{output_path}'.", file=sys.stdout)
    return 0


def cmd_render_all(args: argparse.Namespace) -> int:
    document = load_document_file(args.input)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    for diagram in document.diagrams:
        target = output_dir / f"{diagram.id}.svg"
        target.write_text(render_diagram_svg(diagram, document.title), encoding="utf-8")
        print(f"Rendered '{diagram.id}' to '{target}'.", file=sys.stdout)
    return 0


def cmd_schema(args: argparse.Namespace) -> int:
    del args
    print(format_schema(), file=sys.stdout)
    return 0


def cmd_examples(args: argparse.Namespace) -> int:
    del args
    examples_dir = Path(__file__).resolve().parent.parent / "examples"
    examples = sorted(examples_dir.glob("*.json"))
    if not examples:
        print("No bundled examples found.", file=sys.stdout)
        return 0

    print("Bundled examples:", file=sys.stdout)
    for item in examples:
        print(f"- {item}", file=sys.stdout)
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except (ValidationError, json.JSONDecodeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
