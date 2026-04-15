from __future__ import annotations

from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parent.parent


class CliTests(unittest.TestCase):
    def test_validate_command(self) -> None:
        result = subprocess.run(
            [
                sys.executable,
                "-m",
                "schemegen",
                "validate",
                str(ROOT / "examples" / "flowchart_only.json"),
            ],
            capture_output=True,
            text=True,
            cwd=ROOT,
            check=False,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("OK:", result.stdout)

    def test_render_all_command_writes_svg_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            result = subprocess.run(
                [
                    sys.executable,
                    "-m",
                    "schemegen",
                    "render-all",
                    str(ROOT / "examples" / "coursework_document.json"),
                    "-o",
                    temp_dir,
                ],
                capture_output=True,
                text=True,
                cwd=ROOT,
                check=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            expected = Path(temp_dir) / "context_a0.svg"
            self.assertTrue(expected.exists())

    def test_render_from_stdin(self) -> None:
        input_text = (ROOT / "examples" / "flowchart_only.json").read_text(encoding="utf-8")
        with tempfile.TemporaryDirectory() as temp_dir:
            output_file = Path(temp_dir) / "diagram.svg"
            result = subprocess.run(
                [
                    sys.executable,
                    "-m",
                    "schemegen",
                    "render",
                    "--stdin",
                    "-o",
                    str(output_file),
                ],
                input=input_text,
                capture_output=True,
                text=True,
                cwd=ROOT,
                check=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertTrue(output_file.exists())
            svg = output_file.read_text(encoding="utf-8")
            self.assertIn("<svg", svg)
