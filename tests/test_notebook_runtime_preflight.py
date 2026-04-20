import tempfile
import unittest
from pathlib import Path

from notebook_engine import (
    ROOT,
    notebook_runtime_preflight,
    slugify,
    validate_json_artifact_path,
    validate_output_dir_path,
)


class NotebookRuntimePreflightTests(unittest.TestCase):
    def test_notebook_runtime_preflight_rejects_missing_source(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            missing_path = Path(tmp_dir) / "missing-deck.pptx"

            with self.assertRaisesRegex(RuntimeError, "Source PPTX was not found"):
                notebook_runtime_preflight(missing_path)

    def test_notebook_runtime_preflight_rejects_non_pptx_source(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "notes.txt"
            source_path.write_text("not a deck", encoding="utf-8")

            with self.assertRaisesRegex(RuntimeError, r"must be a \.pptx deck"):
                notebook_runtime_preflight(source_path)

    def test_notebook_runtime_preflight_accepts_existing_pptx_source(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "lesson-deck.pptx"
            source_path.write_bytes(b"placeholder")

            validated_source, output_dir = notebook_runtime_preflight(source_path)

            self.assertEqual(validated_source, source_path.resolve())
            self.assertEqual(output_dir, (ROOT / f"{slugify(source_path.stem)}-notebook-build").resolve())

    def test_validate_output_dir_path_rejects_file_target(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            file_path = Path(tmp_dir) / "output.json"
            file_path.write_text("{}", encoding="utf-8")

            with self.assertRaisesRegex(RuntimeError, "points to a file"):
                validate_output_dir_path(file_path)

    def test_validate_json_artifact_path_requires_json_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            deck_path = Path(tmp_dir) / "source_deck.txt"
            deck_path.write_text("{}", encoding="utf-8")

            with self.assertRaisesRegex(RuntimeError, r"must be a \.json file"):
                validate_json_artifact_path(deck_path, label="Source deck JSON")
