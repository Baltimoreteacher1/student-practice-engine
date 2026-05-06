from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT / "src"
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from run import applyEnvironmentOverrides  # noqa: E402
from utils import parse_trigger_text, sanitize_artifact_path, select_slide_deck  # noqa: E402


class TriggerFlowTests(unittest.TestCase):
    def test_ready_trigger_parses_supported_dates(self) -> None:
        self.assertEqual(parse_trigger_text("Ready 4.7.2026")["date"], "2026-04-07")
        self.assertEqual(parse_trigger_text("Ready 04.07.2026")["date"], "2026-04-07")
        self.assertEqual(parse_trigger_text("Ready 2026-04-07")["date"], "2026-04-07")
        self.assertEqual(parse_trigger_text("LP: 4.7.2026")["date"], "2026-04-07")

    def test_deck_selection_prefers_most_recent_relevant_lesson_deck(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            folder = Path(tmp_dir)
            notebook = folder / "Student Notebook.pptx"
            older_lesson = folder / "Editable Lesson Presentation_Old.pptx"
            newer_lesson = folder / "Editable Lesson Presentation_New.pptx"
            for path in (notebook, older_lesson, newer_lesson):
                path.write_bytes(b"placeholder")
            os.utime(notebook, (1, 1))
            os.utime(older_lesson, (10, 10))
            os.utime(newer_lesson, (20, 20))
            selected = select_slide_deck(folder)
            self.assertEqual(selected.name, "Editable Lesson Presentation_New.pptx")

    def test_sanitize_artifact_path_uses_relative_path_inside_workspace(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            deck = root / "lesson-plan-engine" / "INBOX" / "sample_input_slides.pptx"
            deck.parent.mkdir(parents=True, exist_ok=True)
            deck.write_bytes(b"placeholder")
            self.assertEqual(
                sanitize_artifact_path(deck, root),
                "lesson-plan-engine/INBOX/sample_input_slides.pptx",
            )

    def test_sanitize_artifact_path_uses_filename_outside_workspace(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir) / "workspace"
            external_dir = Path(tmp_dir) / "external"
            root.mkdir(parents=True, exist_ok=True)
            external_dir.mkdir(parents=True, exist_ok=True)
            deck = external_dir / "sample_input_slides.pptx"
            deck.write_bytes(b"placeholder")
            self.assertEqual(sanitize_artifact_path(deck, root), "sample_input_slides.pptx")

    def test_environment_overrides_update_watch_and_output_contract(self) -> None:
        config = {
            "input": {
                "slides_dir": "../lesson-plan-engine/INBOX",
                "agenda_dir": "../lesson-plan-engine/INBOX/agenda",
                "notes_dir": "../lesson-plan-engine/INBOX/notes",
            },
            "extracted": {
                "raw_slide_text": "extracted/raw_slide_text.json",
                "normalized_lesson": "extracted/normalized_lesson.json",
                "source_fidelity_map": "extracted/source_fidelity_map.json",
            },
            "output": {
                "json": "output/lesson_plan.json",
                "markdown": "output/lesson_plan.md",
                "docx": "output/lesson_plan.docx",
                "validation_report": "output/validation_report.md",
            },
        }
        with tempfile.TemporaryDirectory() as tmp_dir:
            watch_dir = Path(tmp_dir) / "watch"
            output_dir = Path(tmp_dir) / "output"
            watch_dir.mkdir(parents=True, exist_ok=True)
            output_dir.mkdir(parents=True, exist_ok=True)
            old_watch = os.environ.get("EDUWONDERLAB_WATCH_DIR")
            old_output = os.environ.get("EDUWONDERLAB_OUTPUT_DIR")
            os.environ["EDUWONDERLAB_WATCH_DIR"] = str(watch_dir)
            os.environ["EDUWONDERLAB_OUTPUT_DIR"] = str(output_dir)
            try:
                updated = applyEnvironmentOverrides(config)
            finally:
                if old_watch is None:
                    os.environ.pop("EDUWONDERLAB_WATCH_DIR", None)
                else:
                    os.environ["EDUWONDERLAB_WATCH_DIR"] = old_watch
                if old_output is None:
                    os.environ.pop("EDUWONDERLAB_OUTPUT_DIR", None)
                else:
                    os.environ["EDUWONDERLAB_OUTPUT_DIR"] = old_output

        self.assertEqual(updated["input"]["slides_dir"], str(watch_dir))
        self.assertEqual(updated["input"]["agenda_dir"], str(watch_dir / "agenda"))
        self.assertEqual(updated["input"]["notes_dir"], str(watch_dir / "notes"))
        self.assertEqual(updated["output"]["docx"], str(output_dir / "lesson_plan.docx"))
        self.assertEqual(updated["extracted"]["raw_slide_text"], str(output_dir / "extracted" / "raw_slide_text.json"))


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
