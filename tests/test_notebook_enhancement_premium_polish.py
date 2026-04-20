import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
ENHANCEMENT_SRC = ROOT / "notebook-engine" / "enhancement" / "src"
if str(ENHANCEMENT_SRC) not in sys.path:
    sys.path.insert(0, str(ENHANCEMENT_SRC))

from premium_polish import (  # type: ignore[import-not-found]
    evaluate_notebook_premium_quality,
    repair_notebook_quality_issues,
    run_notebook_enhancement,
)


def sample_plan() -> dict:
    return {
        "session_1": {
            "slides": [
                {
                    "kind": "practice",
                    "title": "Practice",
                    "subtitle": "",
                    "primary_text": "Students should use the lesson idea.",
                    "secondary_text": "Use the space below to do the work and complete the activity.",
                    "bullets": [],
                    "tasks": [],
                    "response_prompt": "Do the work.",
                    "sentence_starters": [],
                    "vocabulary": [],
                    "activity_name": "",
                    "activity_instructions": "Complete the activity.",
                    "source_slide_numbers": [1],
                }
            ]
        },
        "session_2": {
            "slides": [
                {
                    "kind": "reflection",
                    "title": "Reflection",
                    "subtitle": "",
                    "primary_text": "Think about the lesson.",
                    "secondary_text": "",
                    "bullets": [],
                    "tasks": [],
                    "response_prompt": "",
                    "sentence_starters": [],
                    "vocabulary": [],
                    "activity_name": "",
                    "activity_instructions": "",
                    "source_slide_numbers": [2],
                }
            ]
        },
    }


def sample_deck() -> dict:
    return {
        "slides": [
            {
                "slide_number": 1,
                "title": "Compare the area model and equation",
                "text": "Use the area model to explain which strategy matches the equation.",
                "text_items": ["Compare the area model.", "Explain which strategy matches the equation."],
                "problem_texts": ["Which strategy matches the area model and equation?"],
                "notes": "",
            },
            {
                "slide_number": 2,
                "title": "Write about your strategy",
                "text": "Explain how the model shows the multiplication relationship.",
                "text_items": ["Explain how the model shows the relationship."],
                "problem_texts": ["How does the model show the multiplication relationship?"],
                "notes": "",
            },
        ]
    }


def weak_plan() -> dict:
    return {
        "session_1": {
            "slides": [
                {
                    "kind": "practice",
                    "title": "Explain the strategy",
                    "subtitle": "",
                    "primary_text": "Use the model and the equation to show your thinking.",
                    "secondary_text": "Write your reasoning in the workspace.",
                    "bullets": [],
                    "tasks": [],
                    "response_prompt": "Tell what you notice and why it matters.",
                    "sentence_starters": [],
                    "vocabulary": [],
                    "activity_name": "",
                    "activity_instructions": "Use the source idea to explain your answer.",
                    "source_slide_numbers": [1],
                }
            ]
        },
        "session_2": {
            "slides": []
        },
    }


class NotebookEnhancementPremiumPolishTests(unittest.TestCase):
    def test_evaluate_detects_generic_and_bland_quality_failures(self) -> None:
        report = evaluate_notebook_premium_quality({"plan": sample_plan(), "deck": sample_deck()})

        self.assertFalse(report["passed"])
        self.assertTrue(report["hardFails"])
        self.assertEqual(report["categories"]["lessonAdaptation"]["status"], "fail")
        self.assertEqual(report["qualityTier"], "fail")
        self.assertIn("generic filler text remains", " ".join(report["hardFails"]))

    def test_evaluate_marks_correct_but_bland_quality_as_weak(self) -> None:
        report = evaluate_notebook_premium_quality({"plan": weak_plan(), "deck": sample_deck()})

        self.assertFalse(report["passed"])
        self.assertFalse(report["hardFails"])
        self.assertEqual(report["qualityTier"], "weak")
        self.assertEqual(report["categories"]["lessonAdaptation"]["status"], "weak")
        self.assertEqual(report["categories"]["benchmarkFinish"]["status"], "weak")

    def test_repair_replaces_generic_copy_and_adds_supports(self) -> None:
        repaired = repair_notebook_quality_issues(
            {"plan": sample_plan(), "deck": sample_deck()},
            evaluate_notebook_premium_quality({"plan": sample_plan(), "deck": sample_deck()}),
        )
        slide = repaired["plan"]["session_1"]["slides"][0]

        self.assertNotEqual(slide["activity_instructions"], "Complete the activity.")
        self.assertTrue(slide["sentence_starters"])
        self.assertTrue(slide["tasks"])
        self.assertIn("area model", slide["response_prompt"].lower())
        self.assertTrue(repaired["qaReport"]["repairsApplied"])

    def test_run_notebook_enhancement_writes_polished_bundle_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            bundle_dir = Path(tmp_dir) / "bundle"
            output_dir = Path(tmp_dir) / "output"
            bundle_dir.mkdir()
            (bundle_dir / "notebook_plan.json").write_text(__import__("json").dumps(sample_plan()), encoding="utf-8")
            (bundle_dir / "source_deck.json").write_text(__import__("json").dumps(sample_deck()), encoding="utf-8")

            report = run_notebook_enhancement(bundle_dir, output_dir)

            self.assertTrue((output_dir / "notebook_plan.json").exists())
            self.assertTrue((output_dir / "enhancement_report.json").exists())
            self.assertTrue((output_dir / "rerender_command.txt").exists())
            self.assertIn("finalReport", report)
            self.assertTrue(report["finalReport"]["repairsApplied"])
