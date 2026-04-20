# Lesson Plan Engine Validation Notes

## Canonical runtime

- Primary generator: `../codex-lesson-plan-generator/run.py`
- Desktop launcher: `../Generate Lesson Plan.command`

## Required Python dependencies

- `python-pptx`
- `python-docx`
- `jsonschema`

Install from:

```bash
pip install -r ../codex-lesson-plan-generator/requirements.txt
```

## Latest verified checks

- `python3 -m pytest -q ../codex-lesson-plan-generator/tests`
- `python3 ../codex-lesson-plan-generator/run.py`
- `python3 ../validation-tools/src/verify_lesson_plan_engine.py`

## Canonical sample artifacts

- Input example: `../examples/sample_input_slides.pptx`
- Example output JSON: `../examples/sample_lesson_plan.json`
- Example output Markdown: `../examples/sample_lesson_plan.md`
- Example validation report: `../examples/sample_validation_report.md`
- Gold-standard DOCX: `../gold-standards/sample_lesson_plan.docx`
