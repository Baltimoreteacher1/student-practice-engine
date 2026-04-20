# AGENTS.md

## Mission

Build deterministic, classroom-ready generators that transform source lesson materials into polished final artifacts.

## Locked contracts

- Source lesson PPTX is instructional source of truth.
- Standards must come only from the Learning Target slide.
- Preserve source fidelity and lesson flow.
- Prefer fixed structure over freeform generation.
- Fail loudly when required inputs are missing or ambiguous.
- Produce polished final artifacts, not draft-style output.
- Keep launcher-trigger behavior and folder contracts stable.

Reference contracts:

- `PROMPT.md` for canonical `run.py` workflow behavior and required extract fields.
- `CLAUDE.md` for canonical `run.py` artifact shape and runtime constraints.

## Runtime commands

Project setup:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

Primary pipeline runs:

```bash
python run.py
python run.py --deck examples/sample_input_slides.pptx
python run.py --config config/generator_config.json --school-defaults config/school_defaults.json
```

Test and verification commands:

```bash
python3 -m pytest tests -q
python3 -m pytest tests/test_structure.py tests/test_standards_rule.py tests/test_supports_rule.py tests/test_trigger_flow.py -q
```

## Engineering rules

- Use modular files with focused functions.
- Prefer readable, high-confidence code.
- Use schemas and validation where helpful.
- Do not silently skip failed steps.
- Keep paths relative and portable.
- Update README when behavior changes.
- Add or update tests when core logic changes.

## Output rules

- Always generate JSON + Markdown + DOCX when supported by the run.
- Include a validation report.
- Include a source fidelity appendix.

## Done criteria

Before finalizing changes in this repo:

1. Required tests are passing or a concrete blocker is documented.
2. Structure and standards-source constraints still hold.
3. Output contract remains intact (JSON + Markdown + DOCX + validation report).
4. Any behavior changes are reflected in `README.md`.

## Style

- Deterministic
- Professional
- Classroom-ready
- Maintainable
