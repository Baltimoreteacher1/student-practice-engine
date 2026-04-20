# Validation Tools

Canonical workspace home for validation, QA, extraction, and audit workflows.

Current validation-related material preserved in place includes:

- `../tests/`
- `../codex-lesson-plan-generator/tests/`
- notebook smoke and inspection folders at workspace root

Canonical lesson-plan verification command:

```bash
python3 validation-tools/src/verify_lesson_plan_engine.py
```

This verifier runs the preserved generator test suite, compiles the active Python entrypoints, launches `../Generate Lesson Plan.command` against the sample deck, and confirms the expected output bundle is produced in a fresh temporary output folder.

Use this folder to consolidate durable validators, extractors, reference cases, and QA notes.
