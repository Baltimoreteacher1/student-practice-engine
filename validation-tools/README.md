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

Deterministic lesson-plan benchmark command:

```bash
python3 validation-tools/src/benchmark_lesson_plan_quality.py
```

Canonical notebook verification command:

```bash
python3 validation-tools/src/verify_notebook_engine.py
```

Deterministic notebook benchmark command:

```bash
python3 validation-tools/src/benchmark_notebook_quality.py
```

This verifier runs the preserved generator test suite, compiles the active Python entrypoints, launches `../Generate Lesson Plan.command` against the sample deck, and confirms the expected output bundle is produced in a fresh temporary output folder.

The lesson-plan benchmark reruns the canonical sample deck and compares the generated output bundle against the promoted gold-standard lesson-plan bundle with date/path normalization where needed.

The notebook verifier runs the focused root notebook regression tests, compiles the active notebook runtime files, and then runs the deterministic notebook benchmark cases so source fidelity, compressed structure, rendered quality-report thresholds, and one exact gold-standard Session 1 reference stay locked.

Use this folder to consolidate durable validators, extractors, reference cases, and QA notes.
