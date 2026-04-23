# Lesson Plan Engine AGENTS

## Mission

Keep lesson-plan generation deterministic, source-faithful, and easy to rerun.

## Runtime

- Canonical workflow home: this folder
- Active implementation: `../codex-lesson-plan-generator/`
- Default launcher: `../Generate Lesson Plan.command`
- Default verifier: `../validation-tools/src/verify_lesson_plan_engine.py`

## Working rules

- Read `README.md`, `SPEC.md`, and `TASKS.md` here before changing workflow behavior.
- Make durable workflow-contract changes here, even when the implementation change lands in `../codex-lesson-plan-generator/`.
- Treat source lesson materials and approved standards locations as the source of truth.
- Do not invent objectives, standards, or lesson content.
- Keep outputs editable, teacher-usable, and production-ready.
- Preserve the current launcher and inbox/output contract unless there is a clear reason to change it.

## Folder contract

- `INBOX/`: source lesson decks or source bundles waiting to run
- `OUTPUT/`: final lesson-plan artifacts and extracted support files
- `ARCHIVE/`: completed source bundles kept for traceability
- `examples/`: scrubbed reusable examples
- `gold-standards/`: locked reference outputs
- `logs/`: run notes, validation notes, and change records

## Validation

- Run the smallest relevant check first.
- Use `python3 ../validation-tools/src/verify_lesson_plan_engine.py` before calling the workflow healthy after meaningful contract changes.
- If validation cannot run, record the blocker clearly in `logs/`.
