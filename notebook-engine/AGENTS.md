# Notebook Engine AGENTS

## Mission

Produce source-faithful, editable, production-ready student notebooks with predictable run paths.

## Runtime

- Canonical workflow home: this folder
- Primary local runtime:
  - `../notebook_engine.py`
  - `../notebook_engine_app.py`
  - `../notebook_folder_runner.py`
  - `../notebook_launchers.py`
- Secondary hosted runtime: `../flagship-notebook-generator/`
- Default verifier: `../validation-tools/src/verify_notebook_engine.py`

## Default notebook build contract

- Use the compressed notebook structure by default unless the user explicitly asks for the full notebook.
- Default compressed Session 1 sequence:
  1. `Objectives + Session Map`
  2. `Be Curious`
  3. `Vocabulary + Reference Tool`
  4. `Guided Problem`
  5. `Interactive Activity`
  6. `Best-Fit Interactive Review`
- Preserve source wording for the guided/source problem when available.
- Do not invent vocabulary or swap in generic academic words when the source lesson already provides the math language.

## Working rules

- Read `README.md`, `SPEC.md`, and `TASKS.md` here before changing workflow behavior.
- Change the real implementation in the preserved root runtime unless the task is explicitly about the hosted app.
- Use `enhancement/` only for polish and repair of already-generated notebook bundles.
- Keep outputs editable, readable, and visually cohesive.
- Preserve launcher, folder-runner, and archive contracts when users rely on automation.

## Folder contract

- `INBOX/`: source decks or notebook jobs waiting to run
- `OUTPUT/`: generated notebook bundles
- `ARCHIVE/`: completed source decks or bundles
- `enhancement/`: notebook polish lane for existing notebook bundles
- `examples/`: scrubbed notebook examples
- `gold-standards/`: locked reference notebooks
- `logs/`: run notes, validation notes, and workflow history

## Validation

- Run the smallest relevant check first.
- Use `python3 ../validation-tools/src/verify_notebook_engine.py` before calling the workflow healthy after meaningful contract changes.
- For notebook-engine regressions, also use the targeted root tests that cover the touched area.
