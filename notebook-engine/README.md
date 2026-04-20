# Notebook Engine

Canonical workspace home for student notebook generation.

## Runtime decision

Primary default runtime:

- `../notebook_engine.py`
- `../notebook_engine_app.py`
- `../notebook_folder_runner.py`
- `../notebook_launchers.py`

Secondary runtime:

- `../flagship-notebook-generator/`

Use the root local runtime for normal daily notebook generation, debugging, inbox automation, and production-path fixes.

Use `flagship-notebook-generator/` only when the task is explicitly about the hosted upload workflow, API service, or web UI.

## Where to change code

- change durable workflow contracts here in `notebook-engine/`
- change local runtime behavior in the preserved root notebook files
- change hosted behavior in `../flagship-notebook-generator/`
- use `enhancement/` for polish-only work on existing notebook bundles that should be upgraded without changing the core generator

## Daily workflow

1. Read `SPEC.md` and `TASKS.md`.
2. Default to the local runtime unless the task is clearly hosted-web work.
3. Reproduce with the smallest local command, app flow, or folder-runner path.
4. Fix the real implementation first.
5. Record any durable contract change back in this canonical folder.
