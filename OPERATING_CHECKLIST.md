# Operating Checklist

Use this file for the shortest reliable path through the workspace.

## 1. Lesson plan intake -> build -> output

### Intake

- Put the source deck in `lesson-plan-engine/INBOX/` for the repo-native path.
- The desktop launcher `Generate Lesson Plan.command` already points to `lesson-plan-engine/INBOX/` and `lesson-plan-engine/OUTPUT/` by default.

### Build

- Launcher path:
  - `./Generate Lesson Plan.command`
- Direct path:
  - `python3 codex-lesson-plan-generator/run.py --deck <path-to-source.pptx>`

### Output

- Final artifacts land in `lesson-plan-engine/OUTPUT/`
- Expected bundle members:
  - `lesson_plan.docx`
  - `lesson_plan.md`
  - `lesson_plan.json`
  - `validation_report.md`
  - `extracted/raw_slide_text.json`
  - `extracted/normalized_lesson.json`
  - `extracted/source_fidelity_map.json`

### Archive

- Move accepted source decks or source bundles into `lesson-plan-engine/ARCHIVE/`
- Promote only scrubbed, approved references into `lesson-plan-engine/examples/` or `lesson-plan-engine/gold-standards/`

## 2. Notebook intake -> build -> output

### Intake

- Human-facing daily launcher path:
  - `./Launch Notebook Inbox.command`
  - default inbox: `~/Documents/Chatgpt/Notebook/Notebook Inbox`
  - default output: `~/Documents/Chatgpt/Notebook/Notebook Output`
  - default archive: `~/Documents/Chatgpt/Notebook/Notebook Archive`
- Repo-native path:
  - `notebook-engine/INBOX/`
  - `notebook-engine/OUTPUT/`
  - `notebook-engine/ARCHIVE/`

### Build

- Daily launcher path:
  - `./Launch Notebook Inbox.command`
- Direct single-deck path:
  - `python3 notebook_engine.py run <path-to-source.pptx> --output-dir notebook-engine/OUTPUT/<job>`
- Direct single-deck smoke path:
  - `python3 notebook_engine.py run <path-to-source.pptx> --offline --output-dir notebook-engine/OUTPUT/<job>`
- Repo-native inbox path:
  - `python3 notebook_folder_runner.py --input-dir notebook-engine/INBOX --output-dir notebook-engine/OUTPUT --archive-dir notebook-engine/ARCHIVE --offline`

### Output

- Generated notebook bundles land in `notebook-engine/OUTPUT/` or the user Documents output folder used by the launcher
- Check for:
  - rendered `.pptx`
  - `notebook_plan.json`
  - `source_deck.json`
  - `quality_report.json`

### Archive

- Move accepted source decks or bundles into `notebook-engine/ARCHIVE/`
- Promote only scrubbed, approved references into `notebook-engine/examples/` or `notebook-engine/gold-standards/`

## 3. Notebook enhancement usage

### When to use it

- Use `notebook-engine/enhancement/` when a notebook already exists and needs polish, repair, or rerendering
- Do not use it to replace the core notebook generator

### Intake

- Preferred bundle input: `notebook-engine/enhancement/INBOX/<job>/notebook_plan.json`
- Strongly preferred support file: `source_deck.json`
- Raw `.pptx` notebooks can also be dropped directly into `notebook-engine/enhancement/INBOX/`

### Build

- Launcher path:
  - `./notebook-engine/enhancement/Run Notebook Enhancement.command`
- Direct path:
  - `python3 notebook-engine/enhancement/src/run_enhancement_inbox.py --inbox-dir notebook-engine/enhancement/INBOX --output-dir notebook-engine/enhancement/OUTPUT`

### Output and archive

- Review `enhancement_report.json` or `pptx_polish_report.json`
- Move accepted bundles into `notebook-engine/enhancement/ARCHIVE/`

## 4. Validation quick paths

- Lesson plans:
  - `python3 validation-tools/src/verify_lesson_plan_engine.py`
- Notebooks:
  - `python3 validation-tools/src/verify_notebook_engine.py`
- Focused notebook regression tests:
  - `python3 -m pytest -q tests/test_notebook_session1_pivot.py tests/test_notebook_engine_publisher_polish.py`

## 5. Before calling a run final

- Confirm the source deck or bundle is the real source of truth
- Confirm outputs are editable and readable
- Confirm no placeholder copy or generic filler leaked into the deliverable
- Confirm the output folder contains the expected sidecar files
- Record any blocker or manual-review requirement in the relevant `logs/` folder
- Confirm you are using an active runtime path, not a preserved local one-off from `shared-assets/archive/`
