# AGENTS.md

## Mission

Build reliable education workflow tooling that turns source lesson materials into production-ready classroom artifacts.

## Canonical workspace

Use these folders first for new durable work:

- `lesson-plan-engine/`
- `notebook-engine/`
- `quiz-form-builder/`
- `validation-tools/`
- `shared-assets/`

Existing implementations are preserved in place until they are intentionally migrated:

- `codex-lesson-plan-generator/`
- `flagship-notebook-generator/`
- root notebook scripts such as `notebook_engine.py`, `notebook_engine_app.py`, and `notebook_folder_runner.py`
- legacy root artifacts, smoke runs, and one-off outputs

## Working rules

- Inspect existing code and docs before changing workflow contracts.
- Put new durable docs, specs, examples, and logs in the canonical folders above, not at workspace root.
- Preserve source fidelity. Teacher slides, lesson text, and approved reference materials are the source of truth.
- Prefer extending an existing engine over creating a duplicate pipeline.
- For notebooks generated outside the primary engine, use `notebook-engine/enhancement/` as the polish lane instead of rebuilding the core generator.
- Use `INBOX/`, `OUTPUT/`, and `ARCHIVE/` consistently.
- Treat `gold-standards/` as locked references; copy from them, do not casually edit them.
- Keep implementation code in `src/` or in an existing preserved engine until migration is deliberate.
- Update `SPEC.md`, `TASKS.md`, and the relevant log when workflow behavior changes.

## Notebook Premium Quality Standard

This is a hard quality contract for every notebook-generation or notebook-enhancement task.

Every notebook must be:

- source-faithful
- lesson-adapted
- fully editable
- visually cohesive
- student-usable
- cleanly formatted
- publication-ready

Treat the output as failed and in need of repair if it includes generic filler activities, placeholder wording, bland or repetitive layouts, weak visual hierarchy, crowded text, tiny font, poor spacing, inconsistent alignment, style drift, decorative clutter, lesson-untethered activities, generic supports, template-generated pages, or technically correct but visually underdeveloped slides.

Quality priority order:

1. source fidelity
2. instructional quality
3. student usability
4. visual quality
5. editability
6. premium finish

Default notebook design standard:

- larger readable text
- strong hierarchy
- balanced white space
- clean alignment
- restrained premium color usage
- clear response zones
- interactive-feeling layouts built with editable shapes, text, and tables

Activity standard:

- added activities must feel authored for the exact lesson
- use real lesson vocabulary, representations, examples, misconceptions, and reasoning patterns
- prefer fewer stronger activities over more weaker ones

Writing standard:

- notebook wording must read like polished human publishing copy
- directions must be concise, clear, natural, and teachable
- avoid robotic phrasing, vague prompts, and repetitive wording

Required review pass before finalizing notebook work:

- source fidelity
- lesson adaptation
- activity quality
- vocabulary and support integration
- visual hierarchy
- layout discipline
- typography and readability
- editability
- tone and wording
- benchmark finish

If any category is weak, repair it before finalizing.

## Verification

Before finalizing:

1. Confirm the work landed in the right canonical area.
2. Run the smallest relevant validation, test, or smoke path.
3. Record blockers explicitly if verification cannot run.
4. Summarize what changed, what was verified, and any remaining risk.

## Output discipline

- Production-ready outputs only.
- Fail loudly on missing inputs or ambiguous source.
- Preserve launcher and folder contracts when users rely on automation.
