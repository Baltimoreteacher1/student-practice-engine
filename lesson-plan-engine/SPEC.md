# Lesson Plan Engine Spec

## Purpose

Generate deterministic, classroom-ready lesson plans from source lesson materials with strong source fidelity and validation.

## Current implementation sources

- `../codex-lesson-plan-generator/`
- `../Generate Lesson Plan.command`

The preserved implementation defaults now point at this folder's `INBOX/` and `OUTPUT/` paths so the canonical runtime contract lives here even before full code migration.

Canonical runtime: `../codex-lesson-plan-generator/run.py`

## Inputs

- raw lesson decks or lesson source bundles in `INBOX/`
- reference examples in `examples/`
- approved final references in `gold-standards/`

## Outputs

- final lesson-plan artifacts in `OUTPUT/`
- completed source bundles in `ARCHIVE/`
- run and QA notes in `logs/`

Current locked teacher-facing reference bundle:

- title line with teacher name, session title, and rendered date
- `IEP Students` line with the active initials for that lesson
- `DO NOW`
- `STANDARDS` with a 3-column prerequisite / target / extension table
- `OBJECTIVES` with content objective, before/after self-check line, language objective, and before/after self-check line
- `LESSON PROCEDURES` with the 6-column compact procedures table
- `LESSON PROCEDURES` keeps every active IEP/SPED initial visible in the collaborative and independent modification cells
- lessons with small-group profiles include a source-grounded small-group move tied to the slide or book task
- `VOCABULARY` with the 7-column vocabulary table
- `IEP ACCOMMODATIONS MATRIX` with the active initials and lesson-specific accommodations

## Non-negotiables

- source deck and source lesson text are the truth
- standards and objectives must come from the approved source location
- validation is required before a deliverable is treated as final
- examples are editable references; gold standards are locked exemplars

## Migration target

When code is intentionally promoted into this canonical folder, place durable implementation modules in `src/` and keep the preserved generator as the reference until parity is verified.
