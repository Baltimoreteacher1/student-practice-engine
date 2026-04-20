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

## Non-negotiables

- source deck and source lesson text are the truth
- standards and objectives must come from the approved source location
- validation is required before a deliverable is treated as final
- examples are editable references; gold standards are locked exemplars

## Migration target

When code is intentionally promoted into this canonical folder, place durable implementation modules in `src/` and keep the preserved generator as the reference until parity is verified.
