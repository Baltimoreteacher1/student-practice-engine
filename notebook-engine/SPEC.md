# Notebook Engine Spec

## Purpose

Generate polished, source-faithful student notebook outputs from source lesson materials while preserving lesson fidelity and passing quality validation.

## Current implementation sources

Primary default runtime:

- `../notebook_engine.py`
- `../notebook_engine_app.py`
- `../notebook_folder_runner.py`
- `../notebook_launchers.py`

Secondary hosted runtime:

- `../flagship-notebook-generator/`

## Inputs

- raw lesson decks in `INBOX/`
- reference examples in `examples/`
- approved final notebook references in `gold-standards/`

## Default session build contract

Unless the prompt explicitly asks for the full notebook, default to the compressed Session 1 notebook sequence:

1. `Objectives + Session Map`
2. `Be Curious`
3. `Vocabulary + Reference Tool`
4. `Guided Problem`
5. `Interactive Activity`
6. `Best-Fit Interactive Review`

Rules for this default path:

- preserve the guided/source problem wording verbatim when available
- keep `Be Curious` fully intact instead of compressing it further
- anchor vocabulary to the lesson math content, not to generic launch or mindset slides
- choose the interactive activity based on the lesson type
- use the strongest discussion-ready closing review rather than a decorative filler page

## Outputs

- final notebook bundles in `OUTPUT/`
- default bundle includes the compressed Session 1 notebook
- Session 2 notebook output is optional and should appear only when the prompt, workflow, or notebook plan explicitly calls for it
- completed source decks and bundles in `ARCHIVE/`
- run history and QA notes in `logs/`

## Planning pipeline

- `Source Extract -> Content Map -> Template Build -> Premium Layer -> Flagship Activity Layer -> Notebook Premium QA Enforcement -> QA -> Export`

Current release target:

- next-generation premium notebook output with stronger section headers, richer premium panel layouts, and a more visibly authored flagship-activity arc
- next-generation enhancement output that can upgrade externally generated notebooks into premium, editable PPTX editions without replacing the core engine

The flagship activity layer is additive only:

- preserve the current engine, templates, and render/export path
- keep editable PPTX output with native shapes, text, and tables
- add 2 to 4 lesson-anchored flagship activities per session when the source lesson supports them
- fall back cleanly to the original notebook output when source support is thin

V-next premium output expectations:

- use stronger premium layouts such as evidence ladders and real-world transfer studios when the lesson supports them
- use distinct full-spread compositions for the strongest premium layouts instead of reusing one split-card grammar everywhere
- keep the notebook visually cohesive while making the premium layer feel materially newer than the prior release
- preserve source fidelity and editability even when the premium finish becomes more ambitious

## Enhancement workspace

Use `enhancement/` for polish-only upgrades to notebooks that were already generated elsewhere, including Claude-created draft notebook bundles.

Enhancement pipeline:

- `Existing Notebook Bundle -> Premium QA Rubric -> Targeted Plan Repairs -> Render With Current Notebook Engine -> QA -> Export`

Enhancement rules:

- preserve the existing notebook generator and render path
- repair weak notebook plans instead of rebuilding the notebook system
- prefer targeted copy, support, and structure upgrades over broad redesign
- keep the output editable and source-anchored
- fail clearly when the bundle is too thin to polish safely
- distinguish `premium` from merely `enhanced` in enhancement reporting instead of treating every successful polish pass as equal

## Non-negotiables

- source deck problems and lesson flow stay anchored to the source
- compressed Session 1 is the default deliverable unless the user explicitly requests a different structure
- Session 2 output should be generated only when the workflow explicitly calls for it
- quality reporting is required before final delivery
- launcher and folder-runner contracts must remain stable when automation depends on them
- local runtime is the default daily notebook workflow unless hosted behavior is explicitly required

## Active change locations

- local notebook implementation changes belong in the preserved root notebook runtime files
- hosted notebook implementation changes belong in `../flagship-notebook-generator/`
- durable workflow guidance, examples, task tracking, and migration notes belong here

## Migration target

Promote only stable notebook modules into `src/`. Preserve the current root scripts and the flagship app as the reference until migrated behavior is verified.
