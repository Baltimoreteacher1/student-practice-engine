# Notebook Enhancement Spec

## Purpose

Polish existing notebook bundles to premium classroom-publishing quality without rebuilding the notebook engine.

## Inputs

- draft notebook bundle folders in `INBOX/`
- `notebook_plan.json` as the primary repair target
- `source_deck.json` when source-aware repairs are needed
- optional existing notebook `.pptx` files as references

## Outputs

- repaired notebook bundles in `OUTPUT/`
- `enhancement_report.json` with premium QA findings and repairs applied
- `qualityTier` reporting that distinguishes `premium`, `enhanced`, and `fail`
- render-ready `notebook_plan.json` for the current notebook engine
- accepted bundles in `ARCHIVE/`
- for raw PPTX jobs, `pptx_polish_report.json` must include an explicit post-output audit of the produced deck
- the raw PPTX audit must include deck-wide enhancement evidence, not just extension-slide checks

## Workflow

- `Ingest Existing Bundle -> Premium QA Rubric -> Targeted Plan Repairs -> Render With Current Notebook Engine -> Final QA -> Archive`

For raw PPTX enhancement jobs:

- `Ingest Raw PPTX -> Premium Editorial Polish -> Lesson-Anchored Premium Slide Pack -> Reopen + Audit Produced PPTX -> Optional Safer Rerun -> Quality Tier Gate -> Export`

## Non-negotiables

- do not rebuild the notebook system here
- preserve source fidelity and editable output
- repair weak quality before treating a notebook as ready
- distinguish correct-but-bland from premium-and-finished
- prefer targeted repair over broad replacement

## Quality contract

Enhancement work must enforce the notebook premium quality standard:

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

Hard fails include generic filler, placeholder wording, lesson-untethered activities, weak directions, crowded copy, repetitive flat structure, and pages that feel technically correct but below premium benchmark quality.

V-next enhancement expectations:

- the enhancer should visibly strengthen existing PPTX notebooks, not merely clean them up
- when lesson signal is strong, the enhancer should add multiple lesson-anchored premium slides and editorial framing while preserving editability
- when lesson signal is strong enough for premium extension work, the added slide mix should vary by lesson evidence instead of defaulting to one repeated slide pack
- the PPTX path should extract real lesson focus and vocabulary from notebook structures, not table headers or product chrome
- the PPTX path should inspect the actual produced deck before calling it premium
- premium extension slides should add visible classroom scaffolds, not just more writing space: cue cards, sort cards, talk moves, step ladders, and other editable visual supports
- premium decisions should reflect deck-wide audit categories such as original-slide coverage, enhancement magnitude, lesson anchoring, scaffold density, and student supports
- premium decisions should fail or downgrade crowded extension layouts; layout pressure must be audited from the produced PPTX using added-slide shape count, content-text density, and visual balance heuristics
- added premium slides should prefer larger visual anchors and fewer stronger labels over stacks of tiny chips and text boxes
- the inbox summary should distinguish premium-complete output from output that still needs manual review
- when lesson signal is weak, the enhancer should stop at safe premium polish and report the lower quality tier honestly
