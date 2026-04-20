# Notebook Enhancement

Use this workspace to polish notebook bundles that were already generated elsewhere without changing the main notebook generator.

This lane is for:

- Claude-generated notebook drafts
- previously rendered notebook bundles that need stronger polish
- targeted premium upgrades that should stay inside the current notebook design system

This lane is not for:

- rebuilding the notebook engine
- replacing the core template system
- redesigning the notebook product from scratch

## Bundle contract

Preferred input bundle in `INBOX/<job>/`:

- `notebook_plan.json` required
- `source_deck.json` strongly recommended
- `quality_report.json` optional
- existing `.pptx` notebook files optional as references

Raw `.pptx` notebooks are also supported:

- drop the `.pptx` file directly into `INBOX/`
- the enhancement runner will create `OUTPUT/<file-stem>/`
- it will save a polished PPTX copy plus `pptx_polish_report.json`
- the PPTX polish pass now adds stronger publisher-style structure, including title backplates, prompt cards, panel tabs, response zones, and 2 to 3 lesson-anchored bonus slides when the source supports them

## Workflow

1. Drop either one bundle folder or one raw notebook `.pptx` into `INBOX/`.
2. Click `Run Notebook Enhancement.command` in this folder, or run:
   `python3 notebook-engine/enhancement/src/premium_polish.py notebook-engine/enhancement/INBOX/<job> --output-dir notebook-engine/enhancement/OUTPUT/<job>`
3. Review `enhancement_report.json`.
4. The inbox runner will rerender automatically when `source_deck.json` is present. Manual rerender command:
   `python3 notebook_engine.py render notebook-engine/enhancement/OUTPUT/<job>/notebook_plan.json --deck notebook-engine/enhancement/OUTPUT/<job>/source_deck.json --output-dir notebook-engine/enhancement/OUTPUT/<job>/rendered`
5. Move accepted bundles to `ARCHIVE/`.
