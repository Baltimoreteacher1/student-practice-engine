# Notebook Enhancement AGENTS

## Mission

Repair or upgrade existing notebook bundles without replacing the core notebook generator.

## Working rules

- Read `README.md`, `SPEC.md`, and `TASKS.md` here before changing enhancement behavior.
- Keep the enhancement lane additive: polish, repair, rerender, and report.
- Do not turn enhancement work into a second notebook engine.
- Preserve source fidelity, editability, and the existing notebook structure unless the source bundle is already wrong and needs repair.
- Fail clearly when the bundle is too thin to polish safely.

## Input contract

- Preferred input: `INBOX/<job>/notebook_plan.json`
- Strongly preferred support file: `INBOX/<job>/source_deck.json`
- Optional support files: `quality_report.json`, existing `.pptx` notebook files
- Raw `.pptx` notebooks may also be dropped directly into `INBOX/`

## Output contract

- Polished bundle in `OUTPUT/<job>/`
- Repair report in `enhancement_report.json` or `pptx_polish_report.json`
- Rerendered PPTX output when the source bundle is complete enough to rerender safely
