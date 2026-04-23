# Lesson Plan Engine Tasks

## Now

- [x] Decide whether `../codex-lesson-plan-generator/` stays preserved in place or is promoted into this canonical folder.
- [x] Define one durable input contract for `INBOX/` and one durable output bundle contract for `OUTPUT/`.
- [x] Add one documented verification path that covers generation plus validation (`../validation-tools/src/verify_lesson_plan_engine.py`).

## Next

- [x] Promote at least one scrubbed source deck and one approved output into `examples/` and `gold-standards/`.
- [x] Lock the compact lesson-plan reference contract to the approved DOCX structure and keep the benchmark aligned with it.
- [x] Restore the approved SPED initials line and end-of-plan accommodations matrix to the locked reference bundle.
- [x] Lock row-level SPED initials plus source-grounded small-group lesson moves into the compact procedures contract.
- [x] Consolidate lesson-plan logs and validation notes into `logs/`.
- [x] Document any launcher or automation dependencies that must stay stable.
- [x] Remove or archive legacy generated artifacts that still live under `../codex-lesson-plan-generator/output/`.
- [x] Retire the legacy Node exporter and keep `run.py` as the single canonical generator.

## Later

- [ ] Extract reusable modules into `src/` only after the workflow contract is stable.
- [ ] Add shared asset dependencies in `../shared-assets/` where reuse is proven.
