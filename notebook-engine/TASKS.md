# Notebook Engine Tasks

## Now

- [x] Treat the local root notebook runtime as the primary default path.
- [x] Treat `../flagship-notebook-generator/` as the secondary hosted path.
- [x] Add an additive flagship activity layer after the premium layer without changing the core notebook pipeline.
- [x] Scaffold `enhancement/` as the separate polish lane for externally generated notebook bundles.
- [x] Ship a next-generation premium notebook release across both generator output and enhancement output.
- [ ] Define one durable `INBOX/` to `OUTPUT/` bundle contract for recurring notebook runs.
- [x] Add one documented quality gate path that must pass before outputs are treated as final.
- [x] Tighten the premium QA gate so correct-but-bland notebook output is repaired instead of silently passing.
- [ ] Document which local runtime files are primary, support-only, and secondary snapshots.

## Next

- [ ] Promote at least one scrubbed notebook source deck and one approved final notebook set into `examples/` and `gold-standards/`.
- [x] Centralize notebook QA notes and run history in `logs/`.
- [x] Run one real Claude-generated notebook bundle through `enhancement/` and capture the verified repair path.
- [ ] Identify which root notebook scripts should remain preserved versus migrated.
- [ ] Decide how and when the hosted engine snapshot should be synchronized with the primary local runtime.

## Later

- [ ] Extract stable shared modules into `src/`.
- [ ] Move proven reusable templates or reference art into `../shared-assets/`.
