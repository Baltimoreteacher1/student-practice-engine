# Workspace Pristine Pass - 2026-04-20

## Goal

Make the workspace root read like a production system instead of a mixed runtime plus scratch area.

## Changes

- added the missing `CODEX_OPERATING_PLAYBOOK.md` referenced by the root `AGENTS.md`
- documented shared archive usage in `shared-assets/archive/README.md`
- added a manifest for preserved root artifacts in `shared-assets/archive/root-preserved-manifest.md`
- moved preserved local one-off scripts, a scratch prompt, and a preview render out of the workspace root into `shared-assets/archive/local-preserved-root-artifacts/`
- updated root workflow docs to point future sessions at the active runtimes and away from archived local artifacts

## Why

The root workspace had become visually and operationally noisy. The real lesson-plan and notebook entrypoints were valid, but they sat next to preserved one-off helpers that were easy to mistake for reusable tooling.

This pass keeps those artifacts available without letting them compete with the supported workflows.
