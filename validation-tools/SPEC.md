# Validation Tools Spec

## Purpose

Provide durable validation, QA, extraction, and audit workflows that can be reused across lesson plans, notebooks, and classroom automation outputs.

## Current implementation sources

- `../codex-lesson-plan-generator/tests/`
- `../tests/`
- preserved notebook smoke and inspection folders at workspace root

## Inputs

- source fixtures and repro cases in `examples/`
- reusable validators in `src/`
- shared extraction helpers in `extractors/`

## Outputs

- repeatable validation results
- documented failure modes
- durable QA notes in `logs/`

## Non-negotiables

- validation should run on the smallest credible repro first
- source fidelity and production quality checks matter as much as syntax checks
- preserve useful smoke cases until they are either consolidated or clearly obsolete
- do not rely on one-off output folders as the only evidence of correctness

## Migration target

Consolidate shared validators and extractors here only after they are proven reusable across more than one workflow.
