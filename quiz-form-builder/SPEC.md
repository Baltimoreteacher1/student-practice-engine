# Quiz and Form Builder Spec

## Purpose

Create and maintain durable Google Apps Script builders for quizzes, forms, notebook extraction helpers, and classroom document automation.

## Current implementation sources

- `../apps_script_export_area_triangles/`
- root `.gs` files such as `../Code.gs`, `../NotebookExtractors_Precision.gs`, `../Notebookgrids.gs`, `../notebook_generator.gs`, `../participant_workbook_session1.gs`

## Inputs

- prompt notes, source lesson content, or sheet/form specs in `examples/`
- durable Apps Script source candidates in `apps-script/`

## Outputs

- stable Apps Script project files
- repeatable setup notes and examples
- run and deployment notes in `logs/`

## Non-negotiables

- avoid duplicate function names across separate `.gs` files in the same Apps Script project
- preserve known working entrypoints unless replacement is deliberate
- keep deployment-specific secrets and IDs out of source files
- preserve classroom source fidelity when builders generate lesson artifacts

## Migration target

Promote only the stable, non-duplicated script set into this canonical folder after verifying which preserved root files are still active.
