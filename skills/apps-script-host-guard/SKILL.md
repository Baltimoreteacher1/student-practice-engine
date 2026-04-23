---
name: apps-script-host-guard
description: Use this skill for any Google Apps Script task to verify the correct host service, prevent environment-specific failures, enforce paste-ready syntax, and reduce runtime mistakes before finalizing code.
---

# Apps Script Host Guard

## Purpose
Use this skill for any Google Apps Script work.

This includes:
- Forms scripts
- Sheets scripts
- Docs scripts
- Slides scripts
- Drive-integrated scripts
- trigger-based scripts
- Apps Script debugging
- Apps Script code review
- preflight validation before finalizing code

This skill exists to prevent avoidable runtime failures caused by wrong host assumptions, unsupported methods, syntax issues, missing setup, or incomplete deployment notes.

## Core rule
Before writing or finalizing Apps Script code, verify which Google Apps Script host is actually required.

Use the correct service:
- `FormsApp` for Google Forms
- `SpreadsheetApp` for Google Sheets
- `DocumentApp` for Google Docs
- `SlidesApp` for Google Slides
- `DriveApp` for Drive operations

If a requested script appears to target the wrong Apps Script environment, warn clearly before finalizing.

## Preflight checklist
Run this mental checklist every time:

1. **Host check**
   - What container or standalone Apps Script project will run this?
   - Does the code match that host?
   - Are any methods being used from the wrong service?

2. **API correctness check**
   - Are all service calls real and appropriate for the target host?
   - Are any methods unsupported, brittle, or commonly misused?
   - Are advanced services or libraries required?

3. **Dependency/setup check**
   - Does this script depend on:
     - triggers
     - form IDs
     - folder IDs
     - file IDs
     - spreadsheet tabs
     - document structure
     - permissions
     - manual setup steps
   - If yes, state that clearly.

4. **Syntax/paste-safety check**
   - Check brace balance.
   - Check parentheses and array/object closure.
   - Check for obvious naming mismatches.
   - Check for incomplete helper functions or missing constants.
   - Keep code paste-ready.

5. **Runtime-risk check**
   - Identify the most likely failure points.
   - Warn about them explicitly when relevant.
   - Do not overclaim certainty when execution depends on external IDs, permissions, or document structure.

## Output rules
- Prefer final, paste-ready Apps Script code.
- Keep comments useful but not bloated.
- Preserve the repo’s current architecture and naming conventions when editing existing code.
- Prefer minimal, high-confidence fixes over broad rewrites.
- If the user asks for a repair, patch the smallest failing surface first.

## Response rules
When delivering Apps Script code, include:
- target host/service
- what the script does
- required setup items
- what was checked
- likely risk points, if any

## Common failure prevention
Actively guard against:
- using `FormsApp` logic in the wrong environment
- using the wrong container assumptions
- unsupported method calls
- missing IDs or trigger setup
- incomplete helper functions
- syntax errors from rushed edits
- code that sounds correct but is not runnable

## Validation standard
Before claiming the script is ready:
- verify the host/service match
- verify major method calls are coherent
- verify the code is structurally complete
- verify setup requirements are disclosed
- state clearly what was validated and what could only be inferred

## Refusal boundary
Do not pretend runtime success was confirmed unless the script was actually executed in the target environment.
