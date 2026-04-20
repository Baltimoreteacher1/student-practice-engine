Use these three files together in the same Apps Script project:

- `Code.gs`
- `notebook_generator.gs`
- `NotebookExtractors_Precision.gs`

Important:

- Replace the existing Apps Script file contents with these workspace copies.
- Do not leave an older duplicate copy of `runNotebookGeneratorFlagship()`, `runSession1Only()`, or `runSession2Only()` in a different Apps Script file.
- If an old generator function is still present in `NotebookExtractors_Precision.gs`, Apps Script may keep running that stale version and continue throwing the extractor error.

After replacing the files, run:

- `runNotebookGeneratorFlagship()`

If you only want the standalone notebook builder, you can also run:

- `createFlagshipStudentNotebook()`
