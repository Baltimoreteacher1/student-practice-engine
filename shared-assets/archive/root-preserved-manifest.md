# Root Preserved Artifact Manifest

These items were intentionally moved out of the workspace root to keep the production surface clean.

## Preserved local one-off scripts

Stored in `shared-assets/archive/local-preserved-root-artifacts/scripts/`:

- `adapt_notebooks_esol_newcomer.py`
- `build_premium_student_notebook_93_session2.py`
- `build_tpt_flagship_session2_canoe.py`
- `build_two_session_notebooks.py`
- `cleanup_neft_ppt_v3.py`
- `final_micro_polish_neft_ppt.py`
- `final_sweep_neft_ppt.py`
- `finalize_unit9_test_review_v5.py`
- `fix_requested_slides_v3.py`
- `fix_requested_slides_v4.py`
- `fix_slide12_overflow.py`
- `publisher_polish_neft_ppt.py`
- `rebuild_notebooks_single_object_visuals.py`
- `redesign_edtech_session1.py`
- `refine_neft_ppt_layout.py`
- `refine_unit9_test_review_v4.py`
- `update_neft_ppt.py`
- `upgrade_notebooks_visuals_wida.py`
- `upgrade_unit9_test_review.py`

These were preserved because they capture historical classroom-specific work, but they are not part of the supported lesson-plan or notebook production runtimes.

## Preserved prompts

Stored in `shared-assets/archive/local-preserved-root-artifacts/prompts/`:

- `final_codex_prompt.txt`

## Preserved preview renders

Stored in `shared-assets/archive/local-preserved-root-artifacts/previews/`:

- `pptx_preview/lesson54/Lesson 5.4 - Session 1 Student Notebook.pptx.png`

## Preserved scratch previews and renders

Stored in `shared-assets/archive/local-preserved-root-artifacts/renders/`:

- `preview_slide_13.pptx`
- `preview_slide_5.pptx`
- `unit9_graph1_v4.png`
- `unit9_graph1_v5.png`
- `unit9_graph2_v4.png`
- `unit9_graph2_v5.png`
- `unit9_slide10_render.png`
- `unit9_slide15_v5.png`
- `unit9_slide16_v5.png`

## Reuse rule

If any preserved item needs to become active again, do not move it back to root as-is. First parameterize paths, document the purpose, validate the output, and promote it into the correct canonical folder.
