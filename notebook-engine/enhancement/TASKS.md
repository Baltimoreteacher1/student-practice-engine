# Notebook Enhancement Tasks

## Now

- [x] Define the enhancement bundle contract and rerender handoff.
- [x] Add a compact premium polish runner for existing notebook bundles.
- [x] Verify one real Claude-generated notebook bundle through the enhancement workflow.
- [x] Ship the v-next enhancement release with stronger QA enforcement and a larger premium PPTX polish pack.
- [x] Keep enhancement inbox runs resilient so one broken bundle does not block the rest of the queue.
- [x] Fix raw PPTX lesson-context extraction so ESOL supports and bonus slides use real lesson vocabulary instead of notebook headers.
- [x] Make inbox summaries report review-needed output honestly instead of treating every non-failure as a clean success.
- [x] Add a post-output audit loop for raw PPTX jobs so the enhancer checks the produced slides before it labels the deck premium.
- [x] Strengthen the added premium PPTX slides with more visual and student-friendly scaffolds instead of relying mostly on blank response boxes.
- [x] Add deck-wide audit categories so low-impact or weakly-supported PPTX outputs cannot slip through on a light self-check.
- [x] Add layout-pressure auditing and simplify crowded premium bonus slides so the produced PPTX has fewer text-box stacks and stronger visual anchors.
- [ ] Decide which enhancement repairs should stay plan-level versus later become PPTX-level adjustments.

## Next

- [ ] Add one gold-standard before and after example bundle.
- [x] Capture one verified enhancement run in `logs/RUN_LOG.md`.
- [x] Add stronger visual-balance heuristics using rendered quality report data when it is available.

## Later

- [ ] Extract proven enhancement helpers into reusable notebook modules.
- [ ] Add a launcher or folder-runner path if enhancement becomes a recurring production workflow.
