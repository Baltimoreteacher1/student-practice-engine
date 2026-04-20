# Structure Rules

## Required structure

For each rendered session, the canonical generator must produce this exact structure in order:

1. `Lesson Information`
2. `Standards and Learning Targets`
3. `Lesson Objective and Student Success Criteria`
4. `Materials and Preparation`
5. `Opening / Warm-Up / Launch`
6. `Mini-Lesson / Modeling / Concept Development`
7. `Guided Practice / Collaborative Learning`
8. `Independent Practice / Application / Stations`
9. `Closure / Exit Ticket / Assessment`
10. `Differentiation, SPED/ESOL Supports, and Teacher Notes`

## Determinism rules

- Keep extraction, mapping, composition, validation, and rendering as separate steps.
- Prefer stable heuristics over freeform generation.
- Fail clearly when the deck cannot support a usable lesson plan.
- Preserve session order from the source deck.

## Canonical runtime rule

- `run.py` defines the primary structure contract.
