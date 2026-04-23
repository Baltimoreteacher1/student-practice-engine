# Validation Tools AGENTS

## Mission

Add deterministic validation paths that fail loudly and help future sessions verify the real workflow quickly.

## Working rules

- Prefer local, reproducible checks over network-dependent validation.
- Validate the real implementation, not only the canonical docs.
- Keep validators narrow, explainable, and easy to rerun from the workspace root.
- Use synthetic fixtures when they prove a contract clearly and avoid brittle external dependencies.
- Record durable validation findings in `logs/`.

## Output standard

- A validator should make pass/fail obvious.
- A validator should name the command it ran or the artifact it expected.
- A validator should not silently skip missing prerequisites.
