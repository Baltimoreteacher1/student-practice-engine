# Enhancement Source Notes

Keep enhancement logic here when it operates on existing notebook bundles without changing the primary generator.

Current entrypoint:

- `premium_polish.py`
- `run_enhancement_inbox.py`
- `polish_notebook_pptx.py`

The enhancement runner should:

- inspect an existing notebook bundle
- score premium quality categories
- repair obvious plan-level quality issues safely
- write a render-ready `notebook_plan.json`
- leave the final PPTX rendering to `../notebook_engine.py`
- apply a conservative PPTX polish pass when the inbox contains raw notebook `.pptx` files instead of bundle folders
