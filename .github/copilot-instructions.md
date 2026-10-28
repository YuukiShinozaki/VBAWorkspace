# Workspace-wide instructions

> **Workspace purpose: Excel VBA development only.** Treat Excel workbooks, VBA source, and supporting Excel/VBA project documentation or tooling as the scope of this workspace. By default, interpret requests in this context; do not introduce or scaffold unrelated application types, languages, or project structures. If a request appears unrelated, clarify its connection to this workspace before proceeding.

- Follow the conventions of the project being changed. Keep changes focused, readable, and consistent with nearby code.
- Before editing, read the target project's README and inspect relevant source, configuration, and tests. Follow project-specific instructions where available; do not assume every project uses the same folder layout or Git policy.
- Do not create new folders or restructure an existing project unless the user requests it or the task requires it. Prefer existing project files and directories.
- After code changes, run the smallest relevant existing validation and report what was checked and what could not be checked.

## Excel and VBA development

- Treat the target Excel macro-enabled workbook (`.xlsm`/`.xlsb`) as the runtime artifact. Exported `.bas`, `.cls`, and `.frm` files under `Projects/` are source-control copies; editing them alone does not change code inside the workbook.
- For exported VBA source, follow the target project's README and existing VBA style and structure, and keep changes focused on the requested behavior.
- Do not import or otherwise apply exported VBA code to a workbook, change workbook contents, or run a macro unless the user explicitly requests that operation. When requested, confirm the exact workbook and component before acting; never substitute a similarly named workbook or module.
- Before referencing workbook objects by name, confirm the worksheet, defined name, or UserForm control exists. When investigating workbook behavior, consider relevant formulas, named ranges, data validation, and conditional formatting as well as VBA code.
- Preserve related UserForm `.frx` files with their `.frm` files. For code intended to be applied, validate in Excel/VBE as appropriate (including compile and runtime checks); distinguish static checks from workbook runtime validation.
- Keep each VBA logical statement within the VBA line-continuation limit: at most 24 continuation markers (` _`, up to 25 physical lines). Split longer statements into readable smaller statements or intermediate variables.

## Troubleshooting reports

- When investigating a concrete problem, check `Reports/Defects/` and, when relevant, `Reports/Maintenance/` for related cases. Verify any historical guidance against the current workbook, source, configuration, and environment before using it.
- For report creation, organization, updates, or investigation summaries, follow the repository rules in `.github/instructions/Report Writing Rules.instructions.md`. They define classification, templates, evidence, validation, and naming conventions.

These English-language instructions are the authoritative workspace guidance for Copilot.
