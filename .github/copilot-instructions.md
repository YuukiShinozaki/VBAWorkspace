# Workspace-wide instructions

> **Workspace purpose: Excel VBA development only.** Treat Excel workbooks, VBA source, and supporting Excel/VBA project documentation or tooling as the scope of this workspace. By default, interpret requests in this context; do not introduce or scaffold unrelated application types, languages, or project structures. If a request appears unrelated, clarify its connection to this workspace before proceeding.

- Follow the conventions of the project being changed. Keep changes focused, readable, and consistent with nearby code.
- Before editing, read the target project's README and inspect relevant source, configuration, and tests. Follow project-specific instructions where available; do not assume every project uses the same folder layout or Git policy.
- Do not create new folders or restructure an existing project unless the user requests it or the task requires it. Prefer existing project files and directories.
- After code changes, run the smallest relevant existing validation and report what was checked and what could not be checked.

## Excel and VBA development

### Mandatory operator approval before Excel operations

- Before starting Excel, opening a workbook, running a macro, or applying changes to a live workbook, always ask the operator for explicit approval in clear Japanese and wait for an affirmative reply. This confirmation is required even when the original request asks for the operation; a general request to investigate, fix, or validate is not permission.
- Make it unmistakable that this is a permission request, not a notice that work is starting. State the exact workbook path, the planned operation, and, when applicable, the module/macro and whether workbook contents will change. For example: 「【許可の確認】Excelを起動して、ブック『<フルパス>』を開いてもよいですか？ 読み取りのみで、マクロの実行や変更・保存は行いません。許可をいただくまで操作しません。」 For execution, ask separately: 「【実行許可の確認】ブック『<フルパス>』のマクロ『<モジュール名.マクロ名>』を実行してもよいですか？ 想定される変更：<変更内容>。許可をいただくまで実行しません。」
- This rule also applies to tools, scripts, tasks, COM automation, and delegated agents that may automatically launch Excel or open a workbook, including read-only inspection tools. Obtain approval before invoking them, not after Excel has opened. If the behavior is uncertain, inspect offline files/documentation first or ask for approval before proceeding.
- Approval is limited to the stated workbook and operations. Permission to open/read is not permission to run macros, import code, modify, or save. Ask again before expanding the approved scope. If approval is absent, unclear, or denied, do not proceed; limit work to offline source/documentation and explain any validation limits.

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
