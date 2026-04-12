# AGENTS.md

## Repo Purpose

This repository contains an Excel-based competition management workflow for UKMARS contest registration.

The central artifact is the workbook:

- `UKMARS Contest Registration Processing.xlsx`

The raw registration input is:

- `ExampleGoogleFormOutput.csv`

The normal workflow is:

1. Replace the CSV with the latest Google Forms export.
2. Paste the CSV contents into workbook sheet `GoogleFormResponsesValueCopy`.
3. Open the workbook in desktop Excel.
4. Recalculate and save the workbook.
5. Export downstream PDFs, CSVs, and value-only Excel files.

## Start Here

If workbook logic or export behavior is relevant, read these files first:

1. `CompetitionWorkflow.md`
2. `OperatorChecklist.md`
3. `export_competition_pdfs.ps1`
4. `export_competition_csvs.ps1`
5. `export_value_workbooks.py`

For Python environment expectations, read:

- `PythonEnvironment.md`

## Important Files

- Workbook: `UKMARS Contest Registration Processing.xlsx`
- Raw CSV input: `ExampleGoogleFormOutput.csv`
- PDF export: `export_competition_pdfs.ps1`
- CSV export: `export_competition_csvs.ps1`
- Value-only workbook export: `export_value_workbooks.py`
- Older repair utility: `fix_mouse_workbook.py`

## Workbook Data Flow

The workbook currently flows through these sheets:

1. `GoogleFormResponsesValueCopy`
   Raw pasted CSV data.
2. `Denormalised Responses`
   Expands each form response into robot rows.
3. `RobotLookup Clean`
   Main cleaned lookup table for downstream sheets.
4. `Value Copied Summary`
   Summary totals sheet.
5. 14 competition sheets
   Operational sheets for printing and exports.
6. Data extract sheets
   - `Drag Race Entrants`
   - `RATS DATA`
   - `Line Following Entrants`
   - `Pursuit Entrants`

## Environment Rules

- Use WSL/Linux for Python work in this repo.
- Use the Linux virtual environment at `./.venv-linux`.
- Do not use the Windows venv from WSL.
- When running Python, prefer:
  - `./.venv-linux/bin/python`
- PDF and CSV exports must be run from Windows PowerShell because they use Excel COM automation.

## Standard Commands

Run value-only workbook exports from WSL:

```bash
./.venv-linux/bin/python export_value_workbooks.py
```

Run PDF export from Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\export_competition_pdfs.ps1 -WorkbookPath ".\UKMARS Contest Registration Processing.xlsx" -OutputFolder ".\competition-pdfs"
```

Run CSV export from Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\export_competition_csvs.ps1 -WorkbookPath ".\UKMARS Contest Registration Processing.xlsx" -OutputFolder ".\competition-csvs"
```

## Output Locations

- PDFs: `competition-pdfs`
- Competition CSVs: `competition-csvs`
- Combined competition workbook: `competition-excels/Competition Sheets.xlsx`
- Data extract workbooks: `dataextracts`

Older output snapshots may also exist in sibling `_old` folders.

## Critical Workflow Constraints

- Excel must recalculate and save the workbook before Python value-only exports will work.
- If cached formula values are missing, `export_value_workbooks.py` will fail by design.
- Competition PDF and CSV exports depend on populated competition sheets and desktop Excel automation.
- If the CSV export hits a transient Excel automation error once, rerunning it may succeed.

## Validation Expectations

Before generating outputs, verify:

- `RobotLookup Clean` has populated rows and no `ZZZ` placeholders.
- `Value Copied Summary` totals look sensible.
- A few populated competition sheets contain titles, headers, robot names, and builder names.
- The four data extract sheets are populated and look structurally correct.

## Editing Guidance For Future Agents

- Treat the workbook and export scripts as the primary focus of this repo.
- Do not disturb unrelated files unless the task explicitly requires it.
- Prefer updating documentation and scripts over making speculative workbook changes.
- If workbook formula or sheet-structure edits are required, be explicit about the risk because the workbook is the system of record and is not easy to review as text.
- Keep documentation references rooted in this repository. Do not hard-code external machine-specific paths.
- When describing commands, clearly separate WSL/Linux commands from Windows PowerShell commands.

## Assumptions To Preserve

- The workbook and CSV live at the repo root.
- The Linux venv is `.venv-linux`.
- Output folders are rooted at this repository.
- The current operational docs are `CompetitionWorkflow.md` and `OperatorChecklist.md`.
