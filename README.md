# UKMARS Excel Contest Registration

This repository contains an Excel-based competition management system for UKMARS contest registration.

The main workbook, `UKMARS Contest Registration Processing.xlsx`, takes registration data exported from Google Forms as `ExampleGoogleFormOutput.csv`, transforms it through a series of formula-driven sheets, and produces the competition sheets and data extracts used for event operations.

## What This Repo Is For

This repo is used to:

- import raw registration data from Google Forms
- recalculate and validate the master competition workbook
- generate printable competition PDFs
- generate competition CSV exports
- generate value-only Excel outputs for operational handoff sheets and external systems

The workbook currently drives:

- 14 competition sheets for scoring and administration
- `Drag Race Entrants`
- `RATS DATA`
- `Line Following Entrants`
- `Pursuit Entrants`

## Main Files

- Workbook: `UKMARS Contest Registration Processing.xlsx`
- Raw CSV input: `ExampleGoogleFormOutput.csv`
- PDF export script: `export_competition_pdfs.ps1`
- CSV export script: `export_competition_csvs.ps1`
- Value-only workbook export script: `export_value_workbooks.py`
- Full workflow notes: `CompetitionWorkflow.md`
- Short operator runbook: `OperatorChecklist.md`
- Python environment notes: `PythonEnvironment.md`

## What You Need To View And Run It

To work with the workbook itself, you need:

- Windows with desktop Microsoft Excel
- the workbook and CSV at the repo root

To run the PowerShell export scripts, you need:

- Windows PowerShell
- desktop Excel installed locally

To run the Python export, you need:

- WSL / Linux shell access on the same machine
- the Linux virtual environment at `./.venv-linux`
- Python packages installed in that venv, including `openpyxl`

## Tested Environment

This workflow has only been tested on:

- Windows 10
- Microsoft Excel for Microsoft 365 MSO
- Version `2603`
- Build `16.0.19822.20086`
- `32-bit`

No claim is made here that the workbook automation has been verified on other Excel versions, other Office channels, macOS Excel, or LibreOffice.

## Workflow Summary

1. Export the latest registration responses from Google Forms as CSV.
2. Replace `ExampleGoogleFormOutput.csv`.
3. Open `UKMARS Contest Registration Processing.xlsx` in Excel.
4. Paste the CSV contents into `GoogleFormResponsesValueCopy` starting at `A1`.
5. Ensure Excel calculation mode is `Automatic`.
6. Recalculate with `Ctrl+Alt+F9`.
7. Save the workbook.
8. Validate `RobotLookup Clean`, `Value Copied Summary`, the populated competition sheets, and the four data extract sheets.
9. Run the required exports.

Important: the Python export depends on Excel having recalculated and saved cached formula results first.

## Export Commands

Run these from the repository root.

### Export Competition PDFs

Run in Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\export_competition_pdfs.ps1 -WorkbookPath ".\UKMARS Contest Registration Processing.xlsx" -OutputFolder ".\competition-pdfs"
```

Output folder:

- `competition-pdfs`

### Export Competition CSVs

Run in Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\export_competition_csvs.ps1 -WorkbookPath ".\UKMARS Contest Registration Processing.xlsx" -OutputFolder ".\competition-csvs"
```

Output folder:

- `competition-csvs`

If the CSV export fails once with an Excel automation error, run the same command again.

### Export Value-Only Excel Workbooks

Run in WSL / Linux:

```bash
./.venv-linux/bin/python export_value_workbooks.py
```

Outputs:

- `competition-excels/Competition Sheets.xlsx`
- `dataextracts/Drag Race Entrants.xlsx`
- `dataextracts/RATS DATA.xlsx`
- `dataextracts/Line Following Entrants.xlsx`
- `dataextracts/Pursuit Entrants.xlsx`

## Notes On Environments

- Use WSL for Python work in this repo.
- Use Windows PowerShell for the Excel COM export scripts.
- Do not try to run the Windows virtual environment from WSL.

## More Detail

For the full process and validation checks, see:

- `CompetitionWorkflow.md`
- `OperatorChecklist.md`
- `PythonEnvironment.md`
