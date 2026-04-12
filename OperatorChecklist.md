# Operator Checklist

Use this as the short day-of-use process.

## Update The Workbook

1. Export the latest Google Forms responses as CSV.
2. Replace:
   [ExampleGoogleFormOutput.csv](./ExampleGoogleFormOutput.csv)
3. Open:
   [UKMARS Contest Registration Processing.xlsx](./UKMARS%20Contest%20Registration%20Processing.xlsx)
4. Open the new CSV in Excel.
5. Copy the full used range including headers.
6. Paste into `GoogleFormResponsesValueCopy` starting at `A1`.

## Recalculate And Save

1. Set Excel to `Formulas -> Calculation Options -> Automatic`
2. Press:

```text
Ctrl+Alt+F9
```

3. Save the workbook.

## Validate

Check these before exporting anything:

1. `RobotLookup Clean`
   Confirm rows are populated and there are no `ZZZ` placeholders.
2. `Value Copied Summary`
   Confirm Junior and Senior totals look sensible.
3. Competition sheets
   Spot check a few populated sheets and confirm `Robot Name` and `Builder Name` are present.
4. Data sheets
   Check:
   - `Drag Race Entrants`
   - `RATS DATA`
   - `Line Following Entrants`
   - `Pursuit Entrants`

## What Feeds What

- The 14 competition sheets feed:
  - PDF export
  - CSV export
  - `competition-excels/Competition Sheets.xlsx`
- `Drag Race Entrants` feeds:
  - drag race external system
  - `dataextracts/Drag Race Entrants.xlsx`
- `RATS DATA` feeds:
  - RATS / wall follower / maze solver handoff
  - `dataextracts/RATS DATA.xlsx`
- `Line Following Entrants` feeds:
  - line following entrant handoff
  - `dataextracts/Line Following Entrants.xlsx`
- `Pursuit Entrants` feeds:
  - pursuit entrant handoff
  - `dataextracts/Pursuit Entrants.xlsx`

## Close The Workbook

1. Save again if needed.
2. Close Excel completely.

## Export PDFs

Run in Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\export_competition_pdfs.ps1 -WorkbookPath ".\UKMARS Contest Registration Processing.xlsx" -OutputFolder ".\competition-pdfs"
```

## Export CSVs

Run in Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\export_competition_csvs.ps1 -WorkbookPath ".\UKMARS Contest Registration Processing.xlsx" -OutputFolder ".\competition-csvs"
```

If the CSV export fails once with an Excel automation error, run the same command again.

## Export Value-Only Excel Files

Run in the Linux / WSL shell:

```bash
./.venv-linux/bin/python export_value_workbooks.py
```

This creates:

- `competition-excels/Competition Sheets.xlsx`
- `dataextracts/Drag Race Entrants.xlsx`
- `dataextracts/RATS DATA.xlsx`
- `dataextracts/Line Following Entrants.xlsx`
- `dataextracts/Pursuit Entrants.xlsx`

## Output Folders

- PDFs: `competition-pdfs`
- CSVs: `competition-csvs`
- Combined Excel workbook: `competition-excels`
- Data extract workbooks: `dataextracts`

## Full Runbook

For the full workflow and troubleshooting notes, see:

- [CompetitionWorkflow.md](./CompetitionWorkflow.md)
