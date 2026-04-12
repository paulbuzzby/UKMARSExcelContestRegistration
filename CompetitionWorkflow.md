# Competition Spreadsheet Workflow

This file documents the current end-to-end process for updating the competition workbook from a new Google Forms CSV export, validating the workbook, and generating the downstream PDF, CSV, and value-only Excel outputs.

## Core Files

- Workbook: [UKMARS Contest Registration Processing.xlsx](./UKMARS%20Contest%20Registration%20Processing.xlsx)
- CSV input: [ExampleGoogleFormOutput.csv](./ExampleGoogleFormOutput.csv)
- PDF export script: [export_competition_pdfs.ps1](./export_competition_pdfs.ps1)
- CSV export script: [export_competition_csvs.ps1](./export_competition_csvs.ps1)
- Value-only workbook export script: [export_value_workbooks.py](./export_value_workbooks.py)
- Python environment notes: [PythonEnvironment.md](./PythonEnvironment.md)

## Workbook Structure

The workbook currently works like this:

1. `GoogleFormResponsesValueCopy`
   This is the raw imported form data. It should mirror the CSV exactly.

2. `Denormalised Responses`
   This expands each form response into up to 5 robot rows.

3. `RobotLookup Clean`
   This is the main clean lookup sheet.
   It removes the old `ZZZ` filler-row approach and is the main source for downstream reporting.

4. `Value Copied Summary`
   This totals directly from `RobotLookup Clean`.

5. Competition sheets
   These are the 14 printable competition sheets used for PDF and CSV export.

6. Data / entrant sheets
   These are the additional dynamic sheets used for external systems or admin extracts:
   - `Drag Race Entrants`
   - `RATS DATA`
   - `Line Following Entrants`
   - `Pursuit Entrants`

## Sheet Usage And External Outputs

The workbook now has three main output groups.

### Internal Workbook Summary / Validation

These are used inside the main workbook and should be checked after each CSV refresh:

- `RobotLookup Clean`
  Main clean lookup sheet used as the source for downstream dynamic sheets.
- `Value Copied Summary`
  Summary totals sheet driven from `RobotLookup Clean`.

### Competition Administration And Print Outputs

These 14 sheets are the operational competition sheets used for manual scoring, PDF export, CSV export, and the combined value-only competition workbook:

- `Junior Half Size Line Follower`
- `Senior Half Size Line Follower`
- `Junior Half Size Wall Follower`
- `Senior Half Size Wall Follower`
- `Junior Wall Follower`
- `Senior Wall Follower`
- `Junior Maze Solver`
- `Senior Maze Solver`
- `Junior Half Size Maze Solver`
- `Senior Half Size Maze Solver`
- `Junior Drag Race`
- `Senior Drag Race`
- `Junior Pursuit Race`
- `Senior Pursuit Race`

These feed:

- PDF export via [export_competition_pdfs.ps1](./export_competition_pdfs.ps1)
- CSV export via [export_competition_csvs.ps1](./export_competition_csvs.ps1)
- combined value-only workbook:
  `competition-excels/Competition Sheets.xlsx`

### External Systems / Data Extract Sheets

These sheets are intended for external systems or structured handoff files:

- `Drag Race Entrants`
  Feed for the external drag race system.
  Contains `Robot Name`, `Builder Name`, `Email`, `Class`.
  Also exported as a value-only workbook:
  `dataextracts/Drag Race Entrants.xlsx`

- `RATS DATA`
  Feed for the RATS / wall follower / maze solver workflow.
  Includes entries in `Wall Follower` or `Maze Solver` only, excluding half-size events.
  Contains `Robot`, `Builder`, `Email`, `Class`, `Entries`, `Status`.
  Also exported as a value-only workbook:
  `dataextracts/RATS DATA.xlsx`

- `Line Following Entrants`
  Feed for line following entrant extraction.
  Includes both standard and half-size line following entries.
  Ordered by `Class` then `Builder Name`.
  Contains `Robot Name`, `Builder Name`, `Email`, `Class`.
  Also exported as a value-only workbook:
  `dataextracts/Line Following Entrants.xlsx`

- `Pursuit Entrants`
  Feed for pursuit entrant extraction.
  Ordered by `Class` then `Builder Name`.
  Contains `Robot Name`, `Builder Name`, `Email`, `Class`.
  Also exported as a value-only workbook:
  `dataextracts/Pursuit Entrants.xlsx`

## Capacity / Design Assumptions

- The workbook is designed to support at least `200` Google Form responses.
- Each form response can contain up to `5` robots.
- Formula-driven sheets depend on Excel recalculating and saving the workbook before non-Excel tools can read the results reliably.

## Step 1: Replace The Source CSV

Export the latest responses from Google Forms as CSV.

Replace:

- [ExampleGoogleFormOutput.csv](./ExampleGoogleFormOutput.csv)

Important:

- keep the same column order as the current CSV
- keep the header row
- do not add extra blank columns

## Step 2: Update `GoogleFormResponsesValueCopy`

Open the workbook in Excel and update `GoogleFormResponsesValueCopy` so it exactly matches the new CSV.

Recommended process:

1. Open the new CSV in Excel.
2. Select the full used range including headers.
3. Copy it.
4. Go to `GoogleFormResponsesValueCopy` in the main workbook.
5. Select cell `A1`.
6. Paste as values.

Rules:

- headers stay in row `1`
- data starts in row `2`
- column order must not change

## Step 3: Recalculate In Excel

After updating the CSV data:

1. Confirm Excel calculation mode is `Automatic`
   Excel path: `Formulas -> Calculation Options -> Automatic`
2. Press:

```text
Ctrl+Alt+F9
```

3. Save the workbook.

This recalculates:

- `Denormalised Responses`
- `RobotLookup Clean`
- `Value Copied Summary`
- the 14 competition sheets
- `Drag Race Entrants`
- `RATS DATA`
- `Line Following Entrants`
- `Pursuit Entrants`

## Step 4: Validate The Workbook

Validate the workbook before running any exports.

### `RobotLookup Clean`

Check:

- rows are populated
- `Robot`, `Builder`, `Class`, and `Email` look correct
- there are no `ZZZ` placeholders
- event flags look sensible

### `Value Copied Summary`

Check:

- Junior totals look reasonable
- Senior totals look reasonable
- totals reflect the current registrations

### Competition Sheets

Spot check a few populated sheets such as:

- `Junior Half Size Line Follower`
- `Senior Maze Solver`
- `Senior Pursuit Race`

Check:

- `A1` contains the title
- row `4` contains headers
- `A5` downward contains robot names
- `B5` downward contains builder names
- result columns are blank and ready for manual scoring

### Data / Entrant Sheets

Check:

- `Drag Race Entrants`
- `RATS DATA`
- `Line Following Entrants`
- `Pursuit Entrants`

Expected behavior:

- rows are populated only from real matching entries
- sheets update from `RobotLookup Clean`
- ordering is sheet-specific and formula-driven

## Step 5: Save And Close The Workbook

Before running any automation:

1. Save the workbook.
2. Close Excel completely.

This matters because:

- the PowerShell export scripts open the workbook through Excel automation
- the Python value-only workbook export relies on Excel having saved cached formula results

## Step 6: Export PDFs

Run from Windows PowerShell in the workspace root:

```powershell
powershell -ExecutionPolicy Bypass -File .\export_competition_pdfs.ps1 -WorkbookPath ".\UKMARS Contest Registration Processing.xlsx" -OutputFolder ".\competition-pdfs"
```

What this exports:

- one PDF per populated competition sheet
- output folder is created automatically if missing
- files are named from the sheet names

Current competition sheets included:

- `Junior Half Size Line Follower`
- `Senior Half Size Line Follower`
- `Junior Half Size Wall Follower`
- `Senior Half Size Wall Follower`
- `Junior Wall Follower`
- `Senior Wall Follower`
- `Junior Maze Solver`
- `Senior Maze Solver`
- `Junior Half Size Maze Solver`
- `Senior Half Size Maze Solver`
- `Junior Drag Race`
- `Senior Drag Race`
- `Junior Pursuit Race`
- `Senior Pursuit Race`

PDF layout rules:

- start print area at `A1`
- end at the last entrant row in column `A`
- end at the last visible header column in row `4`
- `A4`
- `Landscape`
- fit `1` page wide
- unlimited pages tall

## Step 7: Export Competition CSVs

Run from Windows PowerShell in the workspace root:

```powershell
powershell -ExecutionPolicy Bypass -File .\export_competition_csvs.ps1 -WorkbookPath ".\UKMARS Contest Registration Processing.xlsx" -OutputFolder ".\competition-csvs"
```

What this exports:

- one CSV per populated competition sheet
- includes the header row from `A4:B4`
- exports only:
  - `Robot Name`
  - `Builder Name`
- output folder is created automatically if missing

Notes:

- Excel COM can occasionally be transiently busy on the first run
- the script has retry logic, but if a run fails and a second run works, that is usually an Excel automation timing issue rather than a workbook data issue

## Step 8: Export Value-Only Excel Workbooks

This step creates `.xlsx` outputs that contain values only, not formulas.

Important:

- open the workbook in Excel first
- run `Ctrl+Alt+F9`
- save
- close the workbook

Then run from the Linux / WSL shell:

```bash
./.venv-linux/bin/python export_value_workbooks.py
```

What this creates:

1. A combined workbook:
   - `competition-excels/Competition Sheets.xlsx`
   - contains the 14 competition sheets as value-only copies

2. Single-sheet data extract workbooks in `dataextracts`:
   - `Drag Race Entrants.xlsx`
   - `RATS DATA.xlsx`
   - `Line Following Entrants.xlsx`
   - `Pursuit Entrants.xlsx`

The Python script will stop with a clear error if the source workbook has not been recalculated and saved in Excel first.

## Output Folders

Current automation outputs go to:

- PDFs: `competition-pdfs`
- CSVs: `competition-csvs`
- combined competition workbook: `competition-excels`
- single-sheet value-only extracts: `dataextracts`

## Troubleshooting

### `RobotLookup Clean` is blank or stale

In Excel:

1. set calculation mode to `Automatic`
2. press `Ctrl+Alt+F9`
3. save again

### `Value Copied Summary` looks wrong

Check `RobotLookup Clean` first. `Value Copied Summary` depends on it.

### Competition sheets are blank

Check:

- `RobotLookup Clean` is populated
- the relevant event actually has entries
- the workbook was recalculated after the CSV paste

### PDF export fails

Check:

- the workbook is closed
- desktop Excel is installed
- the command is run from Windows PowerShell

### CSV export fails once but works on a second run

That is usually a transient Excel COM timing problem. Re-run the same command.

### Value-only workbook export fails with a cached-results error

That means the workbook formulas were not saved with up-to-date values.

Fix:

1. open workbook in Excel
2. press `Ctrl+Alt+F9`
3. save
4. close workbook
5. rerun `export_value_workbooks.py`

## Quick Reference

1. Replace CSV
2. Paste CSV into `GoogleFormResponsesValueCopy`
3. `Ctrl+Alt+F9`
4. Save workbook
5. Validate `RobotLookup Clean`
6. Validate `Value Copied Summary`
7. Validate competition sheets and data sheets
8. Save and close workbook
9. Run PDF export if needed
10. Run CSV export if needed
11. Run value-only workbook export if needed
