# StudioTax Schedule 3 Import Helper

This repository contains a small PowerShell helper for entering capital gains rows from a CSV into the StudioTax Schedule 3 attachment dialog.

The script reads a local CSV file and types values into the currently focused StudioTax grid one cell at a time. It does not include any tax data.

## CSV Columns

The CSV must include these columns:

```text
Number
Name of fund/corp.
Year of acquisition
Proceeds of disposition
Adjusted cost base
Outlays and expenses
```

Extra CSV columns are ignored. `Gain or loss` is assumed to be calculated by StudioTax, so the script does not type into that column.

## Files

- `Import-StudioTaxSchedule3.ps1` - the automation helper.
- `.gitignore` - excludes local CSV, TSV, Excel, and log files.

## Safety Notes

- Test on a copy of the return if possible.
- Start with `-MaxRows 2` until the cursor movement is confirmed.
- Review StudioTax totals before clicking `OK`.
- If the script is typing into the wrong place, click the PowerShell window and press `Ctrl+C`.

## Preview the Data

From PowerShell:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Import-StudioTaxSchedule3.ps1 .\schedule3.csv -Preview
```

This does not type into StudioTax. It prints the row count, expected totals, and the first two rows that would be entered.

## Recommended Two-Row Test

1. Open StudioTax to the Schedule 3 attachment dialog.
2. Clear any existing text from the first grid row.
3. Click the first yellow `Number` cell.
4. Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Import-StudioTaxSchedule3.ps1 .\schedule3.csv -MaxRows 2 -CountdownSeconds 5 -DelayMs 120
```

During the countdown, make sure the cursor is still in the first yellow `Number` cell.

## Full Run

Only use this after the two-row test works:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Import-StudioTaxSchedule3.ps1 .\schedule3.csv -CountdownSeconds 5 -DelayMs 120
```

## Tuning Options

`.\schedule3.csv`

Positional path to the local CSV file to import. `-CsvPath ".\schedule3.csv"` also works.

`-MaxRows 2`

Limits the run to the first two CSV rows. Use this for testing.

`-Preview`

Prints the row count, expected totals, and first two rows without typing into StudioTax.

`-DelayMs 250`

Slows down typing and tabbing. Increase this if StudioTax misses characters or navigation.

`-ExtraTabsAfterRow 0`

Changes how many extra tabs are sent after each row. The default is `1`, intended to skip the calculated `Gain or loss` column. If the cursor lands too far right or too far left on the next row, try `0` or `2`.
