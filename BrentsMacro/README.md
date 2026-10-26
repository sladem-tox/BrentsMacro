# Brents Macro Python replacement

This package reproduces `CopyColumns_ByRowCheckThreshold` from the workbook macro.
It reads the active worksheet (or `--source-sheet`), checks each column in the
requested range, and copies qualifying values from the requested row range into
column D of the destination sheet. Existing columns A:C in that sheet are
preserved; contents from D onward are cleared first.

Install dependencies and run:

```powershell
uv sync
uv run brentsmacro Final_StudentData_2026_MacroVersion_Blank.xlsm `
  --source-sheet Untreated_Normalised --columns FQ:MD --rows 1:35 --threshold 3 --check-row 35 `
  --destination-sheet Filtered
```

For `.xlsm` input, the VBA project is preserved. By default, the named
destination worksheet is updated in the input workbook. Use `--output` to write
to a separate workbook instead.

To list the worksheets without changing the workbook:

```powershell
uv run brentsmacro Final_StudentData_2026_MacroVersion_Blank.xlsm --list-sheets
```
