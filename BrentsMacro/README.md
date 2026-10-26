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
  --columns BR:EG --rows 1:35 --threshold 3 --check-row 35 `
  --destination-sheet Filtered --output Final_StudentData_2026_filtered.xlsm
```

For `.xlsm` input, the VBA project is preserved. If `--output` is omitted, the
script writes a sibling file with `_filtered` appended to its name.