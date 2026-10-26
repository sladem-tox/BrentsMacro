"""Copy columns whose check-row value is at or below a threshold."""

from __future__ import annotations

import argparse
import math
import re
from pathlib import Path
from typing import Optional, Tuple, Union

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


MAX_EXCEL_ROW = 1_048_576


def _parse_column_range(spec: str) -> Tuple[int, int]:
    match = re.fullmatch(r"\s*([A-Za-z]+)\s*:\s*([A-Za-z]+)\s*", spec)
    if not match:
        raise ValueError("column range must look like BR:EG or B:D")
    try:
        start = column_index_from_string(match.group(1).upper())
        end = column_index_from_string(match.group(2).upper())
    except ValueError as exc:
        raise ValueError("column range contains an invalid column") from exc
    if end < start:
        raise ValueError("column range must end at or after its start")
    return start, end


def _parse_row_range(spec: str, max_row: int = MAX_EXCEL_ROW) -> Tuple[int, int]:
    match = re.fullmatch(r"\s*(\d+)\s*:\s*(\d+)\s*", spec)
    if not match:
        raise ValueError("row range must look like 1:35")
    start, end = int(match.group(1)), int(match.group(2))
    if start < 1 or end < start or end > max_row:
        raise ValueError(f"row range must be between 1 and {max_row}")
    return start, end


def _numeric_value(value: object) -> Optional[float]:
    """Return the VBA-style numeric conversion, or None for non-numeric cells."""
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        converted = float(value)
    elif isinstance(value, str) and value.strip():
        try:
            converted = float(value)
        except ValueError:
            return None
    else:
        return None
    return converted if math.isfinite(converted) else None


def copy_columns_by_row_check_threshold(
    workbook_path: Union[str, Path],
    *,
    source_sheet: Optional[str] = None,
    columns: str,
    rows: str,
    threshold: float,
    check_row: int,
    destination_sheet: str = "Filtered",
    output_path: Optional[Union[str, Path]] = None,
) -> Tuple[Path, int]:
    """Copy qualifying source columns to destination column D, preserving values only."""
    workbook_path = Path(workbook_path)
    if not workbook_path.is_file():
        raise FileNotFoundError(workbook_path)
    if check_row < 1:
        raise ValueError("check_row must be at least 1")
    if not destination_sheet or len(destination_sheet) > 31:
        raise ValueError("destination_sheet must contain 1 to 31 characters")
    if any(char in destination_sheet for char in '[]:*?/\\'):
        raise ValueError("destination_sheet contains an invalid Excel sheet-name character")

    keep_vba = workbook_path.suffix.lower() == ".xlsm"
    workbook = load_workbook(workbook_path, keep_vba=keep_vba)
    source = workbook[source_sheet] if source_sheet else workbook.active
    start_col, end_col = _parse_column_range(columns)
    start_row, end_row = _parse_row_range(rows)
    if check_row > MAX_EXCEL_ROW:
        raise ValueError(f"check_row must be between 1 and {MAX_EXCEL_ROW}")

    if destination_sheet in workbook.sheetnames:
        destination = workbook[destination_sheet]
        for row in destination.iter_rows(
            min_row=1, max_row=destination.max_row, min_col=4, max_col=destination.max_column
        ):
            for cell in row:
                cell.value = None
    else:
        destination = workbook.create_sheet(destination_sheet)

    copied = 0
    for source_col in range(start_col, end_col + 1):
        numeric_value = _numeric_value(source.cell(check_row, source_col).value)
        if numeric_value is not None and numeric_value <= threshold:
            for offset, row_number in enumerate(range(start_row, end_row + 1), start=1):
                destination.cell(offset, 4 + copied).value = source.cell(
                    row_number, source_col
                ).value
            copied += 1

    output = Path(output_path) if output_path else workbook_path.with_name(
        f"{workbook_path.stem}_filtered{workbook_path.suffix}"
    )
    workbook.save(output)
    return output, copied


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Copy columns whose check-row value is at or below a threshold."
    )
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--columns", required=True, help="Column range, e.g. BR:EG")
    parser.add_argument("--rows", required=True, help="Row range, e.g. 1:35")
    parser.add_argument("--threshold", required=True, type=float)
    parser.add_argument("--check-row", required=True, type=int)
    parser.add_argument("--destination-sheet", default="Filtered")
    parser.add_argument("--source-sheet")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    output, copied = copy_columns_by_row_check_threshold(
        args.workbook,
        source_sheet=args.source_sheet,
        columns=args.columns,
        rows=args.rows,
        threshold=args.threshold,
        check_row=args.check_row,
        destination_sheet=args.destination_sheet,
        output_path=args.output,
    )
    print(f"Copy complete: {copied} columns copied to '{args.destination_sheet}' in {output}")
