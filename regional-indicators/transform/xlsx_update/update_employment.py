"""
regional-indicators/transform/xlsx_update/update_employment.py

Writes smoothed unemployment rates into the Employment sheet.

Input:  cache/unemployment/regions/*.json, one file per region,
        produced by fetch_salm_unemployment.py (SA2), fetch_qrsis_labour.py
        (Queensland LGAs and state) and fetch_nsw_labour.py (NSW).
Target: the Employment sheet.

Sheet layout:
  Row 1  calendar years from column C.
  Row 2  fiscal labels for the early columns, then the calendar year.
         When a new year column is created, row 2 is filled to match.
  Sections "LGA", "SA2" and "State", each introduced by its name in
  column A. Every data row has the region name in column A and
  "Smoothed Unemployment rate (%)" in column B.

A region name can occur in more than one section (Goondiwindi is both
an LGA and an SA2), so a region is searched for only within the bounds
of its own section.

The year-column, existing-series and write helpers are shared with the
Crime sheet, which has the same header layout, and are imported from
update_crime.

A series-level audit concern does not block the write: the value is
written and marked bold red for review. A cell-level block (a formula
or unexpected content) prevents the write.

The workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_employment.py <workbook.xlsx> <cache/unemployment dir> [--visible]
"""

from __future__ import annotations

import json
import sys
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).parent))
from audit import describe_write_outcome
from update_crime import _find_year_column, _write_crime_row

SHEET_NAME = "Employment"
SECTION_MARKERS = {"LGA": "LGA", "SA2": "SA2", "State": "State"}
SECTION_ORDER = ["LGA", "SA2", "State"]
YEAR_HEADER_ROW = 1
FISCAL_LABEL_ROW = 2


def _column_a(sheet) -> list:
    max_row = sheet.used_range.last_cell.row
    col_a = sheet.range((1, 1), (max_row, 1)).value
    return col_a if isinstance(col_a, list) else [col_a]


def _find_section_rows(sheet) -> dict:
    """Return {section name: row} for the section headings. Each must occur exactly once."""
    col_a = _column_a(sheet)
    found = {}
    for section, marker in SECTION_MARKERS.items():
        rows = [i + 1 for i, v in enumerate(col_a) if v == marker]
        if len(rows) != 1:
            raise ValueError(
                f"Expected exactly one '{marker}' section marker in column A of "
                f"'{sheet.name}', found rows {rows}."
            )
        found[section] = rows[0]
    return found


def _find_region_row(sheet, section_rows: dict, section: str, label: str) -> int:
    """Return the row of `region` within `section`, searching only that section's rows."""
    start = section_rows[section] + 1
    later = [r for r in section_rows.values() if r > section_rows[section]]
    end = (min(later) - 1) if later else sheet.used_range.last_cell.row
    col_a = sheet.range((start, 1), (end, 1)).value
    if not isinstance(col_a, list):
        col_a = [col_a]
    matches = [start + i for i, v in enumerate(col_a) if v == label]
    if not matches:
        raise ValueError(
            f"Could not find '{label}' in the {section} section (rows {start}-{end}) of "
            f"'{sheet.name}'. This function does not guess or create rows."
        )
    if len(matches) > 1:
        raise ValueError(f"AMBIGUOUS: '{label}' appears at rows {matches} in the {section} section.")
    return matches[0]


def _ensure_fiscal_label(sheet, year: int) -> None:
    """Fill row 2 for a newly created year column with the calendar year.

    Matches the existing columns. An existing label is never overwritten.
    """
    col = _find_year_column(sheet, year)
    cell = sheet.cells(FISCAL_LABEL_ROW, col)
    if cell.value is None:
        cell.value = year
        cell.number_format = "General"


def process_region(sheet, section_rows: dict, data: dict):
    """Write one region's latest complete year.

    Returns (result_line, counts_as_written, counts_as_flagged).
    """
    section = data["section"]
    label = data["region"]
    tag = f"{section}:{label}"
    values = data["indicators"]["unemployment"]["values"]
    if not values:
        return f"{tag}: no data, skipped", False, False

    latest = max(values, key=int)
    year, value = int(latest), values[latest]
    source_series = {int(y): v for y, v in values.items()}

    # The current-year column of this sheet is stored to one decimal place,
    # so the value written is rounded to match; otherwise the cell audit
    # would report rounding differences as conflicts. Earlier columns hold
    # unrounded values, so the source series used for the historical
    # comparison is left at full precision.
    value = float(Decimal(str(value)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))

    try:
        row = _find_region_row(sheet, section_rows, section, label)
    except ValueError as exc:
        return f"{tag}: SKIPPED (row-finding) — {exc}", False, True

    report, coord = _write_crime_row(sheet, row, year, value, source_series)
    if report.safe_to_write or report.should_write_with_flag:
        _ensure_fiscal_label(sheet, year)
    return describe_write_outcome(tag, report, coord, f"{year} = {value:.4g}")


def update_employment(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    region_files = sorted((cache_dir / "regions").glob("*.json"))
    if not region_files:
        raise FileNotFoundError(
            f"No region files in {cache_dir / 'regions'} -- run "
            f"`run_update.py --only salm_unemployment` first."
        )

    results, written, flagged = [], 0, 0
    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(str(xlsx_path))
        try:
            sheet = wb.sheets[SHEET_NAME]
            section_rows = _find_section_rows(sheet)
            any_written = False
            for path in region_files:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)
                line, is_written, is_flagged = process_region(sheet, section_rows, data)
                results.append(line)
                written += is_written
                flagged += is_flagged
                any_written |= is_written
            if any_written:
                wb.save()
        finally:
            wb.close()
    finally:
        app.quit()

    results += ["", f"Summary: {written} written, {flagged} flagged for review."]
    return results


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--visible"]
    if len(args) != 2:
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/unemployment dir> [--visible]")
        sys.exit(1)
    for line in update_employment(Path(args[0]), Path(args[1]), visible="--visible" in sys.argv):
        print(line)