"""
regional-indicators/transform/xlsx_update/update_fuel.py

Writes annual average petrol prices into the Exogenous sheet.

Input:  cache/fuel/racq_rulp_annual.json, produced by
        fetchers/fetch_fuel.py (RACQ Annual Fuel Price Report).
Target: the "Fuel" section of the Exogenous sheet.

Sheet layout:
    Fuel                         <years>     section heading and year header
    <location>                               heading, column A only
    Average RULP Price (cents)   <values>    the row written
    <location>
    Average RULP Price (cents)   <values>
    ...
The Fuel section has its own year header (the "Fuel" row), separate
from row 1 of the sheet. The year column is located in that row; a new
year is appended directly after the last year in the row and written
into the header.

Locations are taken from the sheet, not from configuration. Every row
in the section that is followed by an "Average RULP Price (cents)" row
is a location block, and its column A text is looked up in the report's
list of locations (case-insensitive, surrounding whitespace ignored).
A block whose name the report does not publish is reported and left
unchanged. To add a location, add a block to the sheet.

Only the report year is written. The report also carries about eleven
years of history, which is passed to the historical audit; earlier
years on the sheet that differ from the report are listed in the result
line and are not modified.

Out of scope: the derived rows below the location blocks (annual fuel
cost and its year-on-year change) and any state or national benchmark
rows.

The workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_fuel.py <workbook.xlsx> <cache/fuel dir> [--visible]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).parent))
from audit import audit_cell, audit_series, audit_historical_series, WriteAuditReport

SHEET_NAME      = "Exogenous"
SECTION_HEADER  = "Fuel"
INDICATOR_LABEL = "Average RULP Price (cents)"
FIRST_YEAR_COLUMN = 2          # column B
CACHE_FILE      = "racq_rulp_annual.json"


def _is_year(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and 1990 <= value <= 2100


def _column_a(sheet) -> list[str]:
    max_row = sheet.used_range.last_cell.row
    col_a = sheet.range((1, 1), (max_row, 1)).value
    if not isinstance(col_a, list):
        col_a = [col_a]
    return [str(v).strip() if v is not None else "" for v in col_a]


def _find_section_header_row(labels: list[str]) -> int:
    rows = [i + 1 for i, text in enumerate(labels) if text == SECTION_HEADER]
    if len(rows) != 1:
        raise ValueError(
            f"Expected exactly one row labelled '{SECTION_HEADER}' in column A of "
            f"sheet '{SHEET_NAME}', found {len(rows)} (rows {rows})."
        )
    return rows[0]


def _find_blocks(labels: list[str], header_row: int) -> list[tuple[str, int]]:
    """Return (location name, row to write) for every block in the Fuel
    section: a non-empty row directly followed by an INDICATOR_LABEL row.
    """
    blocks = []
    for index in range(header_row, len(labels) - 1):      # labels[index] is sheet row index + 1
        name, below = labels[index], labels[index + 1]
        if name and name != INDICATOR_LABEL and below == INDICATOR_LABEL:
            blocks.append((name, index + 2))
    return blocks


def _find_year_column(sheet, header_row: int, year: int) -> int:
    """Return the column for `year` in the Fuel section's header row,
    appending one directly after the last year if needed.
    """
    max_col = sheet.used_range.last_cell.column
    header = sheet.range((header_row, FIRST_YEAR_COLUMN), (header_row, max_col)).value
    if not isinstance(header, list):
        header = [header]

    last_year_col = None
    for offset, cell_year in enumerate(header):
        col = FIRST_YEAR_COLUMN + offset
        if cell_year == year:
            return col
        if _is_year(cell_year):
            last_year_col = col

    if last_year_col is None:
        raise ValueError(
            f"No year values found in the '{SECTION_HEADER}' header row (row {header_row}) "
            f"-- can't tell where to add {year}."
        )

    new_col = last_year_col + 1
    year_cell = sheet.cells(header_row, new_col)
    year_cell.value = year
    year_cell.number_format = "General"
    return new_col


def _read_existing_series(sheet, header_row: int, row: int, exclude_col: int) -> dict[int, float]:
    max_col = sheet.used_range.last_cell.column
    years = sheet.range((header_row, FIRST_YEAR_COLUMN), (header_row, max_col)).value
    values = sheet.range((row, FIRST_YEAR_COLUMN), (row, max_col)).value
    if not isinstance(years, list):
        years, values = [years], [values]
    series = {}
    for offset, (yr, val) in enumerate(zip(years, values)):
        if FIRST_YEAR_COLUMN + offset == exclude_col:
            continue
        if _is_year(yr) and isinstance(val, (int, float)) and not isinstance(val, bool):
            series[int(yr)] = val
    return series


def write_fuel(sheet, data: dict) -> tuple[list[str], int, int]:
    """Audit and write the report year's price for every location block.

    Separate from the Excel session handling so that it can be run
    against any object with the xlwings Sheet interface (see
    tests/fake_xlwings_sheet.py). Returns
    (result lines, written count, flagged count).
    """
    results: list[str] = []
    written_count = 0
    flagged_count = 0

    year = int(data["report_year"])
    by_name = {name.strip().lower(): series for name, series in data["locations"].items()}

    labels = _column_a(sheet)
    header_row = _find_section_header_row(labels)
    blocks = _find_blocks(labels, header_row)
    if not blocks:
        raise ValueError(
            f"Found the '{SECTION_HEADER}' section at row {header_row} but no town "
            f"blocks under it (a town row followed by '{INDICATOR_LABEL}')."
        )

    for town, row in blocks:
        series = by_name.get(town.lower())
        if series is None:
            results.append(
                f"{town}: SKIPPED — RACQ's report has no location called '{town}'. "
                f"Check the spelling against the report's Appendix 1."
            )
            flagged_count += 1
            continue
        if str(year) not in series:
            results.append(f"{town}: no {year} figure in RACQ's report ('nd'), skipped")
            flagged_count += 1
            continue

        value = series[str(year)]
        source_series = {int(y): v for y, v in series.items()}

        col = _find_year_column(sheet, header_row, year)
        cell = sheet.cells(row, col)
        existing = _read_existing_series(sheet, header_row, row, exclude_col=col)
        report = WriteAuditReport(
            audit_cell(cell, value),
            audit_series(existing, year, value),
            audit_historical_series(existing, source_series),
        )

        if report.safe_to_write:
            cell.value = value
            cell.number_format = "General"
            differing = sorted(
                y for y, v in existing.items()
                if y in source_series and abs(v - source_series[y]) > 0.05
            )
            note = ""
            if differing:
                detail = ", ".join(f"{y} sheet {existing[y]} vs report {source_series[y]}" for y in differing)
                note = f" (note: earlier years differ from this report, not changed — {detail})"
            results.append(f"{town}: WRITTEN {year} = {value} -> {cell.address}{note}")
            written_count += 1
        else:
            results.append(f"{town}: FLAGGED, not written — {report.summary_line()}")
            flagged_count += 1

    return results, written_count, flagged_count


def update_fuel(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    cache_path = cache_dir / CACHE_FILE
    if not cache_path.exists():
        raise FileNotFoundError(
            f"{cache_path} not found -- run fetch_fuel.py first "
            f"(python run_update.py --only fuel)."
        )
    with open(cache_path, encoding="utf-8") as f:
        data = json.load(f)

    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(str(xlsx_path))
        try:
            sheet = wb.sheets[SHEET_NAME]
            results, written_count, flagged_count = write_fuel(sheet, data)
            if written_count:
                wb.save()
        finally:
            wb.close()
    finally:
        app.quit()

    results.insert(0, f"Source: {data.get('source', 'RACQ Annual Fuel Price Report')}")
    results.append("")
    results.append(f"Summary: {written_count} written, {flagged_count} flagged for review.")
    return results


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--visible"]
    visible = "--visible" in sys.argv

    if len(args) != 2:
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/fuel dir> [--visible]")
        sys.exit(1)

    results = update_fuel(Path(args[0]), Path(args[1]), visible=visible)
    for line in results:
        print(line)
