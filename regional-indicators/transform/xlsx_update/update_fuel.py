"""
regional-indicators/transform/xlsx_update/update_fuel.py
-----------------------------------------------------------
Writes fetch_fuel.py's cached output (RACQ annual average regular
unleaded petrol prices) into the Exogenous sheet's Fuel section.

Built 2026-10-06.

SHEET LAYOUT (confirmed against the real 2025 working file)
    Fuel            | <years across this row>      <- section header row
    Bowen           |                              <- town row, column A only
    Average RULP Price (cents) | Bowen RULP | ...  <- the row written
    Brisbane
    Average RULP Price (cents) | ...
    ...
The Fuel section has its OWN year header row (the "Fuel" row itself);
it is not the sheet's row 1 and does not line up with it in every
column, so the year column is found from the Fuel row. If the new year
isn't there yet it is appended straight after the last year in that
row, and the year is written into the Fuel row as its header.

WHICH TOWNS: whichever blocks the sheet has. A block is any row in the
Fuel section whose NEXT row is labelled "Average RULP Price (cents)";
its column A text is looked up (case-insensitive, whitespace stripped)
in RACQ's list of locations. Nothing is hardcoded and towns.toml is
not involved -- RACQ's locations are its own list (Bowen is on it, and
it is not a study town). A block whose name RACQ doesn't publish is
reported and left alone.

Each write goes through audit.py's cell, series and ground-truth
historical checks. RACQ's table carries about eleven years of history,
so the sheet's existing figures are compared with the report; where an
earlier year differs the result line says so (RACQ does restate
figures -- each report "supersedes all previous reports") but only the
NEW year is written.

NOT WRITTEN (yet):
  - the derived rows below the town blocks (annual fuel cost and its
    year-on-year change) -- formulas that need extending by one column;
  - the Queensland / Australia benchmark rows the 2026 reference file
    adds at the bottom of the section. They don't exist in the 2025
    working file, and this pipeline doesn't create rows. The Queensland
    figure is already in the cache ("queensland_mean_of_locations").

Uses xlwings (real Excel via COM automation), NOT openpyxl -- see
base.py's docstring for why.

Tested 2026-10-06 against the real 2025 working file's cell contents
through tests/fake_xlwings_sheet.py (a stand-in for the Excel sheet
object): 8 written, all equal to the 2026 reference file. First real
Excel run still to be confirmed.

Usage:
    python update_fuel.py <path-to-Indicators_Data-Charts.xlsx> <cache/fuel dir> [--visible]
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
    """(town name, row to write) for every block in the Fuel section:
    a non-empty row directly followed by an INDICATOR_LABEL row."""
    blocks = []
    for index in range(header_row, len(labels) - 1):      # index is 0-based for row index+1
        name, below = labels[index], labels[index + 1]
        if name and name != INDICATOR_LABEL and below == INDICATOR_LABEL:
            blocks.append((name, index + 2))
    return blocks


def _find_year_column(sheet, header_row: int, year: int) -> int:
    """Column for `year` in the Fuel section's own header row, creating
    it straight after the last year in that row if it isn't there."""
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
    """Audit and write the report year's figure for every town block in
    the Fuel section. Separate from the Excel open/save handling so it
    can be exercised against any sheet-like object. Returns
    (result lines, written count, flagged count)."""
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
