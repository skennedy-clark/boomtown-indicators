"""
regional-indicators/transform/xlsx_update/update_education.py

Writes school enrolments and teaching staff into the Exogenous sheet.

Input:  cache/schools/<slug>_schools.json, produced by
        fetchers/fetch_schools.py (ACARA School Profile).
Target: the "Education" section of the Exogenous sheet.

Sheet layout:
    Education                             <years>   section heading and year header
    <town>                                          heading, column A only
    Full Time Equivalent Enrolments       <values>
    Full Time Equivalent Teaching Staff   <values>
    <town>
    ...
The Education section has its own year header (the "Education" row),
separate from row 1 of the sheet. The year column is located in that
row; a new year is appended directly after the last year in the row
and written into the header.

A town's block is located by its heading within the Education section
only, and the two rows below it are verified against the expected
labels. Blocks for the same towns in other sections of the sheet are
never matched. Towns that are fetched but have no block are listed once
at the end of the run.

Only the latest year is written. The cache carries the full series
from 2008, which is passed to the historical audit; earlier years on
the sheet that differ from the source are summarised in the result line
and are not modified.

The section helpers (_column_a, _find_year_column,
_read_existing_series) mirror those in update_fuel.py.

The workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_education.py <workbook.xlsx> <cache/schools dir> [--visible]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).parent))
from audit import audit_cell, audit_series, audit_historical_series, WriteAuditReport

SHEET_NAME      = "Exogenous"
SECTION_HEADER  = "Education"
FIRST_YEAR_COLUMN = 2          # column B
CACHE_GLOB      = "*_schools.json"

# (cache key, row label, row offset below the town heading)
INDICATORS = (
    ("fte_enrolments",     "Full Time Equivalent Enrolments",     1),
    ("fte_teaching_staff", "Full Time Equivalent Teaching Staff", 2),
)
NEXT_SECTION_HEADERS = ("Fuel", "Rainfall", "Business", "Crime",
                        "Employment", "Housing", "Income", "Population")


class NoEducationBlock(Exception):
    """Raised when a town has no block in the Education section. Not an error."""


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


def _find_town_row(labels: list[str], header_row: int, town: str) -> int:
    """Return the row of `town`'s heading within the Education section.

    The two rows below it must carry the two indicator labels.
    """
    matches = []
    for index in range(header_row, len(labels)):        # labels[index] is sheet row index + 1
        text = labels[index]
        if text in NEXT_SECTION_HEADERS:
            break
        if text == town.strip():
            matches.append(index + 1)

    if not matches:
        raise NoEducationBlock(town)
    if len(matches) > 1:
        raise ValueError(
            f"AMBIGUOUS: {len(matches)} rows in the Education section are headed "
            f"'{town}' (rows {matches})."
        )

    town_row = matches[0]
    for _, label, offset in INDICATORS:
        found = labels[town_row + offset - 1] if town_row + offset - 1 < len(labels) else ""
        if found != label:
            raise ValueError(
                f"'{town}' block starts at row {town_row}, but row {town_row + offset} "
                f"reads {found!r} where '{label}' was expected -- not writing into it."
            )
    return town_row


def _find_year_column(sheet, header_row: int, year: int) -> int:
    """Return the column for `year` in the Education section's header row,
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


def write_education(sheet, cache_files: list[Path]) -> tuple[list[str], int, int]:
    """Audit and write the latest year for every cached town.

    Separate from the Excel session handling so that it can be run
    against any object with the xlwings Sheet interface (see
    tests/fake_xlwings_sheet.py). Returns
    (result lines, written count, flagged count).
    """
    results: list[str] = []
    written_count = 0
    flagged_count = 0
    no_block: list[str] = []

    labels = _column_a(sheet)
    header_row = _find_section_header_row(labels)

    for path in cache_files:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)

        town = data["town"]
        year = int(data["year"])

        try:
            town_row = _find_town_row(labels, header_row, town)
        except NoEducationBlock:
            no_block.append(town)
            continue
        except ValueError as exc:
            results.append(f"{town}: SKIPPED (row-finding) — {exc}")
            flagged_count += len(INDICATORS)
            continue

        for key, label, offset in INDICATORS:
            series = data.get("indicators", {}).get(key, {})
            if str(year) not in series:
                results.append(f"{town} {label}: no {year} figure, skipped")
                flagged_count += 1
                continue

            value = series[str(year)]
            source_series = {int(y): v for y, v in series.items()}
            row = town_row + offset

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
                gaps = {
                    y: abs(v - source_series[y]) / abs(source_series[y])
                    for y, v in existing.items()
                    if y in source_series and source_series[y] and abs(v - source_series[y]) > 0.05
                }
                note = ""
                if gaps:
                    worst = max(gaps, key=gaps.get)
                    note = (
                        f" (note: {len(gaps)} earlier year(s) differ from ACARA's current "
                        f"file, not changed — {min(gaps)}-{max(gaps)}, largest {gaps[worst]:.1%} "
                        f"in {worst}: sheet {existing[worst]} vs source {source_series[worst]})"
                    )
                results.append(f"{town} {label}: WRITTEN {year} = {value} -> {cell.address}{note}")
                written_count += 1
            else:
                results.append(f"{town} {label}: FLAGGED, not written — {report.summary_line()}")
                flagged_count += 1

    if no_block:
        results.append(
            f"No block in the {SHEET_NAME} sheet's Education section, nothing to write "
            f"(expected): {', '.join(no_block)}"
        )
    return results, written_count, flagged_count


def update_education(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    cache_files = sorted(cache_dir.glob(CACHE_GLOB))
    if not cache_files:
        raise FileNotFoundError(
            f"No {CACHE_GLOB} files found in {cache_dir} -- run fetch_schools.py first "
            f"(python run_update.py --only schools)."
        )

    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(str(xlsx_path))
        try:
            sheet = wb.sheets[SHEET_NAME]
            results, written_count, flagged_count = write_education(sheet, cache_files)
            if written_count:
                wb.save()
        finally:
            wb.close()
    finally:
        app.quit()

    results.append("")
    results.append(f"Summary: {written_count} written, {flagged_count} flagged for review.")
    return results


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--visible"]
    visible = "--visible" in sys.argv

    if len(args) != 2:
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/schools dir> [--visible]")
        sys.exit(1)

    results = update_education(Path(args[0]), Path(args[1]), visible=visible)
    for line in results:
        print(line)
