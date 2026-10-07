"""
regional-indicators/transform/xlsx_update/update_education.py
----------------------------------------------------------------
Writes fetch_schools.py's cached output (ACARA School Profile, summed
by town postcode) into the Exogenous sheet's Education section.

Built 2026-10-07.

SHEET LAYOUT (confirmed against the real 2025 working file)
    Education       | <years across this row>        <- section header row
    Chinchilla      |                                <- town row, column A only
    Full Time Equivalent Enrolments    | FTE Enrolments    | ...
    Full Time Equivalent Teaching Staff| FTE Teaching Staff| ...
    Dalby
    ...
Like the Fuel section, Education has its OWN year header row (the
"Education" row itself, 2008 onward), which is not the sheet's row 1.
The year column is found from that row; if the new year isn't there
yet it is appended straight after the last year in the row and the
year written in as its header.

A town's block is found by its town header row inside the Education
section -- the row whose next two rows carry the two indicator labels,
checked rather than assumed -- so the Rainfall and Fuel blocks that
share the same town names further up and down the sheet are never
touched. Towns that are fetched but have no block (Toowoomba's three
sub-areas, Shepparton, Yarram) are listed once as expected.

Each write goes through audit.py's cell, series and ground-truth
historical checks. The cache carries every year from 2008, so the
sheet's existing figures are compared with the source; where earlier
years differ the result line says how many and the largest gap, but
only the NEW year is written. (ACARA's back series has been restated
for some towns since the early columns were filled in.)

Uses xlwings (real Excel via COM automation), NOT openpyxl -- see
base.py's docstring for why.

Tested 2026-10-07 against the real 2025 working file's cell contents
through tests/fake_xlwings_sheet.py (a stand-in for the Excel sheet
object): 24 written, all equal to the 2026 reference file. First real
Excel run still to be confirmed.

NOTE: the section-handling helpers here (_column_a, _find_year_column,
_read_existing_series) are deliberately the same as update_fuel.py's.
Worth pulling into one shared module once both are confirmed in real
Excel -- see TODO.md.

Usage:
    python update_education.py <path-to-Indicators_Data-Charts.xlsx> <cache/schools dir> [--visible]
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

# (cache key, row label, rows below the town header row)
INDICATORS = (
    ("fte_enrolments",     "Full Time Equivalent Enrolments",     1),
    ("fte_teaching_staff", "Full Time Equivalent Teaching Staff", 2),
)
NEXT_SECTION_HEADERS = ("Fuel", "Rainfall", "Business", "Crime",
                        "Employment", "Housing", "Income", "Population")


class NoEducationBlock(Exception):
    """This town has no block in the Education section. Not an error."""


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
    """Row of `town`'s header inside the Education section. The two
    rows below it must carry the two indicator labels."""
    matches = []
    for index in range(header_row, len(labels)):        # 0-based index == row index+1
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
    """Column for `year` in the Education section's own header row,
    creating it straight after the last year in that row if needed."""
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
    """Audit and write each cached town's latest year into its block.
    Separate from the Excel open/save handling so it can be exercised
    against any sheet-like object. Returns
    (result lines, written count, flagged count)."""
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
