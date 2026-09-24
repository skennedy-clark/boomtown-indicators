"""
regional-indicators/transform/xlsx_update/update_business.py
-----------------------------------------------------------------------
Writes fetch_business.py's cached output (NPP/PP business counts by
turnover band) into the Business sheet's two parallel sections.

CONFIRMED REAL STRUCTURE (2026-09-23/24 inspection of the reference
workbook):
  Row 1: fiscal-year strings ("2008/09", ...) starting at column B.
  Row 2: "Non-primary production" section header.
  Then repeating 5-row region blocks: name row, then "0k-50k",
    "50k-200k", "200k-2m", "2m+" in that fixed order.
  Later: "Primary production" section header, then the SAME 12 regions
    in the SAME order, same 5-row block shape.
  Some region-name rows carry a live SUM formula (e.g. "=SUM(Q4:Q7)")
    totalling the 4 bands below them -- never written to directly,
    only the 4 band rows are write targets.

YEAR-COLUMN FINDING -- the tricky part, confirmed real and deliberately
designed around, not guessed at:
  Row 1 has fiscal-year-LOOKING labels in TWO places: the real data
  range (B onward) and a later range of growth-rate/working-out
  columns Steve confirmed are "on the fly working-out... the sort of
  thing that is problematic with this approach" -- ad hoc analysis
  cells that happen to have year labels above them too, not real data.
  There's a genuine BLANK GAP in row 1 between the two. This finder
  stops at the FIRST gap (a None value) when scanning row 1 rightward
  from column B -- never continues scanning past it, so it can never
  latch onto a lookalike label in the working-out zone, and never
  accidentally creates a new column IN that zone either. Matches
  Steve's own description of how he'd do this by eye: "run my eye down
  the data and see that Q is complete... start recording into R".

VALUE SANITY CHECK, per Steve's own heuristic (worth restating
exactly, since it's a genuinely good one): "a value in R... clearly
not a positive, integer/float in the order of tens to hundreds" is
"growth/decay working or some other data discovery", not real data.
Reused rather than reinvented -- this is exactly what audit_cell's
existing UNEXPECTED_CONTENT/FORMULA detection already catches (a
formula, or an existing value that doesn't match what's expected), so
this script goes through the same audit_cell/audit_series/
WriteAuditReport pipeline as every other indicator rather than adding
a separate, parallel check.

Toowoomba composite: 3 SA2s (Central/Harlaxton/West), NOT 4 -- Steve's
explicit call: "we're not tracking Toowoomba - East because there are
no FIFO workers living there and it is not affected by the boom-town
dynamics". Matches fetch_business.py's SA2_REGIONS exactly already;
nothing to reconcile here, just documenting why.

*** NOT YET TESTED against a live Excel instance. *** Row/column
finding logic is tested against a mock built to match the real
structure, but the actual write against the real workbook hasn't run
yet -- test against a throwaway copy first, same as every other
wiring script's first live test.

Usage:
    python update_business.py <path-to-Indicators_Data-Charts.xlsx> <cache/business dir> [--visible]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).parent))
from audit import audit_cell, audit_series, WriteAuditReport, apply_write_formatting, describe_write_outcome
from base import _fiscal_label

SHEET_NAME = "Business"
FIRST_YEAR_COLUMN = 2   # column B -- confirmed
YEAR_HEADER_ROW = 1
REGION_BLOCK_SIZE = 5   # name row + 4 turnover-band rows, confirmed

SECTION_HEADERS = {
    "NPP": "Non-primary production",
    "PP":  "Primary production",
}
BAND_LABELS = ["0k-50k", "50k-200k", "200k-2m", "2m+"]


def _find_year_column(sheet, year: int) -> int:
    """Scans row 1 rightward from column B, stopping at the FIRST gap
    (a None value) -- never continues past it into the growth-rate
    working-out zone confirmed to sit further right with its own
    lookalike year labels. If `year` is found before the gap, returns
    its column. Otherwise creates a new column immediately after the
    last real one (immediately before the gap), matching exactly how
    Steve described doing this by eye: find the last complete year,
    the next column is where new data goes.
    """
    used = sheet.used_range
    max_col = used.last_cell.column

    header_row = sheet.range((YEAR_HEADER_ROW, FIRST_YEAR_COLUMN), (YEAR_HEADER_ROW, max_col)).value
    if not isinstance(header_row, list):
        header_row = [header_row]

    last_real_col = None
    for offset, label in enumerate(header_row):
        col = FIRST_YEAR_COLUMN + offset
        if label is None:
            break  # the real data range ends here -- stop, don't scan further
        if isinstance(label, str) and "/" in label:
            start_year = int(label.split("/")[0])
            end_year = start_year + 1
            if end_year == year:
                return col
            last_real_col = col

    if last_real_col is None:
        raise ValueError(
            f"Could not find any real fiscal-year labels in row {YEAR_HEADER_ROW} "
            f"of sheet '{sheet.name}' before the first gap -- can't determine "
            f"where to add a new year column safely."
        )

    new_col = last_real_col + 1
    year_cell = sheet.cells(YEAR_HEADER_ROW, new_col)
    year_cell.value = _fiscal_label(year)
    year_cell.number_format = "General"
    return new_col


def _find_section_row(sheet, section_label: str) -> int:
    used = sheet.used_range
    max_row = used.last_cell.row
    col_a = sheet.range((1, 1), (max_row, 1)).value
    if not isinstance(col_a, list):
        col_a = [col_a]

    matches = [1 + offset for offset, val in enumerate(col_a) if val == section_label]
    if len(matches) == 0:
        raise ValueError(f"Could not find section header '{section_label}' in sheet '{sheet.name}'.")
    if len(matches) > 1:
        raise ValueError(f"AMBIGUOUS: section '{section_label}' appears at rows {matches}.")
    return matches[0]


def _find_region_row(sheet, section_row: int, region_name: str, next_section_row: int | None) -> int:
    """Search within one section (between its header and the next
    section's header, or the sheet end) for region_name's own row.
    """
    used = sheet.used_range
    max_row = next_section_row - 1 if next_section_row else used.last_cell.row
    col_a = sheet.range((section_row + 1, 1), (max_row, 1)).value
    if not isinstance(col_a, list):
        col_a = [col_a]

    matches = [
        section_row + 1 + offset for offset, val in enumerate(col_a)
        if val == region_name
    ]
    if len(matches) == 0:
        raise ValueError(
            f"Could not find region '{region_name}' within the section starting "
            f"row {section_row} in sheet '{sheet.name}'. This function does not "
            f"guess or create a new region block for you."
        )
    if len(matches) > 1:
        raise ValueError(f"AMBIGUOUS: region '{region_name}' appears at rows {matches}.")
    return matches[0]


def _find_band_row(sheet, region_row: int, band_label: str) -> int:
    for row in range(region_row + 1, region_row + REGION_BLOCK_SIZE):
        if sheet.cells(row, 1).value == band_label:
            return row
    raise ValueError(
        f"Could not find turnover band '{band_label}' within rows "
        f"{region_row + 1}-{region_row + REGION_BLOCK_SIZE - 1} (region block "
        f"starting row {region_row}) in sheet '{sheet.name}'."
    )


def _col_letter(col: int) -> str:
    """1 -> 'A', 26 -> 'Z', 27 -> 'AA', etc. -- correct for any column,
    not just the single-letter range every other column reference in
    this project has stayed within so far."""
    letters = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _write_region_total_formula(sheet, region_row: int, col: int) -> str | None:
    """Confirmed real pattern (2026-09-24, all 12 regions checked): the
    region-header row carries a live =SUM(<col><first band row>:<col>
    <last band row>) formula in every existing year column -- never
    written to directly (it's the total, not a data cell), but a
    genuinely NEW column starts empty and needs this formula extended
    into it too, or the total for the new year would silently be
    missing even though all 4 band values are there. Different from
    every other write in this project: this DELIBERATELY writes a
    formula, not a value -- safe here specifically because the target
    cell is confirmed empty (a brand new column), never overwriting an
    existing formula or value. Returns the formula written, or None if
    the cell wasn't empty (left untouched rather than guessing whether
    it's safe to overwrite).
    """
    cell = sheet.cells(region_row, col)
    if cell.value is not None or (cell.formula and str(cell.formula).startswith("=")):
        return None  # not empty -- don't guess, leave it alone

    col_letter = _col_letter(col)
    first_band_row = region_row + 1
    last_band_row = region_row + REGION_BLOCK_SIZE - 1
    formula = f"=SUM({col_letter}{first_band_row}:{col_letter}{last_band_row})"
    cell.formula = formula
    cell.number_format = "General"
    return formula


def _read_existing_series(sheet, row: int, exclude_col: int | None = None) -> dict:
    used = sheet.used_range
    max_col = used.last_cell.column
    header_row = sheet.range((YEAR_HEADER_ROW, FIRST_YEAR_COLUMN), (YEAR_HEADER_ROW, max_col)).value
    values = sheet.range((row, FIRST_YEAR_COLUMN), (row, max_col)).value
    if not isinstance(header_row, list):
        header_row, values = [header_row], [values]

    series = {}
    for offset, (label, val) in enumerate(zip(header_row, values)):
        col = FIRST_YEAR_COLUMN + offset
        if col == exclude_col:
            continue
        if label is None:
            break  # stop at the same gap the year-finder stops at
        if isinstance(label, str) and "/" in label and isinstance(val, (int, float)):
            year = int(label.split("/")[0]) + 1
            series[year] = val
    return series


def _write_business_row(sheet, row: int, year: int, value):
    col = _find_year_column(sheet, year)
    cell = sheet.cells(row, col)

    cell_result = audit_cell(cell, value)
    existing_series = _read_existing_series(sheet, row, exclude_col=col)
    series_result = audit_series(existing_series, year, value)

    report = WriteAuditReport(cell_result, series_result)
    if report.safe_to_write or report.should_write_with_flag:
        cell.value = value
    apply_write_formatting(cell, report.safe_to_write, report.should_write_with_flag)

    return report, cell.address


def update_business(xlsx_path: Path, cache_dir: Path, year: int, visible: bool = False) -> list[str]:
    cache_files = sorted(cache_dir.glob("*_business.json"))
    if not cache_files:
        raise FileNotFoundError(
            f"No *_business.json files found in {cache_dir} -- run fetch_business.py first."
        )

    results = []
    written_count = 0
    flagged_count = 0

    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(str(xlsx_path))
        try:
            sheet = wb.sheets[SHEET_NAME]
            any_written = False

            npp_section_row = _find_section_row(sheet, SECTION_HEADERS["NPP"])
            pp_section_row = _find_section_row(sheet, SECTION_HEADERS["PP"])

            for path in cache_files:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)

                region = data["region"]
                indicators = data.get("indicators", {})

                for production, section_row, next_section in (
                    ("npp", npp_section_row, pp_section_row),
                    ("pp",  pp_section_row,  None),
                ):
                    band_values = indicators.get(production, {})
                    if not band_values:
                        results.append(f"{region} {production}: no data, skipped")
                        continue

                    try:
                        region_row = _find_region_row(sheet, section_row, region, next_section)
                    except ValueError as exc:
                        results.append(f"{region} {production}: SKIPPED (row-finding) — {exc}")
                        flagged_count += len(BAND_LABELS)
                        continue

                    for band_label in BAND_LABELS:
                        if band_label not in band_values:
                            results.append(f"{region} {production} {band_label}: no data, skipped")
                            continue
                        value = band_values[band_label]

                        try:
                            row = _find_band_row(sheet, region_row, band_label)
                            report, coord = _write_business_row(sheet, row, year, value)
                            line, is_written, is_flagged = describe_write_outcome(
                                f"{region} {production} {band_label}", report, coord, f"{year} = {value}"
                            )
                            results.append(line)
                            if is_written:
                                written_count += 1
                                any_written = True
                            if is_flagged:
                                flagged_count += 1
                        except ValueError as exc:
                            results.append(f"{region} {production} {band_label}: SKIPPED (row-finding) — {exc}")
                            flagged_count += 1

                    # Region-header row's live SUM formula -- extend it into
                    # the new year's column too (see _write_region_total_formula's
                    # docstring), confirmed only into a genuinely empty cell.
                    year_col = _find_year_column(sheet, year)
                    total_formula = _write_region_total_formula(sheet, region_row, year_col)
                    if total_formula:
                        results.append(f"{region} {production} TOTAL: formula extended -> {total_formula}")
                        any_written = True
                    else:
                        results.append(
                            f"{region} {production} TOTAL: cell not empty, left untouched "
                            f"-- check manually whether a total formula is needed here"
                        )

            if any_written:
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

    if len(args) != 3:
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/business dir> <year> [--visible]")
        sys.exit(1)

    xlsx_path = Path(args[0])
    cache_dir = Path(args[1])
    year = int(args[2])

    results = update_business(xlsx_path, cache_dir, year, visible=visible)
    for line in results:
        print(line)