"""
regional-indicators/transform/xlsx_update/update_population_ucl.py -- reads
fetch_population_ucl.py's cached output
(cache/population/{slug}_population_ucl.json) and writes each town's
LATEST year into Indicators_Data-Charts.xlsx's "UCL" section on the
Population sheet.

Uses xlwings (real Excel via COM automation), NOT openpyxl -- openpyxl
was confirmed to corrupt this specific workbook badly enough that even
Excel's own repair couldn't recover the result. See base.py's docstring
for the full explanation.

Runs the pre-write audits from audit.py on every town before writing --
this workbook has been hand-edited for years and is known to contain
crud (stray values, leftover formulas from ad-hoc analysis) in cells
this script writes into, and its existing historical data has not been
independently verified. A clean-looking write can still be wrong if it
lands in the wrong row, or right if a "flagged" cell turns out to be a
real, correctly-revised figure -- the audits surface exactly that
ambiguity for a human to resolve, they don't resolve it themselves.

*** NOT YET TESTED against a live Excel instance *** -- same caveat as
base.py. Test against a throwaway copy first.

Opens ONE Excel session for the whole run (not one per town).

Deliberately targets the "UCL" block (indicator name "Estimated resident
population (a) by urban centre and locality") and NOT each town's main
"Population (ERP)" row -- those are a different geography level (SA2 or
LGA depending on town) fed by a different, not-yet-built fetcher.

Deliberately matches on indicator name only, no sub_label -- Toowoomba's
sub-label reads "Toowoomba Residents (UCL)" rather than the "Residents
(UCL)" every other town uses.

Usage:
    python update_population_ucl.py <path-to-Indicators_Data-Charts.xlsx> <cache/population dir> [--visible]

--visible runs Excel on-screen rather than in the background -- worth
using for your first real run, so you can watch it happen.

Writes the file in place for every town that passes both audits clean.
Flagged towns are NOT written -- rerun after resolving what the report
says, rather than this script guessing on your behalf.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

from base import _find_year_column, _find_town_indicator_row, read_existing_series
from audit import audit_cell, audit_series, WriteAuditReport

INDICATOR_NAME = "Estimated resident population (a) by urban centre and locality"
SHEET_NAME = "Population"


def update_population_ucl(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    """Update every town found in cache_dir's *_population_ucl.json files
    with its latest available year, auditing each one before writing.
    Returns a list of human-readable result lines -- written towns show
    the cell they landed in, flagged towns show why they were skipped
    and nothing about them was changed.
    """
    cache_files = sorted(cache_dir.glob("*_population_ucl.json"))
    if not cache_files:
        raise FileNotFoundError(
            f"No *_population_ucl.json files found in {cache_dir} -- "
            f"run fetch_population_ucl.py first."
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

            for path in cache_files:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)

                town = data["town"]
                year_vals = data["population_by_year"]
                if not year_vals:
                    results.append(f"{town}: no data in cache file, skipped")
                    continue

                latest_year_str = max(year_vals, key=int)
                latest_value = year_vals[latest_year_str]
                latest_year = int(latest_year_str)

                try:
                    col = _find_year_column(sheet, latest_year)
                    row = _find_town_indicator_row(sheet, town, INDICATOR_NAME, sub_label=None)
                except ValueError as exc:
                    results.append(f"{town}: SKIPPED (row-finding) — {exc}")
                    flagged_count += 1
                    continue

                cell = sheet.cells(row, col)
                cell_result = audit_cell(cell, latest_value)
                existing_series = read_existing_series(sheet, row, exclude_col=col)
                series_result = audit_series(existing_series, latest_year, latest_value)
                report = WriteAuditReport(cell_result, series_result)

                if report.safe_to_write:
                    cell.value = latest_value
                    cell.number_format = "General"
                    coord = cell.address
                    results.append(f"{town}: WRITTEN {latest_year} = {latest_value:,} -> {coord}")
                    written_count += 1
                    any_written = True
                else:
                    results.append(f"{town}: FLAGGED, not written — {report.summary_line()}")
                    flagged_count += 1

            if any_written:
                wb.save()
        finally:
            wb.close()
    finally:
        app.quit()

    results.append("")
    results.append(f"Summary: {written_count} written, {flagged_count} flagged for review, "
                    f"{len(cache_files) - written_count - flagged_count} skipped (no data).")
    return results


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--visible"]
    visible = "--visible" in sys.argv

    if len(args) != 2:
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/population dir> [--visible]")
        sys.exit(1)

    xlsx_path = Path(args[0])
    cache_dir = Path(args[1])

    results = update_population_ucl(xlsx_path, cache_dir, visible=visible)
    for line in results:
        print(line)