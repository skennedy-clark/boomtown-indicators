"""
update_population_ucl.py -- reads fetch_population_ucl.py's cached output
(cache/population/{slug}_population_ucl.json) and writes each town's
LATEST year into Indicators_Data-Charts.xlsx's "UCL" section on the
Population sheet.

Uses xlwings (real Excel via COM automation), NOT openpyxl -- openpyxl
was confirmed to corrupt this specific workbook badly enough that even
Excel's own repair couldn't recover the result. See base.py's docstring
for the full explanation.

*** NOT YET TESTED against a live Excel instance *** -- same caveat as
base.py. Test against a throwaway copy first.

Opens ONE Excel session for the whole run (not one per town) -- starting
and stopping Excel per town would be needlessly slow for a run covering
17 towns.

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

Writes the file in place (saves the same path it opened). Run
fetch_population_ucl.py first so the cache JSON files actually exist.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

from base import _find_year_column, _find_town_indicator_row

INDICATOR_NAME = "Estimated resident population (a) by urban centre and locality"
SHEET_NAME = "Population"


def update_population_ucl(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    """Update every town found in cache_dir's *_population_ucl.json files
    with its latest available year, in a single Excel session. Returns a
    list of human-readable result lines for logging/review.
    """
    cache_files = sorted(cache_dir.glob("*_population_ucl.json"))
    if not cache_files:
        raise FileNotFoundError(
            f"No *_population_ucl.json files found in {cache_dir} -- "
            f"run fetch_population_ucl.py first."
        )

    results = []
    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(str(xlsx_path))
        try:
            sheet = wb.sheets[SHEET_NAME]

            for path in cache_files:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)

                town = data["town"]
                year_vals = data["population_by_year"]
                if not year_vals:
                    results.append(f"{town}: no data in cache file, skipped")
                    continue

                # Latest year determined from the data itself, never
                # hardcoded -- correct in 2026, 2027, or any future year
                # without a code change.
                latest_year_str = max(year_vals, key=int)
                latest_value = year_vals[latest_year_str]
                latest_year = int(latest_year_str)

                try:
                    col = _find_year_column(sheet, latest_year)
                    row = _find_town_indicator_row(sheet, town, INDICATOR_NAME, sub_label=None)
                    cell = sheet.cells(row, col)
                    cell.value = latest_value
                    cell.number_format = "General"  # don't trust Excel's inherited format --
                                                     # confirmed to silently pick up a
                                                     # percentage format from an adjacent
                                                     # cell on real data
                    coord = cell.address
                    results.append(f"{town}: {latest_year} = {latest_value:,} -> {coord}")
                except ValueError as exc:
                    # A town not found, or an ambiguous match, shouldn't
                    # silently corrupt something else -- surface it and
                    # move on to the next town.
                    results.append(f"{town}: SKIPPED — {exc}")

            wb.save()
        finally:
            wb.close()
    finally:
        app.quit()

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