"""
regional-indicators/transform/xlsx_update/update_population_erp.py
-----------------------------------------------------------------------
Writes the main SA2-level "Population (ERP)" figure from
fetch_population_erp.py's cached output into the workbook's SA2 section
on the Population sheet.

Uses section="SA2" (base.py's new disambiguator) to land in the right
section -- this indicator name collides with the LGA section's own
"Population (ERP)" row for the same town names (confirmed real:
Goondiwindi's LGA figure is 42% different from its SA2 figure), so
section is not optional here the way it's been unnecessary for the UCL
writes so far.

IMPORTANT: matches by town.sa2_name, NOT town.name. The SA2 section's
row labels are real ABS SA2 names, not town names -- e.g. Toowoomba's
three sub-areas need "Toowoomba - Central" / "North Toowoomba -
Harlaxton" / "Toowoomba - West", not "Toowoomba (Central)" etc. This
was flagged as an open gap weeks before fetch_population_erp.py existed
to actually need it.

Full source history is available (fetch_population_erp.py's
series_by_year field), so every write goes through the ground-truth
historical cross-check, same as update_population_nrw_lga.py -- more
reliable than the shape-based guess for existing-data problems.

Uses xlwings (real Excel via COM automation), NOT openpyxl -- see
base.py's docstring for why.

*** NOT YET TESTED against a live Excel instance. *** The section
parameter itself is tested (against a mock reproducing the real
Goondiwindi LGA/SA2 collision), but this script's actual write against
the real workbook has not been run yet -- test against a throwaway
copy first, same as every other first run this project.

Usage:
    python update_population_erp.py <path-to-Indicators_Data-Charts.xlsx> <cache/population dir> [--visible]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

from base import write_one

INDICATOR_NAME = "Population (ERP)"
SHEET_NAME = "Population"
SECTION = "SA2"


def update_population_erp(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    cache_files = sorted(cache_dir.glob("*_population_erp.json"))
    if not cache_files:
        raise FileNotFoundError(
            f"No *_population_erp.json files found in {cache_dir} -- "
            f"run fetch_population_erp.py first."
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

                # sa2_name is the real ABS SA2 region name, captured by
                # fetch_population_erp.py from the matched QRSIS region
                # string -- confirmed reliable, including for towns whose
                # SA2 name differs from town.name (Toowoomba's three
                # sub-areas each get their own correct, distinct name).
                town_display_name = data.get("town")
                sa2_name = data.get("sa2_name") or town_display_name
                series_by_year = data.get("series_by_year") or {}
                latest_year = data.get("year")
                latest_value = data.get("value")

                if latest_year is None or latest_value is None:
                    results.append(f"{town_display_name}: no data, skipped")
                    continue

                try:
                    report, coord = write_one(
                        sheet, sa2_name, INDICATOR_NAME, None,
                        int(latest_year), latest_value,
                        source_series={int(y): v for y, v in series_by_year.items()},
                        section=SECTION,
                    )
                    if report.safe_to_write:
                        results.append(
                            f"{town_display_name} (SA2 '{sa2_name}'): WRITTEN "
                            f"{latest_year} = {latest_value:,} -> {coord}"
                        )
                        written_count += 1
                        any_written = True
                    else:
                        results.append(
                            f"{town_display_name} (SA2 '{sa2_name}'): FLAGGED, "
                            f"not written — {report.summary_line()}"
                        )
                        flagged_count += 1
                except ValueError as exc:
                    results.append(f"{town_display_name} (SA2 '{sa2_name}'): SKIPPED (row-finding) — {exc}")
                    flagged_count += 1

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

    if len(args) != 2:
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/population dir> [--visible]")
        sys.exit(1)

    xlsx_path = Path(args[0])
    cache_dir = Path(args[1])

    results = update_population_erp(xlsx_path, cache_dir, visible=visible)
    for line in results:
        print(line)