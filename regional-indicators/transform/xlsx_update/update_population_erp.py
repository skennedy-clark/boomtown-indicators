"""
regional-indicators/transform/xlsx_update/update_population_erp.py

Writes SA2-level estimated resident population (ERP) into the Population
sheet.

Input:  cache/population/<slug>_population_erp.json, produced by
        fetchers/fetch_population_erp.py.
Target: the "SA2" section of the Population sheet, row "Population
        (ERP)" in each SA2's block.

The row name "Population (ERP)" also exists in the LGA section, under
some of the same region names, so every lookup passes section="SA2".

Blocks are matched on the SA2 name recorded in the cache file, not the
town name: the section is headed by ABS SA2 names (for example
"Toowoomba - Central"), which differ from town names for several towns.

The cache carries the full source series, so every write is checked
against it with audit_historical_series.

The LGA-level row is written by update_population_erp_lga.py. The
workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_population_erp.py <workbook.xlsx> <cache/population dir> [--visible]
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

                # The block heading is the ABS SA2 name captured by the fetcher,
                # which differs from the town name for several towns.
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