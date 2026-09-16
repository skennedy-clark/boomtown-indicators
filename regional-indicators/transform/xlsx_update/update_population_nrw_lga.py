"""
regional-indicators/transform/xlsx_update/update_population_nrw_lga.py
--------------------------------------------------------------------------
Writes the LGA-level "Non-resident workers on-shift" figure (Isaac,
Maranoa, Toowoomba, Western Downs) from fetch_population_nrw.py's cached
output. Split out from what used to be a combined LGA+UCL script -- see
update_population_nrw.py for the UCL-level sibling.

Reads the SAME cache/population/*_population_nrw.json files as the UCL
script (each town's cache file carries both its LGA's series and its
own UCL figure), but writes each LGA's figure exactly ONCE, not once
per town sharing that LGA (several towns share an LGA -- e.g.
Chinchilla/Dalby/Miles/Wandoan/Tara/Wallumbilla all under Western
Downs).

Uses xlwings (real Excel via COM automation), NOT openpyxl -- see
base.py's docstring for why.

Unlike the UCL script, every LGA write here has FULL multi-year source
history available (the NRW-LGA source file gives 2006/2008-2025, not
just the latest year), so every write goes through the ground-truth
historical cross-check in audit.py's audit_historical_series -- real
data confirmed this correctly clears false positives the shape-based
check alone produced (Isaac's 2012 figure looked like an isolated
miscopy by shape, but exactly matches the real source).

Confirmed working against real data (2026-09-15, as part of the
combined script this was split from): Western Downs, Maranoa, Isaac,
and Toowoomba all wrote correctly once the ground-truth check was
added.

Usage:
    python update_population_nrw_lga.py <path-to-Indicators_Data-Charts.xlsx> <cache/population dir> [--visible]

Note on sub-labels: most LGA rows read "Non-residents (LGA)". An
earlier version of this hardcoded a Toowoomba-specific override
("Toowoomba Non-residents (LGA)") based on the 2026-Aug reference
file's wording -- confirmed WRONG for the real 2025-based working
file, which uses the plain default for Toowoomba too. No override is
applied by default; if a genuinely different-wording workbook needs
one, check the actual target file's wording first (the way this was
checked before removing the wrong override), don't guess.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

from base import write_one

INDICATOR_NAME = "Non-resident workers on-shift"
SHEET_NAME = "Population"
LGA_SUB_LABEL_DEFAULT = "Non-residents (LGA)"
LGA_SUB_LABEL_OVERRIDES: dict[str, str] = {}


def update_population_nrw_lga(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    cache_files = sorted(cache_dir.glob("*_population_nrw.json"))
    if not cache_files:
        raise FileNotFoundError(
            f"No *_population_nrw.json files found in {cache_dir} -- "
            f"run fetch_population_nrw.py first."
        )

    results = []
    written_count = 0
    flagged_count = 0
    lgas_written: set[str] = set()

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

                lga = data.get("lga")
                lga_series = data.get("lga_nrw_on_shift_by_year")

                if not lga_series or lga in lgas_written:
                    continue

                latest_year_str = max(lga_series, key=int)
                latest_value = lga_series[latest_year_str]
                sub_label = LGA_SUB_LABEL_OVERRIDES.get(lga, LGA_SUB_LABEL_DEFAULT)

                try:
                    report, coord = write_one(
                        sheet, lga, INDICATOR_NAME, sub_label,
                        int(latest_year_str), latest_value,
                        source_series={int(y): v for y, v in lga_series.items()},
                    )
                    if report.safe_to_write:
                        results.append(
                            f"{lga} (LGA): WRITTEN {latest_year_str} = {latest_value:,} -> {coord}"
                        )
                        written_count += 1
                        any_written = True
                    else:
                        results.append(f"{lga} (LGA): FLAGGED, not written — {report.summary_line()}")
                        flagged_count += 1
                except ValueError as exc:
                    results.append(f"{lga} (LGA): SKIPPED (row-finding) — {exc}")
                    flagged_count += 1

                lgas_written.add(lga)

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

    results = update_population_nrw_lga(xlsx_path, cache_dir, visible=visible)
    for line in results:
        print(line)