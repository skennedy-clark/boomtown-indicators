"""
regional-indicators/transform/xlsx_update/update_population_nrw_lga.py

Writes LGA-level non-resident worker counts into the Population sheet.

Input:  cache/population/<slug>_population_nrw.json, produced by
        fetchers/fetch_population_nrw.py. Each town's file carries its
        LGA's full series as well as the town's own UCL figure.
Target: the "LGA" section of the Population sheet, row "Non-resident
        workers on-shift" with sub-label "Non-residents (LGA)".

Several towns share an LGA; each LGA is written once.

The source provides the full multi-year LGA series, so every write is
checked against it with audit_historical_series in addition to the cell
audit.

Sub-labels: LGA_SUB_LABEL_DEFAULT applies to every LGA.
LGA_SUB_LABEL_OVERRIDES is available for a workbook whose wording
differs for a particular LGA; verify the wording in the target workbook
before adding an entry.

The UCL-level figures are written by update_population_nrw.py. The
workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_population_nrw_lga.py <workbook.xlsx> <cache/population dir> [--visible]
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