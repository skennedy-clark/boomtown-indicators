"""
regional-indicators/transform/xlsx_update/update_population_nrw.py

Writes urban-centre (UCL) non-resident worker counts into the Population
sheet.

Input:  cache/population/<slug>_population_nrw.json, produced by
        fetchers/fetch_population_nrw.py.
Target: the "UCL" section of the Population sheet, row "Non-resident
        workers on-shift" with sub-label "Non-resident workers (UCL)".

The LGA-level figures in the same cache files are written by
update_population_nrw_lga.py.

The source publishes only the latest year at UCL level, so these writes
are checked with the cell and shape-based series audits only; no
historical comparison is possible.

--deep-audit adds context to any flagged town: the corresponding LGA
series (already in the cache file) is examined for a change in the same
year. A genuine regional change normally appears at both levels. The
output informs review; it does not change what is written.

The workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_population_nrw.py <workbook.xlsx> <cache/population dir> [--visible] [--deep-audit]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

from base import write_one, deep_audit_context

INDICATOR_NAME = "Non-resident workers on-shift"
SHEET_NAME = "Population"
UCL_SUB_LABEL = "Non-resident workers (UCL)"


def update_population_nrw(
    xlsx_path: Path, cache_dir: Path, visible: bool = False, deep_audit: bool = False
) -> list[str]:
    cache_files = sorted(cache_dir.glob("*_population_nrw.json"))
    if not cache_files:
        raise FileNotFoundError(
            f"No *_population_nrw.json files found in {cache_dir} -- "
            f"run fetch_population_nrw.py first."
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
                lga = data.get("lga")
                lga_series = data.get("lga_nrw_on_shift_by_year")
                ucl_latest = data.get("ucl_nrw_latest")

                if not ucl_latest:
                    results.append(f"{town} (UCL): no data, skipped")
                    continue

                try:
                    report, coord = write_one(
                        sheet, town, INDICATOR_NAME, UCL_SUB_LABEL,
                        ucl_latest["year"], ucl_latest["value"],
                    )
                    if report.safe_to_write:
                        results.append(
                            f"{town} (UCL): WRITTEN {ucl_latest['year']} = "
                            f"{ucl_latest['value']:,} -> {coord}"
                        )
                        written_count += 1
                        any_written = True
                    else:
                        results.append(f"{town} (UCL): FLAGGED, not written — {report.summary_line()}")
                        flagged_count += 1
                        if deep_audit:
                            flagged_years = (
                                report.series_audit.outlier_years
                                if report.series_audit.outlier_years
                                else [ucl_latest["year"]]
                            )
                            lga_series_int = (
                                {int(y): v for y, v in lga_series.items()} if lga_series else {}
                            )
                            results.append(deep_audit_context(lga, lga_series_int, flagged_years))
                except ValueError as exc:
                    results.append(f"{town} (UCL): SKIPPED (row-finding) — {exc}")
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
    flags = {"--visible", "--deep-audit"}
    args = [a for a in sys.argv[1:] if a not in flags]
    visible = "--visible" in sys.argv
    deep_audit = "--deep-audit" in sys.argv

    if len(args) != 2:
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/population dir> [--visible] [--deep-audit]")
        sys.exit(1)

    xlsx_path = Path(args[0])
    cache_dir = Path(args[1])

    results = update_population_nrw(xlsx_path, cache_dir, visible=visible, deep_audit=deep_audit)
    for line in results:
        print(line)