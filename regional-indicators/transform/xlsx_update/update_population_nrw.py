"""
regional-indicators/transform/xlsx_update/update_population_nrw.py
---------------------------------------------------------------------
Writes the per-town UCL-level "Non-resident workers on-shift" figure
(sub_label "Non-resident workers (UCL)") from fetch_population_nrw.py's
cached output. Split out from what used to be a combined LGA+UCL
script -- see update_population_nrw_lga.py for the LGA-level sibling,
which writes a different section of the same sheet from the same
cache files.

Uses xlwings (real Excel via COM automation), NOT openpyxl -- openpyxl
was confirmed to corrupt this specific workbook. See base.py's
docstring for the full explanation.

Runs the pre-write audits from audit.py on every town before writing.
The UCL/FTE source file only ever gives the latest year (no history),
so these writes use the shape-based series audit only, NOT the
ground-truth historical cross-check (that requires full source
history, which update_population_nrw_lga.py has for its own writes,
but this script doesn't).

--deep-audit: for any flagged entry, cross-checks the corresponding
LGA's own history in the same year(s), using the SAME cache file's
lga_nrw_on_shift_by_year field (already fetched, no extra request
needed) -- a real regional workforce event should show up at both
levels; a UCL-only blip is more likely a genuine data-entry error
specific to that cell. This is context for a human to weigh, not a
verdict. Confirmed on real data: correctly shows Isaac LGA's +26%
swing in 2012 corroborating Moranbah's flagged UCL figure that year,
and a weaker -9% case for Moranbah's flagged 2016.

*** NOT YET TESTED against a live Excel instance for this specific
split -- the combined version this was split from WAS tested live and
worked correctly; this split preserves that logic unchanged, but
re-confirm after placing it. ***

Usage:
    python update_population_nrw.py <path-to-Indicators_Data-Charts.xlsx> <cache/population dir> [--visible] [--deep-audit]
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