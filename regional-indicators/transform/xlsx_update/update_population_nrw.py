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

The current source publishes only the two latest years at UCL level.
Earlier years come from the source register,
docs/non-resident-population-register.csv (built from the QGSO reports by
transform/non_resident_register.py): its verified UCL figures are the
town's history. With that history:
  - a town's row with no figures yet (a town just added) is filled with
    every verified year; years the reports do not publish stay blank;
  - existing figures are compared with the publications
    (audit_historical_series) instead of the shape-based outlier check;
  - a town with no current figure (Mitchell: grouped with other towns
    since 2018) still receives its history.

--deep-audit adds context to any flagged town: the corresponding LGA
series (already in the cache file) is examined for a change in the same
year. A genuine regional change normally appears at both levels. The
output informs review; it does not change what is written.

The workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_population_nrw.py <workbook.xlsx> <cache/population dir> [--visible] [--deep-audit]
"""

from __future__ import annotations

import csv
import json
import sys
from pathlib import Path

import xlwings as xw

from base import write_one, deep_audit_context, _find_town_indicator_row, read_existing_series

INDICATOR_NAME = "Non-resident workers on-shift"
SHEET_NAME = "Population"
UCL_SUB_LABEL = "Non-resident workers (UCL)"
REGISTER = Path(__file__).resolve().parents[3] / "docs" / "non-resident-population-register.csv"


def _row_is_empty(sheet, town: str) -> bool:
    """Whether the town's UCL row has no figures (False if not found)."""
    try:
        row = _find_town_indicator_row(sheet, town, INDICATOR_NAME, UCL_SUB_LABEL)
    except ValueError:
        return False
    return not read_existing_series(sheet, row)


def register_history(path: Path = REGISTER) -> dict[str, dict[int, int]]:
    """{town: {year: non-resident workers on-shift}} from the verified UCL
    rows of the source register; {} if there is no register."""
    history: dict[str, dict[int, int]] = {}
    if not path.exists():
        return history
    with open(path, encoding="utf-8", newline="") as f:
        for row in csv.DictReader(f):
            if (row["geography_level"] == "UCL" and row["verification_status"] == "verified"
                    and row["measure"] == "non-resident workers on-shift" and row["value"]):
                history.setdefault(row["geography"], {})[int(row["reference_year"])] = int(float(row["value"]))
    return history


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
            histories = register_history()

            for path in cache_files:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)

                town = data["town"]
                lga = data.get("lga")
                lga_series = data.get("lga_nrw_on_shift_by_year")
                ucl_latest = data.get("ucl_nrw_latest")

                history = dict(histories.get(town, {}))
                if not ucl_latest and history and _row_is_empty(sheet, town):
                    # No current figure, and a row with no figures yet (a town
                    # just added): the latest published year carries the history.
                    year = max(history)
                    ucl_latest = {"year": year, "value": history[year]}
                if not ucl_latest:
                    results.append(f"{town} (UCL): no data, skipped")
                    continue
                if history:
                    history[int(ucl_latest["year"])] = ucl_latest["value"]

                try:
                    report, coord = write_one(
                        sheet, town, INDICATOR_NAME, UCL_SUB_LABEL,
                        ucl_latest["year"], ucl_latest["value"],
                        source_series=history or None,
                    )
                    if report.safe_to_write:
                        filled = getattr(report, "history_filled", None)
                        results.append(
                            f"{town} (UCL): WRITTEN {ucl_latest['year']} = "
                            f"{ucl_latest['value']:,} -> {coord}"
                            + (f" (new row: history {min(filled)}-{max(filled)} filled, {len(filled)} years)"
                               if filled else "")
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