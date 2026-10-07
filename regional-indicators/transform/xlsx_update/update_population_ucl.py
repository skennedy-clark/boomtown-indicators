"""
regional-indicators/transform/xlsx_update/update_population_ucl.py

Writes urban-centre (UCL) resident population into the Population sheet.

Input:  cache/population/<slug>_population_ucl.json, produced by
        fetchers/fetch_population_ucl.py.
Target: the "UCL" section of the Population sheet, row "Estimated
        resident population (a) by urban centre and locality" in each
        town's block. Only the latest year in each cache file is written.

Rows are matched on the indicator name alone. The column B sub-label is
not used because its wording varies between towns.

Each write is preceded by the cell and series audits in audit.py. A
flagged value is reported and not written; an exact entry in
verified_overrides.toml is not consulted by this writer.

All towns are processed in a single Excel session. The workbook is
edited through Excel (xlwings); see base.py.

Usage:
    python update_population_ucl.py <workbook.xlsx> <cache/population dir> [--visible]

--visible shows the Excel window while the script runs.
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
    """Write the latest UCL population for every cached town.

    Returns one result line per town (written, with its cell address, or
    flagged, with the reason) followed by a summary line.
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