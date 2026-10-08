"""
regional-indicators/transform/xlsx_update/update_population_erp_lga.py

Writes LGA-level estimated resident population (ERP) into the Population
sheet.

Input:  cache/population/lga_<slug>_population_erp_lga.json, produced
        by fetchers/fetch_population_erp_lga.py.
Target: the "LGA" section of the Population sheet, row "Population
        (ERP)" in each LGA's block.

The same row name exists in the SA2 section (written by
update_population_erp.py) under some of the same region names, so every
lookup passes section="LGA". The two writers read different cache files
(*_population_erp_lga.json and *_population_erp.json) and cannot pick up
each other's input.

Blocks are matched on the LGA name (the `lga` field of the towns in
towns.toml). The column B sub-label is not used because its wording
varies between workbooks; within a block the row name is unique.

The cache carries the full source series from 2001, so every write is
checked against it with audit_historical_series. The ABS revises recent
years with each release, so small differences from existing workbook
values are expected; they are reported and do not block the write. Only
the latest year is written; earlier years are not modified.

The workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_population_erp_lga.py <workbook.xlsx> <cache/population dir> [--visible]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

from base import write_one

INDICATOR_NAME = "Population (ERP)"
SHEET_NAME = "Population"
SECTION = "LGA"
CACHE_GLOB = "*_population_erp_lga.json"


def write_lga_erp(sheet, cache_files: list[Path]) -> tuple[list[str], int, int]:
    """Write every cached LGA figure into `sheet`.

    Separate from the Excel session handling so that it can be run
    against any object with the xlwings Sheet interface (see
    tests/fake_xlwings_sheet.py). Returns
    (result lines, written count, flagged count).
    """
    results: list[str] = []
    written_count = 0
    flagged_count = 0

    for path in cache_files:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)

        region = data.get("region")
        series_by_year = data.get("series_by_year") or {}
        latest_year = data.get("year")
        latest_value = data.get("value")
        label = f"{region} (LGA)"

        if not region or latest_year is None or latest_value is None:
            results.append(f"{path.name}: no data, skipped")
            continue

        try:
            report, coord = write_one(
                sheet, region, INDICATOR_NAME, None,
                int(latest_year), latest_value,
                source_series={int(y): v for y, v in series_by_year.items()},
                section=SECTION,
            )
        except ValueError as exc:
            results.append(f"{label}: SKIPPED (row-finding) — {exc}")
            flagged_count += 1
            continue

        if report.safe_to_write:
            results.append(f"{label}: WRITTEN {latest_year} = {latest_value:,} -> {coord}")
            written_count += 1
        else:
            results.append(f"{label}: FLAGGED, not written — {report.summary_line()}")
            flagged_count += 1

    return results, written_count, flagged_count


def update_population_erp_lga(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    cache_files = sorted(cache_dir.glob(CACHE_GLOB))
    if not cache_files:
        raise FileNotFoundError(
            f"No {CACHE_GLOB} files found in {cache_dir} -- "
            f"run fetch_population_erp_lga.py first "
            f"(python run_update.py --only population_erp_lga)."
        )

    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(str(xlsx_path))
        try:
            sheet = wb.sheets[SHEET_NAME]
            results, written_count, flagged_count = write_lga_erp(sheet, cache_files)
            if written_count:
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

    results = update_population_erp_lga(xlsx_path, cache_dir, visible=visible)
    for line in results:
        print(line)
