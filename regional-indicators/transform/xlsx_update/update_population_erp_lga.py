"""
regional-indicators/transform/xlsx_update/update_population_erp_lga.py
---------------------------------------------------------------------------
Writes the LGA-level "Population (ERP)" figure from
fetch_population_erp_lga.py's cached output into the Population sheet's
LGA section -- the "Residents (LGA)" row under Brisbane, Goondiwindi,
Isaac, Maranoa, Narrabri, Toowoomba and Western Downs
(Population!AA4, AA6, AA9, AA12, AA14, AA17, AA20 for 2025).

Sibling of update_population_erp.py, which writes the same indicator
name into the SA2 section. The two must not be confused:
  - this script   : section="LGA", reads *_population_erp_lga.json
  - the SA2 script: section="SA2", reads *_population_erp.json
The indicator name "Population (ERP)" is identical in both sections and
several block names are too (Goondiwindi and Toowoomba appear in LGA,
SA2 and UCL sections with very different values), so section="LGA" is
what keeps each figure in its own row. Requires the 2026-10-06 fix in
base.py's _find_town_indicator_row (the "LGA" header is in row 2, above
where the scan used to start).

Matches on the block name only (the JSON's "region", which is
towns.toml's [lga_regions.*] name). No sub-label is used: column B
reads "Residents (LGA)" in the 2025 working file but "Toowoomba
Residents (LGA)" for one block in the 2026 reference file -- the same
wording drift that caught update_population_nrw_lga.py out. Within a
block in the LGA section the indicator name alone is unique.

Full source history is available (2001 onwards), so every write goes
through audit.py's ground-truth historical cross-check. Note the ABS
revises recent years of ERP each release (rebasing after a Census), so
the existing workbook history typically differs from the source by a
fraction of a percent for the last several years. That is expected and
does not block the write; this script writes the NEW year only and
does not rewrite earlier years.

Uses xlwings (real Excel via COM automation), NOT openpyxl -- see
base.py's docstring for why.

Tested 2026-10-06 against the real 2025 working file's cell contents
through a stand-in for the Excel sheet object (this script cannot be
run for real outside Windows/Mac Excel): all seven rows found, all
seven values land in the right cells and equal the 2026 reference
file. First real Excel run still to be confirmed by Steve.

Usage:
    python update_population_erp_lga.py <path-to-Indicators_Data-Charts.xlsx> <cache/population dir> [--visible]
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
    """Write every cached LGA figure into `sheet`. Split out from the
    Excel open/save handling so the row-finding and audit logic can be
    exercised against any sheet-like object. Returns
    (result lines, written count, flagged count)."""
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
