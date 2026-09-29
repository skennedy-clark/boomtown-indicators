"""
regional-indicators/transform/xlsx_update/update_housing.py
-----------------------------------------------------------------------
Writes fetch_qgso_housing.py's per-REGION cache files
(cache/housing/regions/{lga|sa2|state}_*.json) into the Housing sheet.

CONFIRMED REAL STRUCTURE (2026-09-29, direct inspection of the 2026
reference): same three-section shape as the Employment sheet -- "LGA"
marker (row 4), "SA2" marker (row 41), "State" marker (row 120) -- so this
script reuses update_employment.py's section-finding helpers directly
rather than duplicating them (SECTION_MARKERS happens to be identical:
LGA/SA2/State). Year header is plain calendar-year integers in row 1 from
column C, SAME layout as Crime/Employment, but starting at 2000 (not 2001)
-- update_crime.py's _find_year_column searches by VALUE, not a fixed
offset, so this works unchanged.

Each region is a 6-row block: 1 name row + up to 5 indicator rows, in this
fixed order where present:
  "Building Approvals: Non-residential dwelling units (Private) Total (Number)"
  "Detached dwelling: median sale price ($)"
  "Detached dwelling: number of sales (Number)"
  "House - 3 bedrooms - median rent of lodgements ($/week)"
  "Building Approvals: Residential dwelling units (Private) New Houses (Number)"
The State section's only region (Queensland) has just 3 of these (no
approvals rows at all) -- confirmed real, not a gap in this script.
THE SAME REGION NAME IS USED IN TWO SECTIONS ("Goondiwindi" is both an LGA
and an SA2 row) -- same disambiguation as Employment: every row-find is
bounded to one section via update_employment.py's _find_region_row.

NOT YET COVERED: both approvals rows, on every block. fetch_qgso_housing.py
was extended 2026-09-29 to fetch region-level sales/price/rent at LGA, SA2,
and State level, but building approvals could not be made to return any
regions at LGA level in live testing (worse than previously documented,
which assumed at least LGA worked) -- see that file's _fetch_regions()
docstring. Still open; this script simply has no data for those two rows
and never writes them.
ALSO NOT COVERED: the separate Narrabri (NSW) block starting row 131,
which uses a completely different nested category structure and needs its
own fetcher entirely -- out of scope here, same as BOCSAR was initially
separated from QPS for Crime.

Reuses audit.py's shared write-with-flag mechanism via update_crime.py's
_write_crime_row -- a shape-based series concern still writes the value
(bold red), never silently withheld; only a genuine cell-level block stops
a write.

*** NOT YET TESTED against a live Excel instance. *** The fetcher's
region-level output was live-tested and three real bugs were found and
fixed there (a dropped state-region parser case, a label-collision dict
merge that silently dropped LGA-level Goondiwindi, and a missing SA2/
prefix that dropped every SA2 region) -- but the actual xlwings write
against this sheet's real row/column layout has not run yet. Test on a
throwaway copy first.

Usage:
    python update_housing.py <xlsx> <cache/housing dir> [--visible]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).parent))
from audit import describe_write_outcome
from update_crime import _find_year_column, _write_crime_row
from update_employment import _find_section_rows, _find_region_row

SHEET_NAME = "Housing"
BLOCK_SIZE = 6  # 1 name row + up to 5 indicator rows


def _find_housing_indicator_row(sheet, block_start_row: int, label: str) -> int:
    for row in range(block_start_row + 1, block_start_row + BLOCK_SIZE):
        if sheet.cells(row, 1).value == label:
            return row
    raise ValueError(
        f"Could not find '{label}' within rows {block_start_row + 1}-"
        f"{block_start_row + BLOCK_SIZE - 1} (block starting row {block_start_row}) "
        f"in sheet '{sheet.name}'. Expected for the State section's approvals rows "
        f"(genuinely absent there) -- unexpected for anything else."
    )


def process_region(sheet, section_rows: dict, data: dict):
    """One region's every indicator, latest complete year -> its row.
    Returns a list of (result_line, is_written, is_flagged) tuples, one
    per indicator actually present in this region's cache file."""
    section = data["section"]
    label = data["region"]
    results = []

    try:
        block_row = _find_region_row(sheet, section_rows, section, label)
    except ValueError as exc:
        return [(f"{section}:{label}: SKIPPED (row-finding) — {exc}", False, True)]

    for key, entry in data["indicators"].items():
        tag = f"{section}:{label} [{key}]"
        row_label = entry.get("label")
        values = entry.get("values", {})
        if not row_label or not values:
            results.append((f"{tag}: no data, skipped", False, False))
            continue

        latest = max(values, key=int)
        year, value = int(latest), values[latest]
        source_series = {int(y): v for y, v in values.items()}

        try:
            row = _find_housing_indicator_row(sheet, block_row, row_label)
        except ValueError as exc:
            results.append((f"{tag}: SKIPPED (row-finding) — {exc}", False, True))
            continue

        report, coord = _write_crime_row(sheet, row, year, value, source_series)
        results.append(describe_write_outcome(tag, report, coord, f"{year} = {value:,.4g}"))

    return results


def update_housing(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    region_files = sorted((cache_dir / "regions").glob("*.json"))
    if not region_files:
        raise FileNotFoundError(
            f"No region files in {cache_dir / 'regions'} -- run "
            f"`run_update.py --only qgso_housing` first."
        )

    results, written, flagged = [], 0, 0
    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(str(xlsx_path))
        try:
            sheet = wb.sheets[SHEET_NAME]
            section_rows = _find_section_rows(sheet)
            any_written = False
            for path in region_files:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)
                for line, is_written, is_flagged in process_region(sheet, section_rows, data):
                    results.append(line)
                    written += is_written
                    flagged += is_flagged
                    any_written |= is_written
            if any_written:
                wb.save()
        finally:
            wb.close()
    finally:
        app.quit()

    results += ["", f"Summary: {written} written, {flagged} flagged for review."]
    return results


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--visible"]
    if len(args) != 2:
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/housing dir> [--visible]")
        sys.exit(1)
    for line in update_housing(Path(args[0]), Path(args[1]), visible="--visible" in sys.argv):
        print(line)