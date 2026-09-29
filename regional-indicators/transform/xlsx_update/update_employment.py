"""
regional-indicators/transform/xlsx_update/update_employment.py
-----------------------------------------------------------------------
Writes fetch_salm_unemployment.py's per-REGION cache files
(cache/unemployment/regions/{sa2|lga}_*.json) into the Employment sheet.

CONFIRMED REAL STRUCTURE (2026-09-28, direct inspection of the 2026 reference):
  Row 1: plain calendar-year integers from column C (2001 ... 2025) -- the same
    layout as the Crime sheet, so this script REUSES Crime's already-tested
    year-column / existing-series / write helpers rather than duplicating them
    (import from update_crime; both sheets must keep this layout in step).
  Row 2: fiscal labels for the early columns ("2000/01" ...) then the same
    plain integer as row 1 from 2011 on. When a NEW year column is created,
    row 2 is filled with the year too, to keep that pattern.
  Three sections, each introduced by a marker in column A:
      "LGA"   (row 3, a header row: A="LGA", B="Label")  -> 5 LGA rows
      "SA2"   (row 9)                                    -> 15 SA2 rows
      "State" (row 25)                                   -> NSW, Queensland (benchmark)
  Every data row has column B = "Smoothed Unemployment rate (%)" and column A =
    the region name. THE SAME NAME APPEARS IN TWO SECTIONS ("Goondiwindi" is
    both an LGA and an SA2), so a region row is only ever searched for WITHIN
    its own section's bounds -- same disambiguation as Business's NPP/PP.
  Values are plain constants (no formulas) in every SA2/LGA cell.

NOT IN SCOPE: the two State rows. SALM does not publish states; their source
(looks like ABS Labour Force Survey) is still to be confirmed. This script never
touches the State section.

Reuses audit.py's shared write-with-flag mechanism: a shape-based series concern
still writes the value (bold red), never silently withheld; only a genuine
cell-level block (formula / unexpected content) stops a write.

*** NOT YET TESTED against a live Excel instance. *** The row/section finding and
the per-region write logic are tested against a mock built to the real structure
and the real region files; the actual xlwings write hasn't run yet -- test on a
throwaway copy first.

Usage:
    python update_employment.py <xlsx> <cache/unemployment dir> [--visible]
"""

from __future__ import annotations

import json
import sys
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).parent))
from audit import describe_write_outcome
from update_crime import _find_year_column, _write_crime_row

SHEET_NAME = "Employment"
SECTION_MARKERS = {"LGA": "LGA", "SA2": "SA2", "State": "State"}
SECTION_ORDER = ["LGA", "SA2", "State"]
YEAR_HEADER_ROW = 1
FISCAL_LABEL_ROW = 2


def _column_a(sheet) -> list:
    max_row = sheet.used_range.last_cell.row
    col_a = sheet.range((1, 1), (max_row, 1)).value
    return col_a if isinstance(col_a, list) else [col_a]


def _find_section_rows(sheet) -> dict:
    """{ 'LGA': 3, 'SA2': 9, 'State': 25 } -- each marker must appear exactly once."""
    col_a = _column_a(sheet)
    found = {}
    for section, marker in SECTION_MARKERS.items():
        rows = [i + 1 for i, v in enumerate(col_a) if v == marker]
        if len(rows) != 1:
            raise ValueError(
                f"Expected exactly one '{marker}' section marker in column A of "
                f"'{sheet.name}', found rows {rows}."
            )
        found[section] = rows[0]
    return found


def _find_region_row(sheet, section_rows: dict, section: str, label: str) -> int:
    """Search ONLY within `section`'s rows (marker+1 up to the next marker)."""
    start = section_rows[section] + 1
    later = [r for r in section_rows.values() if r > section_rows[section]]
    end = (min(later) - 1) if later else sheet.used_range.last_cell.row
    col_a = sheet.range((start, 1), (end, 1)).value
    if not isinstance(col_a, list):
        col_a = [col_a]
    matches = [start + i for i, v in enumerate(col_a) if v == label]
    if not matches:
        raise ValueError(
            f"Could not find '{label}' in the {section} section (rows {start}-{end}) of "
            f"'{sheet.name}'. This function does not guess or create rows."
        )
    if len(matches) > 1:
        raise ValueError(f"AMBIGUOUS: '{label}' appears at rows {matches} in the {section} section.")
    return matches[0]


def _ensure_fiscal_label(sheet, year: int) -> None:
    """When a NEW year column was just created, row 2 is empty; later columns
    repeat the plain year there, so do the same. Never overwrites a label."""
    col = _find_year_column(sheet, year)
    cell = sheet.cells(FISCAL_LABEL_ROW, col)
    if cell.value is None:
        cell.value = year
        cell.number_format = "General"


def process_region(sheet, section_rows: dict, data: dict):
    """One region's latest complete year -> its row. Returns
    (result_line, counts_as_written, counts_as_flagged)."""
    section = data["section"]
    label = data["region"]
    tag = f"{section}:{label}"
    values = data["indicators"]["unemployment"]["values"]
    if not values:
        return f"{tag}: no data, skipped", False, False

    latest = max(values, key=int)
    year, value = int(latest), values[latest]
    source_series = {int(y): v for y, v in values.items()}

    # ROUNDING FIX 2026-09-29: found on the first live run against the real
    # 2026 exemplar -- every current-year cell on this sheet (all 21 tested,
    # flagged and clean alike) is stored at 1 decimal place, but this writer
    # was passing full precision (e.g. 2.525) into both the write and the
    # cell-conflict check. Result: 8 of 12 flags that run produced were pure
    # rounding artifacts (source rounds to exactly the stored value -- e.g.
    # Maranoa stored=2.2, source=2.225, round(source,1)=2.2) rather than real
    # discrepancies; only 4 were genuine (Roma, Queensland, and two others).
    # Historical cells are NOT rounded this way (many hold long decimals --
    # e.g. 3.966666666666667) -- this appears to be a human-entry convention
    # for the newest column specifically, not a sheet-wide rule, so only the
    # value being written this year is rounded; source_series (used for the
    # historical-series comparison against existing SAVED years) is left at
    # full precision, unchanged.
    value = float(Decimal(str(value)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))

    try:
        row = _find_region_row(sheet, section_rows, section, label)
    except ValueError as exc:
        return f"{tag}: SKIPPED (row-finding) — {exc}", False, True

    report, coord = _write_crime_row(sheet, row, year, value, source_series)
    if report.safe_to_write or report.should_write_with_flag:
        _ensure_fiscal_label(sheet, year)
    return describe_write_outcome(tag, report, coord, f"{year} = {value:.4g}")


def update_employment(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    region_files = sorted((cache_dir / "regions").glob("*.json"))
    if not region_files:
        raise FileNotFoundError(
            f"No region files in {cache_dir / 'regions'} -- run "
            f"`run_update.py --only salm_unemployment` first."
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
                line, is_written, is_flagged = process_region(sheet, section_rows, data)
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
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/unemployment dir> [--visible]")
        sys.exit(1)
    for line in update_employment(Path(args[0]), Path(args[1]), visible="--visible" in sys.argv):
        print(line)