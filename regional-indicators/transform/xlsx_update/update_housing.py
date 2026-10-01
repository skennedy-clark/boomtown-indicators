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

NOT YET COVERED (before 2026-10-01): both approvals rows, on every block.
fetch_qgso_housing.py was extended 2026-09-29 to fetch region-level sales/
price/rent at LGA, SA2, and State level, but building approvals could not
be made to return any regions at LGA level in live testing -- FIXED
2026-09-30: the root cause was a stray "concorded": "Y" setting, not a
real limitation -- see fetch_qgso_housing.py's COLLECTIONS config. Both
approvals rows now populate on every LGA/SA2/State block.

NARRABRI (NSW) BLOCK, ADDED 2026-10-01: the separate block starting row
131 (confirmed real, direct inspection) has a completely different,
one-off nested structure -- a single named block, not a repeated LGA/SA2/
State pattern, so it's handled by its own process_narrabri() below rather
than forced through the generic section-based process_region(). Reads
TWO separate cache files (cache/housing/regions/narrabri_approvals.json
and narrabri_sales_rent.json, from fetch_narrabri_approvals.py and
fetch_narrabri_sales_rent.py respectively) since the two underlying
sources (ABS SDMX API for approvals, NSW DCJ Rent and Sales Report for
sales/rent) are fetched independently. Seven indicators total, each
located by its own distinctive column-B label within the block, EXCEPT
the "Total" 3-bedroom rent row, which has no column-B label of its own in
the real sheet -- confirmed live, found instead as the row immediately
below "House 3-Bed", which it always sits beside.

Per Steve's explicit request, every Narrabri write displays as a whole
number (number_format "0") -- these are dollar amounts and small counts,
not fractional rates like Crime, so unlike Crime's confirmed 1-decimal
convention, Narrabri doesn't need decimal precision shown. This is a
display-only format, same principle as Crime's "0.0" -- the underlying
stored value is unchanged, matching how Excel's own number formats work.

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


# ── Narrabri block (ADDED 2026-10-01) ────────────────────────────────────

NARRABRI_NUMBER_FORMAT = "0"   # whole numbers, per Steve's explicit request -- see module docstring

# key -> column-B label that uniquely identifies this row within the block.
# Confirmed by direct inspection of the real starting file, 2026-10-01.
NARRABRI_ROW_LABELS = {
    "mean_sales_price":         "Mean Sale Price",
    "median_sales_price":       "Median Sale Price",
    "rent_house_3bed":          "House 3-Bed",
    "sales_no":                 "Sales No.",
    "new_residential_building": "New Residential Building",
    "new_houses":                "New Houses",
    # "rent_total_3bed" deliberately absent here -- it has no column-B
    # label of its own in the real sheet; see _find_narrabri_data_row.
}


def _find_narrabri_block_row(sheet) -> int:
    """Locate the one-off "Narrabri (LGA)" block's name row by searching
    column A directly, rather than hardcoding row 131 -- robust against
    the block shifting if rows are added elsewhere in the sheet, same
    principle as every other row-finder in this project."""
    used = sheet.used_range
    last_row = used.last_cell.row
    for row in range(1, last_row + 1):
        if sheet.cells(row, 1).value == "Narrabri (LGA)":
            return row
    raise ValueError(f"Could not find 'Narrabri (LGA)' in column A of sheet '{sheet.name}'.")


def _find_narrabri_data_row(sheet, block_row: int, key: str) -> int:
    """Find one of the Narrabri block's 7 data rows. Six are found by
    their own distinctive column-B label (NARRABRI_ROW_LABELS). The
    seventh, "rent_total_3bed", has no column-B label at all in the real
    sheet -- confirmed live -- so it's found instead as the row
    immediately below "House 3-Bed", which it always sits beside there.
    Searches within a generous window below the block start (40 rows,
    comfortably past row 145, the last confirmed real row) rather than a
    tight fixed range, since this is a single one-off block, not a
    repeated per-region pattern with a known fixed size."""
    SEARCH_WINDOW = 40
    if key == "rent_total_3bed":
        house_row = _find_narrabri_data_row(sheet, block_row, "rent_house_3bed")
        return house_row + 1
    label = NARRABRI_ROW_LABELS[key]
    for row in range(block_row + 1, block_row + SEARCH_WINDOW):
        if sheet.cells(row, 2).value == label:
            return row
    raise ValueError(
        f"Could not find column-B label '{label}' (for '{key}') within "
        f"{SEARCH_WINDOW} rows of the Narrabri block (starting row {block_row}) "
        f"in sheet '{sheet.name}'."
    )


def process_narrabri(sheet, cache_dir: Path) -> list[tuple[str, bool, bool]]:
    """Narrabri's 7 indicators, read from its two separate cache files
    (approvals from ABS, sales/rent from NSW DCJ -- see module docstring),
    each written to its own row within the single Narrabri block."""
    results = []
    files = {
        "narrabri_approvals.json": "Building Approvals",
        "narrabri_sales_rent.json": "Sales/Rent",
    }
    found_any_file = False

    try:
        block_row = _find_narrabri_block_row(sheet)
    except ValueError as exc:
        return [(f"Narrabri: SKIPPED (block not found) — {exc}", False, True)]

    for filename, source_label in files.items():
        path = cache_dir / "regions" / filename
        if not path.exists():
            results.append((f"Narrabri ({source_label}): no cache file, skipped — run the matching fetcher first", False, False))
            continue
        found_any_file = True
        with open(path, encoding="utf-8") as f:
            data = json.load(f)

        for key, entry in data.get("indicators", {}).items():
            tag = f"Narrabri [{key}]"
            row_label = entry.get("label")
            values = entry.get("values", {})
            if not row_label or not values:
                results.append((f"{tag}: no data, skipped", False, False))
                continue

            latest = max(values, key=int)
            year, value = int(latest), values[latest]
            source_series = {int(y): v for y, v in values.items()}

            try:
                row = _find_narrabri_data_row(sheet, block_row, key)
            except ValueError as exc:
                results.append((f"{tag}: SKIPPED (row-finding) — {exc}", False, True))
                continue

            report, coord = _write_crime_row(sheet, row, year, value, source_series)
            line, is_written, is_flagged = describe_write_outcome(tag, report, coord, f"{year} = {value:,.4g}")
            if is_written or is_flagged:
                # Integer display format, per Steve's explicit request --
                # applied AFTER _write_crime_row, same reasoning as
                # update_crime.py's CRIME_NUMBER_FORMAT override: that
                # function's own apply_write_formatting() unconditionally
                # sets number_format based on flag status, so anything set
                # earlier would just be overwritten.
                col = _find_year_column(sheet, year)
                sheet.cells(row, col).number_format = NARRABRI_NUMBER_FORMAT
            results.append((line, is_written, is_flagged))

    if not found_any_file:
        results.append(("Narrabri: SKIPPED — neither cache file found, run narrabri_approvals and narrabri_sales_rent first", False, True))

    return results


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
    all_region_files = sorted((cache_dir / "regions").glob("*.json"))
    narrabri_filenames = {"narrabri_approvals.json", "narrabri_sales_rent.json"}
    region_files = [p for p in all_region_files if p.name not in narrabri_filenames]
    if not region_files and not any((cache_dir / "regions" / n).exists() for n in narrabri_filenames):
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

            # Narrabri is a single one-off block, not a repeated LGA/SA2/
            # State region -- handled once here, not through the generic
            # per-file loop above (its two cache files were excluded from
            # region_files precisely so they don't get force-fed through
            # process_region, which expects a recognised section).
            for line, is_written, is_flagged in process_narrabri(sheet, cache_dir):
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
