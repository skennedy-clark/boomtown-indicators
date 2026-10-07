"""
regional-indicators/transform/xlsx_update/update_housing.py

Writes housing indicators into the Housing sheet.

Input:  cache/housing/regions/*.json -- one file per region from
        fetch_qgso_housing.py (Queensland LGA, SA2 and state), plus
        narrabri_approvals.json and narrabri_sales_rent.json from
        fetch_narrabri_approvals.py and fetch_narrabri_sales_rent.py.
Target: the Housing sheet.

Sheet layout:
  Row 1  calendar years from column C.
  Sections "LGA", "SA2" and "State", each introduced by its name in
  column A. Each region is a block of a name row and up to five
  indicator rows:
    Building Approvals: Non-residential dwelling units (Private) Total (Number)
    Detached dwelling: median sale price ($)
    Detached dwelling: number of sales (Number)
    House - 3 bedrooms - median rent of lodgements ($/week)
    Building Approvals: Residential dwelling units (Private) New Houses (Number)
  The State section's Queensland block has the sales and rent rows only.
  A region name can occur in more than one section, so rows are searched
  for only within the bounds of their own section.

  Narrabri (NSW) has a separate block, headed "Narrabri (LGA)", with its
  own structure and its own sources (ABS building approvals; NSW
  Department of Communities and Justice Rent and Sales Report). It is
  handled by process_narrabri(). Its rows are located by their column B
  labels, except the all-dwellings rent row, which has no label and is
  taken as the row below "House 3-Bed". Narrabri values are displayed as
  whole numbers.

Section finding is shared with update_employment.py and the year-column
and write helpers with update_crime.py; all three sheets have the same
header layout.

A series-level audit concern does not block the write: the value is
written and marked bold red for review. A cell-level block (a formula
or unexpected content) prevents the write.

The workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_housing.py <workbook.xlsx> <cache/housing dir> [--visible]
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
BLOCK_SIZE = 6  # name row + up to 5 indicator rows


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


# ── Narrabri block ───────────────────────────────────────────────────────

NARRABRI_NUMBER_FORMAT = "0"   # dollar amounts and counts: display as whole numbers

# Cache key -> the column B label that identifies its row within the block.
NARRABRI_ROW_LABELS = {
    "mean_sales_price":         "Mean Sale Price",
    "median_sales_price":       "Median Sale Price",
    "rent_house_3bed":          "House 3-Bed",
    "sales_no":                 "Sales No.",
    "new_residential_building": "New Residential Building",
    "new_houses":                "New Houses",
    # "rent_total_3bed" has no column B label; see _find_narrabri_data_row.
}


def _find_narrabri_block_row(sheet) -> int:
    """Return the row of the "Narrabri (LGA)" block heading, found by its
    column A text so that inserted rows elsewhere do not matter.
    """
    used = sheet.used_range
    last_row = used.last_cell.row
    for row in range(1, last_row + 1):
        if sheet.cells(row, 1).value == "Narrabri (LGA)":
            return row
    raise ValueError(f"Could not find 'Narrabri (LGA)' in column A of sheet '{sheet.name}'.")


def _find_narrabri_data_row(sheet, block_row: int, key: str) -> int:
    """Return the row for one Narrabri indicator.

    Rows are found by their column B label (NARRABRI_ROW_LABELS), except
    "rent_total_3bed", which is the row immediately below "House 3-Bed".
    The search covers a fixed window below the block heading.
    """
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
    """Write Narrabri's indicators from its two cache files into its block."""
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
                # Whole-number display format, set after _write_crime_row because
                # that function's formatting step resets the number format.
                col = _find_year_column(sheet, year)
                sheet.cells(row, col).number_format = NARRABRI_NUMBER_FORMAT
            results.append((line, is_written, is_flagged))

    if not found_any_file:
        results.append(("Narrabri: SKIPPED — neither cache file found, run narrabri_approvals and narrabri_sales_rent first", False, True))

    return results


def process_region(sheet, section_rows: dict, data: dict):
    """Write every indicator in one region's cache file for its latest
    complete year.

    Returns a list of (result_line, is_written, is_flagged), one per
    indicator.
    """
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

            # Narrabri's block is not part of the LGA/SA2/State sections and is
            # written separately; its two cache files are excluded from the
            # per-region loop above.
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
