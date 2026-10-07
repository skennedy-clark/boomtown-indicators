"""
regional-indicators/transform/xlsx_update/update_business.py

Writes business counts by turnover band into the Business sheet.

Input:  cache/business/<region>_business.json, produced by
        fetchers/fetch_business.py.
Target: the Business sheet.

Sheet layout:
  Row 1  fiscal-year labels ("2008/09", ...) from column B.
  Section "Non-primary production", then section "Primary production".
  Each section lists the same regions in the same order. A region is a
  block of five rows: a name row, then the turnover bands "0k-50k",
  "50k-200k", "200k-2m" and "2m+".
  The name row holds a SUM formula over the four band rows for each
  year. Only band rows receive values; when a new year column is
  created the SUM formula is added to the name row for that column.

Year columns: to the right of the data, after a blank column, row 1
carries a second run of year-like labels above derived (growth-rate)
columns. The year finder therefore scans row 1 from column B and stops
at the first blank cell. It never matches a label in the derived area
and always places a new column directly after the last data column.

Every write goes through the cell and series audits in audit.py. A
series-level concern does not block the write: the value is written and
marked bold red for review. A cell-level block (a formula or unexpected
content) prevents the write.

The Toowoomba composite region is the sum of the Toowoomba - Central,
North Toowoomba - Harlaxton and Toowoomba - West SA2s, as defined in
fetch_business.py. Toowoomba - East is outside the study area.

The workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_business.py <workbook.xlsx> <cache/business dir> <year> [--visible]

<year> is the calendar year in which the financial year ends (2025 for
2024/25).
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).parent))
from audit import audit_cell, audit_series, WriteAuditReport, apply_write_formatting, describe_write_outcome
from base import _fiscal_label

SHEET_NAME = "Business"
FIRST_YEAR_COLUMN = 2   # column B
YEAR_HEADER_ROW = 1
REGION_BLOCK_SIZE = 5   # name row + 4 turnover-band rows

SECTION_HEADERS = {
    "NPP": "Non-primary production",
    "PP":  "Primary production",
}
BAND_LABELS = ["0k-50k", "50k-200k", "200k-2m", "2m+"]


def _find_year_column(sheet, year: int) -> int:
    """Return the column for `year` in row 1, appending one if needed.

    Scans from column B and stops at the first blank header cell, which
    marks the end of the data columns. A new column is created directly
    after the last data column.
    """
    used = sheet.used_range
    max_col = used.last_cell.column

    header_row = sheet.range((YEAR_HEADER_ROW, FIRST_YEAR_COLUMN), (YEAR_HEADER_ROW, max_col)).value
    if not isinstance(header_row, list):
        header_row = [header_row]

    last_real_col = None
    for offset, label in enumerate(header_row):
        col = FIRST_YEAR_COLUMN + offset
        if label is None:
            break  # end of the data columns
        if isinstance(label, str) and "/" in label:
            start_year = int(label.split("/")[0])
            end_year = start_year + 1
            if end_year == year:
                return col
            last_real_col = col

    if last_real_col is None:
        raise ValueError(
            f"Could not find any real fiscal-year labels in row {YEAR_HEADER_ROW} "
            f"of sheet '{sheet.name}' before the first gap -- can't determine "
            f"where to add a new year column safely."
        )

    new_col = last_real_col + 1
    year_cell = sheet.cells(YEAR_HEADER_ROW, new_col)
    year_cell.value = _fiscal_label(year)
    year_cell.number_format = "General"
    return new_col


def _find_section_row(sheet, section_label: str) -> int:
    used = sheet.used_range
    max_row = used.last_cell.row
    col_a = sheet.range((1, 1), (max_row, 1)).value
    if not isinstance(col_a, list):
        col_a = [col_a]

    matches = [1 + offset for offset, val in enumerate(col_a) if val == section_label]
    if len(matches) == 0:
        raise ValueError(f"Could not find section header '{section_label}' in sheet '{sheet.name}'.")
    if len(matches) > 1:
        raise ValueError(f"AMBIGUOUS: section '{section_label}' appears at rows {matches}.")
    return matches[0]


def _find_region_row(sheet, section_row: int, region_name: str, next_section_row: int | None) -> int:
    """Return the row of `region_name` within one section (from its heading
    to the next section heading or the end of the sheet).
    """
    used = sheet.used_range
    max_row = next_section_row - 1 if next_section_row else used.last_cell.row
    col_a = sheet.range((section_row + 1, 1), (max_row, 1)).value
    if not isinstance(col_a, list):
        col_a = [col_a]

    matches = [
        section_row + 1 + offset for offset, val in enumerate(col_a)
        if val == region_name
    ]
    if len(matches) == 0:
        raise ValueError(
            f"Could not find region '{region_name}' within the section starting "
            f"row {section_row} in sheet '{sheet.name}'. This function does not "
            f"guess or create a new region block for you."
        )
    if len(matches) > 1:
        raise ValueError(f"AMBIGUOUS: region '{region_name}' appears at rows {matches}.")
    return matches[0]


def _find_band_row(sheet, region_row: int, band_label: str) -> int:
    for row in range(region_row + 1, region_row + REGION_BLOCK_SIZE):
        if sheet.cells(row, 1).value == band_label:
            return row
    raise ValueError(
        f"Could not find turnover band '{band_label}' within rows "
        f"{region_row + 1}-{region_row + REGION_BLOCK_SIZE - 1} (region block "
        f"starting row {region_row}) in sheet '{sheet.name}'."
    )


def _col_letter(col: int) -> str:
    """Convert a 1-based column number to its letter: 1 -> 'A', 27 -> 'AA'."""
    letters = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _write_region_total_formula(sheet, region_row: int, col: int) -> str | None:
    """Add the region total formula for a new year column.

    The region name row holds =SUM(<first band>:<last band>) in every
    year column. A newly created column needs the same formula. It is
    written only if the cell is empty; an occupied cell is left
    unchanged. Returns the formula written, or None.
    """
    cell = sheet.cells(region_row, col)
    if cell.value is not None or (cell.formula and str(cell.formula).startswith("=")):
        return None  # occupied: leave unchanged

    col_letter = _col_letter(col)
    first_band_row = region_row + 1
    last_band_row = region_row + REGION_BLOCK_SIZE - 1
    formula = f"=SUM({col_letter}{first_band_row}:{col_letter}{last_band_row})"
    cell.formula = formula
    cell.number_format = "General"
    return formula


def _read_existing_series(sheet, row: int, exclude_col: int | None = None) -> dict:
    used = sheet.used_range
    max_col = used.last_cell.column
    header_row = sheet.range((YEAR_HEADER_ROW, FIRST_YEAR_COLUMN), (YEAR_HEADER_ROW, max_col)).value
    values = sheet.range((row, FIRST_YEAR_COLUMN), (row, max_col)).value
    if not isinstance(header_row, list):
        header_row, values = [header_row], [values]

    series = {}
    for offset, (label, val) in enumerate(zip(header_row, values)):
        col = FIRST_YEAR_COLUMN + offset
        if col == exclude_col:
            continue
        if label is None:
            break  # end of the data columns
        if isinstance(label, str) and "/" in label and isinstance(val, (int, float)):
            year = int(label.split("/")[0]) + 1
            series[year] = val
    return series


def _write_business_row(sheet, row: int, year: int, value):
    col = _find_year_column(sheet, year)
    cell = sheet.cells(row, col)

    cell_result = audit_cell(cell, value)
    existing_series = _read_existing_series(sheet, row, exclude_col=col)
    series_result = audit_series(existing_series, year, value)

    report = WriteAuditReport(cell_result, series_result)
    if report.safe_to_write or report.should_write_with_flag:
        cell.value = value
    apply_write_formatting(cell, report.safe_to_write, report.should_write_with_flag)

    return report, cell.address


def update_business(xlsx_path: Path, cache_dir: Path, year: int, visible: bool = False) -> list[str]:
    cache_files = sorted(cache_dir.glob("*_business.json"))
    if not cache_files:
        raise FileNotFoundError(
            f"No *_business.json files found in {cache_dir} -- run fetch_business.py first."
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

            npp_section_row = _find_section_row(sheet, SECTION_HEADERS["NPP"])
            pp_section_row = _find_section_row(sheet, SECTION_HEADERS["PP"])

            for path in cache_files:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)

                region = data["region"]
                indicators = data.get("indicators", {})

                for production, section_row, next_section in (
                    ("npp", npp_section_row, pp_section_row),
                    ("pp",  pp_section_row,  None),
                ):
                    band_values = indicators.get(production, {})
                    if not band_values:
                        results.append(f"{region} {production}: no data, skipped")
                        continue

                    try:
                        region_row = _find_region_row(sheet, section_row, region, next_section)
                    except ValueError as exc:
                        results.append(f"{region} {production}: SKIPPED (row-finding) — {exc}")
                        flagged_count += len(BAND_LABELS)
                        continue

                    for band_label in BAND_LABELS:
                        if band_label not in band_values:
                            results.append(f"{region} {production} {band_label}: no data, skipped")
                            continue
                        value = band_values[band_label]

                        try:
                            row = _find_band_row(sheet, region_row, band_label)
                            report, coord = _write_business_row(sheet, row, year, value)
                            line, is_written, is_flagged = describe_write_outcome(
                                f"{region} {production} {band_label}", report, coord, f"{year} = {value}"
                            )
                            results.append(line)
                            if is_written:
                                written_count += 1
                                any_written = True
                            if is_flagged:
                                flagged_count += 1
                        except ValueError as exc:
                            results.append(f"{region} {production} {band_label}: SKIPPED (row-finding) — {exc}")
                            flagged_count += 1

                    # Extend the region name row's SUM formula into the new year column
                    # (only if that cell is empty).
                    year_col = _find_year_column(sheet, year)
                    total_formula = _write_region_total_formula(sheet, region_row, year_col)
                    if total_formula:
                        results.append(f"{region} {production} TOTAL: formula extended -> {total_formula}")
                        any_written = True
                    else:
                        results.append(
                            f"{region} {production} TOTAL: cell not empty, left untouched "
                            f"-- check manually whether a total formula is needed here"
                        )

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
    args = [a for a in sys.argv[1:] if a != "--visible"]
    visible = "--visible" in sys.argv

    if len(args) != 3:
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/business dir> <year> [--visible]")
        sys.exit(1)

    xlsx_path = Path(args[0])
    cache_dir = Path(args[1])
    year = int(args[2])

    results = update_business(xlsx_path, cache_dir, year, visible=visible)
    for line in results:
        print(line)