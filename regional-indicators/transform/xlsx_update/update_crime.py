"""
regional-indicators/transform/xlsx_update/update_crime.py
-----------------------------------------------------------------------
Writes fetch_crime_qps.py's cached output into the Crime sheet, for the
11 confirmed QLD towns this fetcher covers.

NOT IN SCOPE FOR THIS SCRIPT, confirmed real and deliberately excluded,
not overlooked:
  - Narrabri / NSW: confirmed via direct inspection that Narrabri and
    the NSW benchmark use a COMPLETELY DIFFERENT 5-category structure
    (Assault, Malicious damage to property, Other offences, Other
    offences against the person, Robbery) from the 12-category QLD
    structure -- consistent with Narrabri needing a genuinely different
    data source (NSW BOCSAR, not QPS) that hasn't been built yet.
    fetch_crime_qps.py is already correctly QLD-only
    (SUPPORTED_STATES = ["QLD"]).
  - The Queensland state benchmark row (confirmed same 12-category
    structure as individual towns, but QPS's own data has no native
    statewide division to read directly -- would need its own
    aggregation across all QLD divisions, similar to Income's QLD/NSW
    benchmark work). Real, separate piece of work, not built here.

CONFIRMED REAL STRUCTURE (2026-09-24, direct inspection of the
reference workbook):
  Row 1: plain calendar-year integers (2001, 2002, ...) starting at
    column C -- NOT fiscal-year strings like Income, NOT the two-row
    fiscal+calendar convention like Population. Row 2: "QPS Division"/
    "Chart label" column headers, not data.
  Each town: a 13-row block -- name row, then 12 indicator rows in a
    FIXED order, confirmed identical across every QLD town checked
    (Chinchilla vs Dalby, row-by-row):
      Breach Domestic Violence Protection Order, Drug Offences,
      Good Order Offences, Offences Against Property,
      Offences Against the Person, Other Offences,
      Other Theft (excl. Unlawful Entry), Prostitution Offences,
      Traffic and Related Offences, Unlawful Entry,
      Weapons Act Offences, Total offences (person, property, other)
  Column B repeats column A for every row except Total, where it's a
    town-specific chart label ("{Town} total crime rate") -- not
    needed for row-finding, column A alone is unique within each
    town's block.

Reuses audit.py's shared write-with-flag mechanism (same as Income and
Business) -- a shape-based series concern still writes the value, just
visually flagged (bold red), never silently withheld; only a genuine
cell-level block (a formula, unexpected existing content) prevents a
write outright.

*** NOT YET TESTED against a live Excel instance. *** Row/column
finding logic is tested against a mock built to match the real
structure, but the actual write against the real workbook hasn't run
yet -- test against a throwaway copy first, same as every other wiring
script's first live test.

Usage:
    python update_crime.py <path-to-Indicators_Data-Charts.xlsx> <cache/crime dir> [--visible]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).parent))
from audit import audit_cell, audit_series, audit_historical_series, WriteAuditReport, apply_write_formatting, describe_write_outcome

SHEET_NAME = "Crime"
FIRST_YEAR_COLUMN = 3   # column C -- confirmed
YEAR_HEADER_ROW = 1
TOWN_BLOCK_SIZE = 13    # name row + 12 indicator rows, confirmed

# Confirmed real, fixed row order (see module docstring) -- matches
# fetch_crime_qps.py's INDICATOR_COLS keys plus "total".
INDICATOR_ROW_LABELS = {
    "breach_dv":         "Breach Domestic Violence Protection Order",
    "drug":              "Drug Offences",
    "good_order":        "Good Order Offences",
    "offences_property": "Offences Against Property",
    "offences_person":   "Offences Against the Person",
    "other_offences":    "Other Offences",
    "theft":             "Other Theft (excl. Unlawful Entry)",
    "prostitution":      "Prostitution Offences",
    "traffic":           "Traffic and Related Offences",
    "unlawful_entry":    "Unlawful Entry",
    "weapons":           "Weapons Act Offences",
    "total":             "Total offences (person, property, other)",
}


def _find_year_column(sheet, year: int) -> int:
    """Plain calendar-year integers in row 1, starting column C.
    Searches for a real gap the same defensive way update_rainfall.py's
    and update_business.py's finders do -- stop at the first None
    rather than trusting the sheet-wide used_range, and never scan past
    it when creating a new column.
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
            break
        # BUG FIXED 2026-09-24: xlwings returns whole-number cells as
        # floats (2001.0) via Excel's COM interface, confirmed real on
        # a live run -- isinstance(label, int) alone silently rejected
        # every real year value, since openpyxl (used only for
        # inspection, not the live write path) happens to preserve int
        # type but xlwings doesn't. Broadened to (int, float), compared
        # via int() conversion.
        if isinstance(label, (int, float)) and int(label) == year:
            return col
        if isinstance(label, (int, float)):
            last_real_col = col

    if last_real_col is None:
        raise ValueError(
            f"Could not find any year values in row {YEAR_HEADER_ROW} of "
            f"sheet '{sheet.name}' before the first gap."
        )

    new_col = last_real_col + 1
    year_cell = sheet.cells(YEAR_HEADER_ROW, new_col)
    year_cell.value = year
    year_cell.number_format = "General"
    return new_col


def _find_town_row(sheet, town_name: str) -> int:
    used = sheet.used_range
    max_row = used.last_cell.row
    col_a = sheet.range((1, 1), (max_row, 1)).value
    if not isinstance(col_a, list):
        col_a = [col_a]

    matches = [1 + offset for offset, val in enumerate(col_a) if val == town_name]
    if len(matches) == 0:
        raise ValueError(
            f"Could not find town '{town_name}' in sheet '{sheet.name}'. "
            f"This function does not guess or create a new block for you."
        )
    if len(matches) > 1:
        raise ValueError(f"AMBIGUOUS: '{town_name}' appears at rows {matches}.")
    return matches[0]


def _find_indicator_row(sheet, town_row: int, label: str) -> int:
    for row in range(town_row + 1, town_row + TOWN_BLOCK_SIZE):
        if sheet.cells(row, 1).value == label:
            return row
    raise ValueError(
        f"Could not find '{label}' within rows {town_row + 1}-"
        f"{town_row + TOWN_BLOCK_SIZE - 1} (town block starting row {town_row}) "
        f"in sheet '{sheet.name}'."
    )


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
            break
        if isinstance(label, (int, float)) and isinstance(val, (int, float)):
            series[int(label)] = val
    return series


def _write_crime_row(sheet, row: int, year: int, value, source_series: dict | None):
    col = _find_year_column(sheet, year)
    cell = sheet.cells(row, col)

    cell_result = audit_cell(cell, value)
    existing_series = _read_existing_series(sheet, row, exclude_col=col)
    series_result = audit_series(existing_series, year, value)

    historical_result = None
    if source_series:
        historical_result = audit_historical_series(existing_series, source_series)

    report = WriteAuditReport(cell_result, series_result, historical_result)
    if report.safe_to_write or report.should_write_with_flag:
        cell.value = value
    apply_write_formatting(cell, report.safe_to_write, report.should_write_with_flag)

    return report, cell.address


def update_crime(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    cache_files = sorted(cache_dir.glob("*_crime_qps.json"))
    if not cache_files:
        raise FileNotFoundError(
            f"No *_crime_qps.json files found in {cache_dir} -- run fetch_crime_qps.py first."
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
                indicators = data.get("indicators", {})

                try:
                    town_row = _find_town_row(sheet, town)
                except ValueError as exc:
                    results.append(f"{town}: SKIPPED (row-finding) — {exc}")
                    flagged_count += len(INDICATOR_ROW_LABELS)
                    continue

                for key, label in INDICATOR_ROW_LABELS.items():
                    values_by_year = indicators.get(key, {})
                    if not values_by_year:
                        results.append(f"{town} {key}: no data, skipped")
                        continue

                    latest_year_str = max(values_by_year, key=int)
                    latest_year = int(latest_year_str)
                    value = values_by_year[latest_year_str]
                    source_series = {int(y): v for y, v in values_by_year.items()}

                    try:
                        row = _find_indicator_row(sheet, town_row, label)
                        report, coord = _write_crime_row(sheet, row, latest_year, value, source_series)
                        line, is_written, is_flagged = describe_write_outcome(
                            f"{town} {key}", report, coord, f"{latest_year} = {value}"
                        )
                        results.append(line)
                        if is_written:
                            written_count += 1
                            any_written = True
                        if is_flagged:
                            flagged_count += 1
                    except ValueError as exc:
                        results.append(f"{town} {key}: SKIPPED (row-finding) — {exc}")
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
    args = [a for a in sys.argv[1:] if a != "--visible"]
    visible = "--visible" in sys.argv

    if len(args) != 2:
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/crime dir> [--visible]")
        sys.exit(1)

    xlsx_path = Path(args[0])
    cache_dir = Path(args[1])

    results = update_crime(xlsx_path, cache_dir, visible=visible)
    for line in results:
        print(line)