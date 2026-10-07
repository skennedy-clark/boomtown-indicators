"""
regional-indicators/transform/xlsx_update/update_crime.py

Writes reported-offence rates into the Crime sheet.

Input:  cache/crime/<slug>_crime_<source>.json, one file per town or
        benchmark, produced by fetch_crime_qps.py (Queensland) and
        fetch_crime_bocsar.py (New South Wales).
Target: the Crime sheet.

Sheet layout:
  Row 1  calendar years from column C.
  Row 2  column headings ("QPS Division", "Chart label").
  Queensland towns and the Queensland benchmark: a 13-row block -- a
  name row followed by twelve offence-category rows, ending with "Total
  offences (person, property, other)". Column A is unique within a
  block and is used to find rows.
  Narrabri and the NSW benchmark: a separate block lower on the sheet
  with its own copy of the year header row (BOCSAR_YEAR_HEADER_ROW) and
  BOCSAR's offence categories. The aggregate rate sits on the name row
  itself, not on a separate total row.

The writer is source-agnostic. Each cache file carries, for every
indicator, the row label to write to and its values, so a fetcher for
another state needs no change here provided it writes the same schema.

A series-level audit concern does not block the write: the value is
written and marked bold red for review. A cell-level block (a formula
or unexpected content) prevents the write.

_find_year_column, _read_existing_series and _write_crime_row are also
used by the Employment and Housing writers, whose sheets share this
header layout.

The workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_crime.py <workbook.xlsx> <cache/crime dir> [--visible]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).parent))
from audit import audit_cell, audit_series, audit_historical_series, WriteAuditReport, apply_write_formatting, describe_write_outcome

SHEET_NAME = "Crime"
FIRST_YEAR_COLUMN = 3   # column C
YEAR_HEADER_ROW = 1
# The Narrabri/NSW block has its own copy of the year header row,
# immediately above it. It is kept in step with row 1 by
# _mirror_bocsar_year_header. This is specific to the Crime sheet and is
# deliberately not part of _find_year_column, which other sheets share.
BOCSAR_YEAR_HEADER_ROW = 162
TOWN_BLOCK_SIZE = 13    # name row + 12 offence-category rows

# Row labels are not defined here. Each fetcher's output carries the
# label for every indicator (indicators[key] = {"label": ..., "values":
# {...}}), so the writer handles any state's category set. See
# fetch_crime_qps.py and fetch_crime_bocsar.py for the schema.


def _find_year_column(sheet, year: int) -> int:
    """Return the column for `year` in row 1, appending one if needed.

    Scans from column C and stops at the first empty header cell, so
    that content further right on the sheet is never mistaken for a
    year column and a new column is always placed directly after the
    last year.
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
        # xlwings returns whole numbers as floats (2001.0), so both int and
        # float are accepted and compared as int.
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


def _mirror_bocsar_year_header(sheet, year: int) -> bool:
    """Copy the year in row 1 into the Narrabri/NSW block's header row.

    Idempotent: writes only when the block's header cell for that column
    is missing or different. Returns True if a cell was written, so the
    caller knows the workbook needs saving.
    """
    col = _find_year_column(sheet, year)  # the column already exists by this point; nothing is created here
    existing = sheet.cells(BOCSAR_YEAR_HEADER_ROW, col).value
    if existing != year:
        sheet.cells(BOCSAR_YEAR_HEADER_ROW, col).value = year
        sheet.cells(BOCSAR_YEAR_HEADER_ROW, col).number_format = "General"
        return True
    return False


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


# Display conventions for this sheet, applied after
# apply_write_formatting() (which sets number format and font from the
# audit result and would otherwise overwrite them): rates are shown to
# one decimal place, and the aggregate on a name row is bold and
# underlined.
CRIME_NUMBER_FORMAT = "0.0"
BOCSAR_TOTAL_ROWS = {163, 172}  # name rows that hold an aggregate rate


def _apply_crime_specific_formatting(cell, row: int) -> None:
    cell.number_format = CRIME_NUMBER_FORMAT
    if row in BOCSAR_TOTAL_ROWS:
        cell.font.bold = True
        # Font.underline is set through the same xlwings Font API as bold and
        # color.
        cell.font.underline = True


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
    if report.safe_to_write or report.should_write_with_flag:
        _apply_crime_specific_formatting(cell, row)

    return report, cell.address


def update_crime(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    # Any source's files are picked up, provided they are named
    # <town>_crime_<source>.json and follow the shared schema.
    cache_files = sorted(cache_dir.glob("*_crime_*.json"))
    if not cache_files:
        raise FileNotFoundError(
            f"No *_crime_*.json files found in {cache_dir} -- run a crime fetcher first."
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

            # Determine the year being written from the first cache file and
            # mirror its header into the Narrabri/NSW block once, before the main
            # loop.
            if cache_files:
                with open(cache_files[0], encoding="utf-8") as f:
                    _first_data = json.load(f)
                _first_indicators = _first_data.get("indicators", {})
                _all_years = {
                    int(y)
                    for entry in _first_indicators.values()
                    for y in entry.get("values", {})
                }
                if _all_years:
                    if _mirror_bocsar_year_header(sheet, max(_all_years)):
                        any_written = True

            for path in cache_files:
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)

                town = data["town"]
                indicators = data.get("indicators", {})

                try:
                    town_row = _find_town_row(sheet, town)
                except ValueError as exc:
                    results.append(f"{town}: SKIPPED (row-finding) — {exc}")
                    flagged_count += len(indicators)
                    continue

                # Some sources (BOCSAR) publish no separately labelled total. Their
                # aggregate rate belongs on the block's name row. This is detected
                # from the data -- no indicator labelled as a total -- not from the
                # state or source, so it applies to any fetcher whose output has that
                # shape.
                has_own_total_row = any(
                    "total" in (entry.get("label") or "").lower()
                    for entry in indicators.values()
                )
                town_total_series: dict[int, float] = {}

                # Row label and values come from the cache file.
                for key, entry in indicators.items():
                    label = entry.get("label")
                    values_by_year = entry.get("values", {})
                    if not label:
                        results.append(
                            f"{town} {key}: SKIPPED (no 'label' in fetcher output) — "
                            f"fetcher must supply indicators[key]['label']"
                        )
                        flagged_count += 1
                        continue
                    if not values_by_year:
                        results.append(f"{town} {key}: no data, skipped")
                        continue

                    if not has_own_total_row:
                        for y, v in values_by_year.items():
                            town_total_series[int(y)] = town_total_series.get(int(y), 0) + v

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

                if not has_own_total_row and town_total_series:
                    total_latest_year = max(town_total_series)
                    total_value = town_total_series[total_latest_year]
                    report, coord = _write_crime_row(
                        sheet, town_row, total_latest_year, total_value, town_total_series
                    )
                    line, is_written, is_flagged = describe_write_outcome(
                        f"{town} total (own name row)", report, coord,
                        f"{total_latest_year} = {total_value}"
                    )
                    results.append(line)
                    if is_written:
                        written_count += 1
                        any_written = True
                    if is_flagged:
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