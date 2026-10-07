"""
regional-indicators/transform/xlsx_update/update_rainfall.py

Writes annual and seasonal rainfall into the Exogenous sheet.

Input:  cache/rainfall/<slug>_bom_rainfall.json, produced by
        fetchers/fetch_bom_rainfall.py.
Target: the "Rainfall" section of the Exogenous sheet.

Sheet layout:
  Row 1  calendar years from column B (a single header row).
  Each town is a block of five rows:
      <town name>            heading, column A only
      <station label>        annual total
      Summer ...             October-March total
      Winter ...             April-September total
      Historic Average       long-term mean annual rainfall

Row finding: a block is located by its town heading, compared with
surrounding whitespace removed. The three labels below the station row
are then verified, so a block with a different layout is rejected
instead of being written to. Station labels are free text and are not
used for matching; their wording differs between towns and between
editions of the workbook.

Station numbers: the station a town uses is defined only by
`bom_station` in towns.toml. A station can close and be replaced, and
the sheet's label may keep the earlier station's number. A number in
the label that differs from towns.toml is reported as a note and does
not block the write.

Towns that are fetched but have no block on the sheet are listed once
at the end of the run.

Annual, summer and winter totals are written for the latest year only.
The cache carries the full series, so each write is checked against it
with audit_historical_series in addition to the cell and series audits.

Historic Average holds the Bureau of Meteorology's published mean
annual rainfall for the station and is constant across the row. When
the fetcher obtained a current figure it is written to every year
column, provided it is within 20% of the existing value; a larger
difference is reported and nothing is written. When no current figure
is available the existing value is carried into the new year column.

The workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_rainfall.py <workbook.xlsx> <cache/rainfall dir> [--visible]
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).parent))
from audit import audit_cell, audit_series, audit_historical_series, WriteAuditReport

SHEET_NAME = "Exogenous"
SECTION_HEADER = "Rainfall"
FIRST_YEAR_COLUMN = 2   # column B
YEAR_HEADER_ROW = 1     # single calendar-year header row


def _find_year_column(sheet, year: int) -> int:
    """Return the column for `year` in row 1, appending one if needed.

    The last year column is found by scanning row 1 for year values, not
    from the sheet's used range, which extends beyond the year columns
    because of other sections lower on the sheet. A new column is placed
    directly after the last year.
    """
    used = sheet.used_range
    max_col = used.last_cell.column

    header_row = sheet.range((YEAR_HEADER_ROW, FIRST_YEAR_COLUMN), (YEAR_HEADER_ROW, max_col)).value
    if not isinstance(header_row, list):
        header_row = [header_row]

    last_year_col = None
    for offset, cell_year in enumerate(header_row):
        col = FIRST_YEAR_COLUMN + offset
        if cell_year == year:
            return col
        if isinstance(cell_year, (int, float)) and 1990 <= cell_year <= 2100:
            last_year_col = col

    if last_year_col is None:
        raise ValueError(
            f"Could not find any year values in row {YEAR_HEADER_ROW} of "
            f"sheet '{sheet.name}' -- can't determine where to add a new "
            f"year column safely."
        )

    new_col = last_year_col + 1
    year_cell = sheet.cells(YEAR_HEADER_ROW, new_col)
    year_cell.value = year
    year_cell.number_format = "General"
    return new_col


class NoRainfallBlock(Exception):
    """Raised when a town has no block in the Rainfall section.

    Not an error: some configured towns are not presented on this sheet.
    """


SUB_ROW_LABEL_PREFIXES = ("Summer", "Winter", "Historic Average")
NEXT_SECTION_HEADERS = ("Education", "Fuel", "Business", "Crime",
                        "Employment", "Housing", "Income", "Population")


def _find_rainfall_station_row(sheet, town: str, station_number: str) -> int:
    """Return the station (annual total) row of `town`'s block.

    Summer is row+1, Winter row+2 and Historic Average row+3. The block
    is found by its town heading and the three labels below the station
    row are verified.

    `station_number` is not used to find the row; see
    _station_label_note.

    Raises NoRainfallBlock if the town has no block, and ValueError if
    the heading is duplicated or the block is not laid out as expected.
    """
    used = sheet.used_range
    max_row = used.last_cell.row

    col_a = sheet.range((1, 1), (max_row, 1)).value
    if not isinstance(col_a, list):
        col_a = [col_a]
    labels = [str(v).strip() if v is not None else "" for v in col_a]

    in_section = False
    header_rows = []
    for offset, text in enumerate(labels):
        row = 1 + offset
        if text == SECTION_HEADER:
            in_section = True
            continue
        if not in_section:
            continue
        if text in NEXT_SECTION_HEADERS:
            break
        if text == town.strip():
            header_rows.append(row)

    if not header_rows:
        raise NoRainfallBlock(town)
    if len(header_rows) > 1:
        raise ValueError(
            f"AMBIGUOUS: {len(header_rows)} rows in the Rainfall section are "
            f"headed '{town}' (rows {header_rows}). Check for a duplicated block."
        )

    station_row = header_rows[0] + 1

    def label_at(row: int) -> str:
        return labels[row - 1] if row - 1 < len(labels) else ""

    for offset, expected in enumerate(SUB_ROW_LABEL_PREFIXES, start=1):
        found = label_at(station_row + offset)
        if not found.startswith(expected):
            raise ValueError(
                f"'{town}' block starts at row {header_rows[0]}, but row "
                f"{station_row + offset} reads {found!r} where a label starting "
                f"'{expected}' was expected. The block isn't laid out as "
                f"station / Summer / Winter / Historic Average -- not writing "
                f"into it."
            )

    station_label = label_at(station_row)
    if not station_label:
        raise ValueError(
            f"'{town}' block at row {header_rows[0]} has no station label in "
            f"row {station_row}."
        )

    return station_row


def _station_label_note(sheet, station_row: int, station_number: str) -> str:
    """Return a note if the station label names a different station.

    If the label in column A contains a station number and none of the
    numbers in it matches `station_number`, a one-line note is returned;
    otherwise "". The note is advisory. A mismatch is expected after a
    station has been replaced, and can otherwise indicate an incorrect
    `bom_station` in towns.toml.
    """
    label = sheet.cells(station_row, 1).value
    label = str(label).strip() if label is not None else ""
    wanted = str(station_number).lstrip("0")
    numbers_in_label = [n.lstrip("0") for n in re.findall(r"\d{5,6}", label)]
    if numbers_in_label and wanted not in numbers_in_label:
        return (
            f"note: the sheet's label reads {label!r} but towns.toml uses "
            f"station {station_number} for this town. Fine if the station "
            f"was replaced -- consider updating the label; otherwise check "
            f"towns.toml."
        )
    return ""


def _year_columns(sheet) -> list[int]:
    """Return the columns whose header cell in row 1 holds a year.

    Used when a whole row is written, so that no column beyond the year
    columns is touched.
    """
    used = sheet.used_range
    max_col = used.last_cell.column
    header_row = sheet.range((YEAR_HEADER_ROW, FIRST_YEAR_COLUMN), (YEAR_HEADER_ROW, max_col)).value
    if not isinstance(header_row, list):
        header_row = [header_row]
    return [
        FIRST_YEAR_COLUMN + offset
        for offset, cell_year in enumerate(header_row)
        if isinstance(cell_year, (int, float)) and 1990 <= cell_year <= 2100
    ]


def _read_existing_series(sheet, row: int, exclude_col: int | None = None) -> dict:
    used = sheet.used_range
    max_col = used.last_cell.column
    years = sheet.range((YEAR_HEADER_ROW, FIRST_YEAR_COLUMN), (YEAR_HEADER_ROW, max_col)).value
    values = sheet.range((row, FIRST_YEAR_COLUMN), (row, max_col)).value
    if not isinstance(years, list):
        years, values = [years], [values]
    series = {}
    for offset, (yr, val) in enumerate(zip(years, values)):
        col = FIRST_YEAR_COLUMN + offset
        if col == exclude_col:
            continue
        if isinstance(yr, (int, float)) and isinstance(val, (int, float)):
            series[int(yr)] = val
    return series


def _carry_forward_historic_average(sheet, row: int, year: int) -> tuple[bool, str]:
    """Copy the existing Historic Average into `year`'s column.

    Used when no current figure is available from the source. The row
    is constant, so any populated cell supplies the value. Other columns
    are not changed.

    Returns (written, message). Nothing is written if the row has no
    existing value.
    """
    used = sheet.used_range
    max_col = used.last_cell.column
    existing_values = sheet.range((row, FIRST_YEAR_COLUMN), (row, max_col)).value
    if not isinstance(existing_values, list):
        existing_values = [existing_values]

    existing_numeric = [v for v in existing_values if isinstance(v, (int, float))]
    if not existing_numeric:
        return False, (
            "no previous average available either (row has no existing data) — "
            "needs manual entry, nothing to carry forward"
        )

    previous_value = existing_numeric[0]
    col = _find_year_column(sheet, year)
    cell = sheet.cells(row, col)
    cell.value = previous_value
    cell.number_format = "General"
    return True, f"used previous average ({previous_value}) for {year} -> row {row}"


def _write_historic_average_row(sheet, row: int, new_value: float) -> tuple[bool, str]:
    """Write `new_value` to every year column of the Historic Average row.

    The row holds one constant value. If the row already has a value and
    `new_value` differs from it by more than 20%, nothing is written and
    the difference is reported for review.

    Returns (written, message).
    """
    used = sheet.used_range
    max_col = used.last_cell.column
    existing_values = sheet.range((row, FIRST_YEAR_COLUMN), (row, max_col)).value
    if not isinstance(existing_values, list):
        existing_values = [existing_values]

    existing_numeric = [v for v in existing_values if isinstance(v, (int, float))]
    if existing_numeric:
        current = existing_numeric[0]
        if current != 0:
            pct_diff = abs(new_value - current) / abs(current)
            if pct_diff > 0.20:
                return False, (
                    f"FLAGGED, not written — new BOM official average ({new_value}) "
                    f"differs from the existing Historic Average ({current}) by "
                    f"{pct_diff:.0%}, more than the 20% sanity threshold. Could be a "
                    f"genuine correction (the whole point of fetching this) or a wrong "
                    f"station/bad fetch -- worth a human look before overwriting."
                )

    # Only columns that hold a year in row 1; see _year_columns.
    for col in _year_columns(sheet):
        cell = sheet.cells(row, col)
        cell.value = new_value
        cell.number_format = "General"

    return True, f"WRITTEN {new_value} across all year columns -> row {row}"


def _write_rainfall_row(sheet, row: int, year: int, value, source_series: dict | None):
    col = _find_year_column(sheet, year)
    cell = sheet.cells(row, col)

    cell_result = audit_cell(cell, value)
    existing_series = _read_existing_series(sheet, row, exclude_col=col)
    series_result = audit_series(existing_series, year, value)

    historical_result = None
    if source_series:
        historical_result = audit_historical_series(existing_series, source_series)

    report = WriteAuditReport(cell_result, series_result, historical_result)
    if report.safe_to_write:
        cell.value = value
        cell.number_format = "General"

    return report, cell.address


def write_rainfall(sheet, cache_files: list[Path]) -> tuple[list[str], int, int]:
    """Audit and write the latest year for every cached town.

    Separate from the Excel session handling so that it can be run
    against any object with the xlwings Sheet interface (see
    tests/fake_xlwings_sheet.py). Returns
    (result lines, written count, flagged count).
    """
    results: list[str] = []
    written_count = 0
    flagged_count = 0
    no_block: list[str] = []

    for path in cache_files:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)

        town = data["town"]
        station = str(data["bom_station"])
        indicators = data.get("indicators", {})
        total_by_year  = indicators.get("rainfall", {})
        summer_by_year = indicators.get("rainfall_summer", {})
        winter_by_year = indicators.get("rainfall_winter", {})

        if not total_by_year:
            results.append(f"{town}: no data, skipped")
            continue

        latest_year_str = max(total_by_year, key=int)
        latest_year = int(latest_year_str)

        try:
            station_row = _find_rainfall_station_row(sheet, town, station)
        except NoRainfallBlock:
            no_block.append(town)
            continue
        except ValueError as exc:
            results.append(f"{town} (station {station}): SKIPPED (row-finding) — {exc}")
            flagged_count += 3  # total, summer and winter
            continue

        label_note = _station_label_note(sheet, station_row, station)
        if label_note:
            results.append(f"{town}: {label_note}")

        summer_row = station_row + 1
        winter_row = station_row + 2

        for label, row, values_by_year in (
            ("total",  station_row, total_by_year),
            ("summer", summer_row,  summer_by_year),
            ("winter", winter_row,  winter_by_year),
        ):
            if latest_year_str not in values_by_year:
                results.append(f"{town} {label}: no {latest_year} data, skipped")
                continue

            value = values_by_year[latest_year_str]
            source_series = {int(y): v for y, v in values_by_year.items()}

            try:
                report, coord = _write_rainfall_row(sheet, row, latest_year, value, source_series)
                if report.safe_to_write:
                    results.append(
                        f"{town} {label}: WRITTEN {latest_year} = {value} -> {coord}"
                    )
                    written_count += 1
                else:
                    results.append(
                        f"{town} {label}: FLAGGED, not written — {report.summary_line()}"
                    )
                    flagged_count += 1
            except ValueError as exc:
                results.append(f"{town} {label}: SKIPPED (row-finding) — {exc}")
                flagged_count += 1

        # Historic Average is written to the whole row from a different
        # source (the Bureau's published station average), so it is handled
        # separately from the three annual figures above.
        historic_row = station_row + 3
        official_avg = data.get("bom_official_historic_avg_mm")
        official_note = data.get("bom_official_historic_avg_note", "")

        if official_avg is not None:
            try:
                written, message = _write_historic_average_row(sheet, historic_row, official_avg)
                results.append(f"{town} historic average: {message}")
                if written:
                    written_count += 1
                else:
                    flagged_count += 1
            except Exception as exc:
                results.append(f"{town} historic average: SKIPPED (row-finding) — {exc}")
                flagged_count += 1
        else:
            written, message = _carry_forward_historic_average(sheet, historic_row, latest_year)
            results.append(
                f"{town} historic average: {message} — fresh fetch unavailable: {official_note}"
            )
            if written:
                written_count += 1
            else:
                flagged_count += 1

    if no_block:
        results.append(
            f"No block on the {SHEET_NAME} sheet's Rainfall section, nothing to write "
            f"(expected): {', '.join(no_block)}"
        )

    return results, written_count, flagged_count


def update_rainfall(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    cache_files = sorted(cache_dir.glob("*_bom_rainfall.json"))
    if not cache_files:
        raise FileNotFoundError(
            f"No *_bom_rainfall.json files found in {cache_dir} -- "
            f"run fetch_bom_rainfall.py first."
        )

    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(str(xlsx_path))
        try:
            sheet = wb.sheets[SHEET_NAME]
            results, written_count, flagged_count = write_rainfall(sheet, cache_files)
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
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/rainfall dir> [--visible]")
        sys.exit(1)

    xlsx_path = Path(args[0])
    cache_dir = Path(args[1])

    results = update_rainfall(xlsx_path, cache_dir, visible=visible)
    for line in results:
        print(line)