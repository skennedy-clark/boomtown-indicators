"""
regional-indicators/transform/xlsx_update/update_rainfall.py
-----------------------------------------------------------------------
Writes fetch_bom_rainfall.py's cached output (total/summer/winter) into
the Exogenous sheet's Rainfall section.

DELIBERATELY DOES NOT REUSE base.py's _find_year_column/
_find_town_indicator_row -- confirmed real, structural differences from
the Population sheet (2026-09-21 inspection of the actual reference
workbook):
  - Single header row (row 1, calendar years only, no separate fiscal-
    year row) starting at COLUMN B, not column C.
  - The row label is the station's own name+number (e.g. "Harewood
    042078"), not a fixed indicator string shared across every town.
  - Confirmed real inconsistencies in the row labels themselves: Miles'
    label has a trailing space ("Miles Post Office 042023 "), Wandoan
    uses ASCII "->" where Goondiwindi uses a real arrow "->" character,
    Narrabri's sub-rows are plain "Summer"/"Winter" where every other
    town uses "Summer (Jan-Mar, Oct-Dec)"/"Winter (Apr-Sept)". Matching
    on exact label text would break on several real towns.
  - What IS confirmed consistent everywhere, including the messy cases:
    row POSITION. Station/total row, then Summer directly below, then
    Winter, then Historic Average, in that exact order, every single
    town checked (Chinchilla through Wandoan, including Moranbah's
    currently-empty block). This script matches by searching for the
    station NUMBER as a substring of column A (robust to the label
    inconsistencies above), then writes by fixed row offset.

DOES NOW WRITE "Historic Average" (added 2026-09-22, previously out of
scope). Confirmed via Steve reading BOM's own "Climate Averages" page
directly (a different, genuinely accessible BOM product from the
interactive portal that blocks automated access) that this row is
meant to be BOM's own official Mean Annual Rainfall figure for the
station, not a self-computed rolling average -- fetch_bom_rainfall.py
now fetches this automatically per station. If unavailable (not every
station has this page -- confirmed real, e.g. Chinchilla's station
genuinely doesn't have one), the existing value is left untouched and
a clear note is produced instead, per Steve's explicit instruction:
never blank it or guess, just flag that it needs a manual check.
Written differently from total/summer/winter: the SAME value goes
across every year column in the row (confirmed that's how this row is
actually used in the real workbook -- a flat constant for charting,
not a per-year series), with a 20%-difference sanity check against
whatever's already there before overwriting.

Reuses audit.py's audit_cell/audit_series/audit_historical_series/
WriteAuditReport directly (those are genuinely generic, not
Population-sheet-specific) -- just not base.py's row/column finders,
which are.

Every write gets the ground-truth historical audit, not just the
shape-based one -- fetch_bom_rainfall.py's cache already carries the
full year-by-year record for each town, so source_series is always
available here (unlike some indicators where only the latest year is
ever fetched).

*** NOT YET TESTED against a live Excel instance. *** Row/column
finding logic is tested against a mock built to exactly match the real
Exogenous sheet structure (including the label inconsistencies above),
but the actual write against the real workbook hasn't run yet -- test
against a throwaway copy first, same as everything else in this
project's history.

Usage:
    python update_rainfall.py <path-to-Indicators_Data-Charts.xlsx> <cache/rainfall dir> [--visible]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).parent))
from audit import audit_cell, audit_series, audit_historical_series, WriteAuditReport

SHEET_NAME = "Exogenous"
SECTION_HEADER = "Rainfall"
FIRST_YEAR_COLUMN = 2   # column B -- confirmed different from Population's column C
YEAR_HEADER_ROW = 1     # single header row -- confirmed different from Population's row 2


def _find_year_column(sheet, year: int) -> int:
    """Exogenous-specific: single calendar-year header row starting at
    column B. Creates a new column (no separate fiscal-year row to
    also fill in, unlike Population) if the year isn't there yet.

    Deliberately does NOT use sheet.used_range.last_cell.column to
    decide where "the end of the year data" is -- confirmed real bug
    (2026-09-22): the Exogenous sheet's used_range extends well past
    the Rainfall section's actual last year column, because of
    unrelated content further down the sheet (Education/Fuel sections,
    or leftover formatting) that has nothing to do with row 1's year
    headers. Trusting used_range caused a real write to land in AA
    instead of the correct Z on a real test file. Searches row 1
    itself for the rightmost cell that actually contains a year,
    and appends immediately after that -- not vulnerable to whatever
    else is going on elsewhere in the sheet.
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


def _find_rainfall_station_row(sheet, station_number: str) -> int:
    """Search the Rainfall section (between the 'Rainfall' section
    header and the next section header) for a row whose column A
    CONTAINS station_number as a substring -- not an exact match,
    since real label text is confirmed inconsistent (trailing spaces,
    different arrow characters) but the station number itself is the
    one thing guaranteed present and unambiguous in every case.

    Returns the station/total row number. Summer is row+1, Winter is
    row+2, Historic Average is row+3 -- confirmed consistent by
    position across every town checked, including the inconsistently-
    labelled ones.
    """
    used = sheet.used_range
    max_row = used.last_cell.row

    col_a = sheet.range((1, 1), (max_row, 1)).value
    if not isinstance(col_a, list):
        col_a = [col_a]

    in_section = False
    matches = []
    for offset, val in enumerate(col_a):
        row = 1 + offset
        text = str(val).strip() if val is not None else ""

        if text == SECTION_HEADER:
            in_section = True
            continue
        if in_section and val is not None and text and text != SECTION_HEADER:
            # A bare town-name header row (no station number in it) --
            # or a genuine data row. Only station rows matter here.
            if in_section and station_number in text:
                matches.append(row)
            # Detect leaving the Rainfall section: a short non-town-like
            # label that doesn't contain digits and isn't a known
            # sub-row label is treated as the next section's header.
            if text in ("Education", "Fuel", "Business", "Crime",
                        "Employment", "Housing", "Income", "Population"):
                in_section = False

    if len(matches) == 1:
        return matches[0]
    if len(matches) == 0:
        raise ValueError(
            f"Could not find a Rainfall row containing station number "
            f"'{station_number}' in sheet '{sheet.name}'. Check the "
            f"number is right and the row actually exists -- this "
            f"function does not guess or create a new row."
        )
    raise ValueError(
        f"AMBIGUOUS: {len(matches)} rows in the Rainfall section contain "
        f"'{station_number}' (rows {matches}). Check for a genuine "
        f"duplicate or a substring collision (e.g. one station number "
        f"contained within another)."
    )


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
    """Fallback for when a fresh BOM fetch isn't available: per Steve's
    explicit instruction, "use the previous average" means actually
    writing the existing value into this year's new column too -- not
    leaving it blank. Confirmed real bug (2026-09-22): an earlier
    version did nothing at all when the fresh fetch failed, which left
    a genuine gap at the newly-created year column even though every
    other column in the row had real data, because appending a new
    year column always starts blank and nothing was filling it back in.

    Reads the existing value from the row (the row is meant to be a
    flat constant, so any already-populated cell works) and writes it
    into year's column specifically, leaving every other column as-is.
    Returns (written, message); written=False (nothing to carry
    forward) only when the row has no existing value anywhere, i.e.
    a town with no history at all yet, like Moranbah currently.
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
    """Write new_value across EVERY year column in the Historic Average
    row -- confirmed this row is a flat constant repeated across the
    whole row in the real workbook (for charting convenience), not a
    per-year value like total/summer/winter, so it needs a different
    write pattern: same value everywhere, not one cell at a time.

    Runs one sanity check before overwriting: compares new_value
    against whatever's CURRENTLY in the row (read from any populated
    cell, since they're all meant to already be identical) and flags
    rather than silently overwrites if the two disagree by more than
    20% -- catches a genuinely wrong new value (bad fetch, wrong
    station) without needing a full per-cell audit for a row that's
    supposed to be uniform anyway.

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

    for offset in range(max_col - FIRST_YEAR_COLUMN + 1):
        col = FIRST_YEAR_COLUMN + offset
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


def update_rainfall(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    cache_files = sorted(cache_dir.glob("*_bom_rainfall.json"))
    if not cache_files:
        raise FileNotFoundError(
            f"No *_bom_rainfall.json files found in {cache_dir} -- "
            f"run fetch_bom_rainfall.py first."
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
                    station_row = _find_rainfall_station_row(sheet, station)
                except ValueError as exc:
                    results.append(f"{town} (station {station}): SKIPPED (row-finding) — {exc}")
                    flagged_count += 3  # would have been 3 writes (total/summer/winter)
                    continue

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
                            any_written = True
                        else:
                            results.append(
                                f"{town} {label}: FLAGGED, not written — {report.summary_line()}"
                            )
                            flagged_count += 1
                    except ValueError as exc:
                        results.append(f"{town} {label}: SKIPPED (row-finding) — {exc}")
                        flagged_count += 1

                # Historic Average: a different write pattern (whole row,
                # not one year) and a different source (BOM's official
                # Climate Averages page, not SILO/manual monthly data) --
                # handled separately from the total/summer/winter loop above.
                historic_row = station_row + 3
                official_avg = data.get("bom_official_historic_avg_mm")
                official_note = data.get("bom_official_historic_avg_note", "")

                if official_avg is not None:
                    try:
                        written, message = _write_historic_average_row(sheet, historic_row, official_avg)
                        results.append(f"{town} historic average: {message}")
                        if written:
                            written_count += 1
                            any_written = True
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
                        any_written = True
                    else:
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
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/rainfall dir> [--visible]")
        sys.exit(1)

    xlsx_path = Path(args[0])
    cache_dir = Path(args[1])

    results = update_rainfall(xlsx_path, cache_dir, visible=visible)
    for line in results:
        print(line)