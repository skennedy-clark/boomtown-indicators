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
    block LAYOUT. A town header row, then the station/total row, then
    Summer, Winter and Historic Average directly below, in that order,
    for every town (Chinchilla through Wandoan, including Moranbah's
    near-empty block).

ROW FINDING REWRITTEN 2026-10-06 -- by TOWN header, not station number.
The first version searched column A for the station number. Those
numbers exist only in the hand-built 2026 reference file's labels; the
file this actually runs on each year (last year's delivered workbook)
has plain station names, so against the real 2025 working file it
found no rows and wrote nothing (0 written, every town skipped).
Now: find the block by its town header row and check the Summer /
Winter / Historic Average labels really are in the three rows below the
station row. The station number is NOT used to find anything: a station
can close and a town move to another (the sheet already records two
such changes, Goondiwindi and Wandoan, whose labels keep the original
station's number), so towns.toml alone says which station is current.
If a label's number differs from towns.toml the run prints a one-line
note and carries on. Works on both the 2025 and 2026 label styles.
Towns with no block on the sheet (Toowoomba's three sub-areas,
Shepparton, Yarram) are listed once at the end as expected, not
reported as failures -- which also stops Toowoomba's rows being
written four times over via its sub-areas' shared station.

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

Tested 2026-10-06 against the real 2025 working file's cell contents
and the 2026 reference file's, through tests/fake_xlwings_sheet.py (a
stand-in for the Excel sheet object -- this script cannot run for real
outside Windows/Mac Excel). First real Excel run still to be confirmed.

Usage:
    python update_rainfall.py <path-to-Indicators_Data-Charts.xlsx> <cache/rainfall dir> [--visible]
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


class NoRainfallBlock(Exception):
    """This town has no block in the Rainfall section at all. Not an
    error -- several towns.toml entries (Toowoomba's three sub-areas,
    Shepparton, Yarram) are fetched but were never on this sheet."""


SUB_ROW_LABEL_PREFIXES = ("Summer", "Winter", "Historic Average")
NEXT_SECTION_HEADERS = ("Education", "Fuel", "Business", "Crime",
                        "Employment", "Housing", "Income", "Population")


def _find_rainfall_station_row(sheet, town: str, station_number: str) -> int:
    """Find `town`'s block in the Rainfall section and return its
    station/total row. Summer is row+1, Winter row+2, Historic Average
    row+3.

    REWRITTEN 2026-10-06. The first version searched column A for the
    station NUMBER ("Harewood 042078"). Those numbers were only added
    to the labels in the hand-built 2026 reference file; the real
    starting file each year is last year's delivered workbook, whose
    labels are just the station name ("Harewood", "Dalby Airport").
    Confirmed against the 2025 working file: not one station number in
    column A, so every town was SKIPPED and nothing was written at all.

    What IS the same in both files is the block layout:
        <town name>              <- header row, column A only
        <station label>          <- total   (returned row)
        Summer ...               <- row+1
        Winter ...               <- row+2
        Historic Average         <- row+3
    so the block is found by its TOWN header row (compared with
    surrounding whitespace stripped -- "Chinchilla " has a trailing
    space in the real sheet), and the three sub-row labels are then
    CHECKED rather than assumed, so a shifted or malformed block stops
    here instead of writing into the wrong rows.

    The station number plays NO part in finding the row. A station can
    close and a town move to another one (Steve, 2026-10-06) -- the
    sheet already shows this: "New Kildonan 041507/ WTP (2020->)" and
    "Gililgulgul 035029/ TM (2021->)" keep the ORIGINAL station's number
    in the label after the change. So towns.toml is the one place that
    says which station is current, the town header is the one thing
    that identifies the block, and a label number that differs from
    towns.toml is not an error. See _station_label_note for the
    advisory line printed in that case.

    Raises NoRainfallBlock if the town has no block on the sheet, and
    ValueError for anything that needs a human to look.
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
    """Advisory only -- never blocks a write. If the station label in
    column A mentions a station number and none of the numbers in it is
    the one towns.toml gives for this town, return a one-line note
    saying so; otherwise "".

    Expected and harmless after a station change (the label keeps the
    old number, towns.toml holds the new one). Worth a glance in any
    other case, since it could also mean towns.toml points at the wrong
    station -- which is why it is reported rather than ignored.
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
    """Column numbers that actually hold a year in the header row. Used
    wherever a whole row is touched, so nothing is ever written under a
    column that isn't a year -- used_range can run past the last year
    (the same trap documented in _find_year_column)."""
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

    # CHANGED 2026-10-06: only columns that hold a year in row 1, not
    # every column out to used_range's edge (which can run past the last
    # year -- see _year_columns).
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
    """Audit and write every cached town's latest year into `sheet`.
    Split out from the Excel open/save handling (2026-10-06) so the
    row-finding and audit logic can be exercised against any sheet-like
    object (see tests/fake_xlwings_sheet.py). Returns
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
            flagged_count += 3  # would have been 3 writes (total/summer/winter)
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