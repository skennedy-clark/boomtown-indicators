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
# BUG FOUND 2026-09-30, real run against the real 2025-origional file: row
# 162 is an EXACT mirror of row 1's year headers (2001-2024, same columns),
# sitting immediately above the Narrabri/BOCSAR block (which starts row
# 163). Confirmed by direct inspection, not assumed -- this is a
# section-local year header for the BOCSAR block specifically, distinct
# from the sheet-wide row 1 header the rest of Crime uses. Kept as a
# SEPARATE constant, and mirrored via its own small function below, rather
# than folded into _find_year_column -- that function is SHARED (Employment
# and Housing both import it), and row 162 in those sheets is an unrelated
# cell; this must never run for anything but Crime.
BOCSAR_YEAR_HEADER_ROW = 162
TOWN_BLOCK_SIZE = 13    # name row + 12 indicator rows, confirmed

# GENERALIZED 2026-09-24, per Steve's explicit direction ("knowing
# there will be a fetch victoria, tasmania, nt, wa at some point"):
# this used to be a hardcoded QLD-specific label dict. Now each
# fetcher's own JSON output carries its own label alongside every
# indicator (indicators[key] = {"label": ..., "values": {...}}), so
# this script is state-agnostic -- it works for QLD's 12 categories,
# NSW's different 8, and whatever future states turn out to use,
# without ever needing a per-state dict here. See
# fetch_crime_qps.py and fetch_crime_bocsar.py for the schema both
# fetchers (and any future state fetcher) must produce.


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


def _mirror_bocsar_year_header(sheet, year: int) -> bool:
    """Keep row 162 (the BOCSAR block's own local year-header row) in sync
    with row 1 whenever a year column exists/gets created. Idempotent --
    safe to call every run; only writes if row 162's copy is missing or
    wrong for that column. See BOCSAR_YEAR_HEADER_ROW's definition for why
    this exists and why it's Crime-specific, not part of the shared
    _find_year_column. Returns True if it actually wrote something, so the
    caller can fold this into any_written -- otherwise, on a hypothetical
    run where every other write gets flagged rather than written, this
    change would silently never reach wb.save()."""
    col = _find_year_column(sheet, year)  # never creates a second time; row 1 already has it by this point
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


# ADDED 2026-09-30, per Steve's explicit direction: Crime data displays to
# 1 decimal place, and the BOCSAR aggregate rows (163 Narrabri, 172 NSW --
# see BOCSAR_YEAR_HEADER_ROW above for why these two are special) are
# bold+underlined regardless of flag status, as a structural marker for
# "this is a name-row aggregate", not a review flag. Both are applied
# AFTER apply_write_formatting() below, never before -- that function
# unconditionally sets number_format/bold/color based on flag status, so
# anything set earlier would just get overwritten.
CRIME_NUMBER_FORMAT = "0.0"
BOCSAR_TOTAL_ROWS = {163, 172}  # Narrabri, NSW -- the name rows holding the aggregate, see the fix above


def _apply_crime_specific_formatting(cell, row: int) -> None:
    cell.number_format = CRIME_NUMBER_FORMAT
    if row in BOCSAR_TOTAL_ROWS:
        cell.font.bold = True
        # NOT YET CONFIRMED against real Excel -- .bold and .color are
        # already proven in this project's real runs, .underline is not.
        # If this errors, tell me the exact message; same fix pattern as
        # the earlier Font.color-cannot-be-None crash.
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
    # GENERALIZED 2026-09-24: was "*_crime_qps.json" only. Picks up
    # any state fetcher's output now (crime_qps for QLD, crime_bocsar
    # for NSW, and future crime_vic/crime_tas/crime_nt/crime_wa),
    # provided each writes its cache file as <town>_crime_<source>.json
    # in this same directory, in the shared schema.
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

            # Determine the year being written from the first cache file's
            # latest data, and mirror row 1's header into row 162 (the
            # BOCSAR block's own copy) BEFORE the main loop, once, rather
            # than repeating this per indicator write.
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

                # BUG FIXED 2026-09-30, found via a real run against
                # Narrabri/NSW: some sources (confirmed: BOCSAR) put their
                # aggregate total directly ON the town's own name row,
                # rather than as a separate labelled indicator row the way
                # QLD's QPS fetcher does ("Total offences (person,
                # property, other)" gets its own row below the town).
                # Confirmed on the real 2025-origional file: row 163
                # ("Narrabri") and row 172 ("NSW") both hold a plain
                # aggregate NUMBER already (not a formula) -- e.g. Q163 =
                # sum of Q164:Q171 -- so this writer needs to also update
                # that cell, which it previously never touched at all.
                # Detected structurally (no "total"-labelled indicator
                # anywhere in this town's JSON), not by checking which
                # state/source this is, so a future state fetcher
                # (VIC/TAS/NT/WA) gets the right behaviour automatically
                # based on whether ITS OWN json includes an explicit total.
                has_own_total_row = any(
                    "total" in (entry.get("label") or "").lower()
                    for entry in indicators.values()
                )
                town_total_series: dict[int, float] = {}

                # GENERALIZED 2026-09-24: reads label and values directly
                # from each fetcher's own JSON, not a hardcoded per-state
                # dict -- works for any state's indicator set.
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