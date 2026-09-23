"""
regional-indicators/transform/xlsx_update/update_income.py
-----------------------------------------------------------------------
Writes fetch_income.py's (ATO Table 8) and fetch_income_table6.py's
(ATO Table 6) cached output into the Income sheet.

DELIBERATELY DOES NOT REUSE base.py's _find_town_indicator_row --
confirmed real, structural incompatibility (2026-09-23 inspection of
the real reference workbook): a town's block here is town-name row,
then POSTCODE row (a bare number in column A, column B empty), then
10 indicator rows. base.py's scanner treats any row with an empty
column B and a non-matching column A as "we've left this town's
block" -- which is exactly the postcode row's shape, so it would
incorrectly fall out of the town's block one row after entering it,
before ever reaching the real indicator rows. Needs its own row-
finder, same reasoning as update_rainfall.py needed one for Exogenous.

Confirmed real sheet structure:
  Row 1: fiscal-year labels ("2000/01", "2001/02", ...) starting at
    column C -- same format as base.py's _fiscal_label, reused
    directly. NOT the same as Population's convention though: there's
    no separate calendar-year row 2 here -- row 2 is "Postcode"/
    "Label" column headers, not years at all.
  Each town: name row, postcode row (confirmed matches towns.toml's
    postcode -- used as a sanity check when finding the row, not just
    trusting the name match alone), then a FIXED set of 10 indicator
    rows in the same order for every town, 12 rows total per town
    block, confirmed consistent for all 12 towns this sheet covers
    (the same 12 as Exogenous's Rainfall section -- VIC towns aren't
    covered here either).
  Six "Individual ABN..." rows exist in the template -- confirmed
    genuinely unpopulated for every town before this change (neither
    fetcher produced this category); now sourced from Table 6B's
    business-income fields, per Notes_DD.docx's documented mapping
    (the ATO file doesn't use the literal words "Individual ABN").
  A state-level benchmark section (State/Benchmark/Queensland/NSW)
    follows after row ~147 -- out of scope, not touched.

Indicator mapping, CORRECTED 2026-09-23 per Notes_DD.docx (Steve's
documented process for how the 2026 reference workbook was actually
built) and confirmed against real data:
  Table 6B's avg_income_all (latest year only)
    -> "Average Taxable Income or Loss (all individuals)"
    CHANGED from Table 8: Table 8's own pre-published "average taxable
    income" figure was confirmed WRONG for this row -- a real,
    non-trivial discrepancy (Chinchilla 2023-24: Table 8 gives
    $72,289, but Table 6B's documented method -- Taxable income or
    loss $ / no. -- gives $73,581, an exact match against the 2026
    reference workbook, verified directly). Table 8 still backfills
    OLDER years Table 6 doesn't cover in its single-year-per-release
    snapshot -- for any given town, whichever year Table 6 covers
    takes precedence entirely; Table 8's value for that same year is
    never written at all, avoiding any same-run conflict.
  Table 6A's avg_income_taxable (latest year only, filtered to
    Taxable status)
    -> "Average Taxable Income or Loss (taxable individuals)"
    Unchanged, already confirmed correct (exact match, $91,431).
  Table 6B's earners_no / wages_total (latest year only)
    -> "No. wage and salary earners" / "Total wage and salary earnings"
    Unchanged.
  Table 6B's six abn_*_income_total/no fields (latest year only)
    -> the six "Individual ABN..." rows, per Notes_DD.docx's mapping:
       NPP = non-primary production, PP = primary production.
    Notes don't explicitly say 6A vs 6B for these -- used 6B
    (unfiltered), consistent with wages already using 6B and none of
    these six labels carrying a "(taxable individuals)" qualifier the
    way the two average-income rows do. Worth Steve confirming this
    reading. Sanity-checked against real data: PP + NPP income exactly
    equals Total income for Chinchilla ($41,377,276 = $41,377,276),
    supporting that these are genuinely the right fields.

Table 8's full series is passed as source_series for the ground-truth
historical audit (like population_nrw_lga.py) when it's actually used
(i.e. the years Table 6 doesn't cover) -- Table 6's ten indicators
only ever have the latest year, so no source_series for those (like
population_nrw.py's UCL writes).

*** NOT YET TESTED against a live Excel instance. *** Row/column
finding logic is tested against a mock built to match the real Income
sheet structure, but the actual write against the real workbook
hasn't run yet -- test against a throwaway copy first, same as
everything else in this project's history.

Usage:
    python update_income.py <path-to-Indicators_Data-Charts.xlsx> <cache/ato dir> [--visible]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).parent))
from audit import audit_cell, audit_series, audit_historical_series, WriteAuditReport
from base import _fiscal_label

SHEET_NAME = "Income"
FIRST_YEAR_COLUMN = 3    # column C -- confirmed same as Population
YEAR_HEADER_ROW = 1      # fiscal-year strings -- confirmed same row Population uses,
                          # but Income has NO separate calendar-year row 2 to also match
FIRST_DATA_ROW = 3        # confirmed: row 1-2 headers, town blocks start row 3
TOWN_BLOCK_SIZE = 12      # confirmed consistent across all 12 towns this sheet covers

INDICATOR_ROW_LABELS = {
    "avg_taxable_income_all":     "Average Taxable Income or Loss (all individuals)",
    "avg_taxable_income_taxable": "Average Taxable Income or Loss (taxable individuals)",
    "earners_no":                 "No. wage and salary earners",
    "wages_total":                "Total wage and salary earnings",
    "abn_npp_income_total":       "Individual ABN NPP Total Income $",
    "abn_npp_income_no":          "Individual ABN NPP Total Income no.",
    "abn_pp_income_total":        "Individual ABN PP Total Income $",
    "abn_pp_income_no":           "Individual ABN PP Total Income no.",
    "abn_total_income_total":     "Individual ABN Total Income $",
    "abn_total_income_no":        "Individual ABN Total Income no.",
}


def _find_year_column(sheet, year: int) -> int:
    """Income-specific: row 1 holds fiscal-year STRINGS directly
    ("2000/01"), not a separate calendar-year number row like
    Population has. Parses each string's starting year to match
    against the target, and creates a new column (fiscal label only --
    there's no second header row to also fill in here) if needed.
    Confirmed via direct inspection not to trust used_range blindly
    (same lesson as update_rainfall.py's Exogenous bug) -- searches
    row 1 itself for the rightmost real fiscal-year label.
    """
    used = sheet.used_range
    max_col = used.last_cell.column

    header_row = sheet.range((YEAR_HEADER_ROW, FIRST_YEAR_COLUMN), (YEAR_HEADER_ROW, max_col)).value
    if not isinstance(header_row, list):
        header_row = [header_row]

    last_year_col = None
    for offset, label in enumerate(header_row):
        col = FIRST_YEAR_COLUMN + offset
        if isinstance(label, str) and "/" in label:
            # BUG FIXED 2026-09-23: was comparing start_year (2022 from
            # "2022/23") directly against `year` -- but `year` is always
            # passed in as the CALENDAR year the fiscal year ENDS in
            # (matching _fiscal_label's own convention: 2023 -> "2022/23"),
            # so this comparison could never match an existing column.
            # Confirmed real, serious damage on the first live run: every
            # write call created ANOTHER new column instead of finding the
            # one from the previous call, corrupting row 1 with 11 duplicate
            # "2023/24" headers (columns Z through AJ) in a single run.
            # Fixed to parse the label the same way _fy_to_calendar_year
            # does, so the two sides of the comparison use one convention.
            end_year = int(label.split("/")[0]) + 1
            if end_year == year:
                return col
            last_year_col = col

    if last_year_col is None:
        raise ValueError(
            f"Could not find any fiscal-year labels in row {YEAR_HEADER_ROW} "
            f"of sheet '{sheet.name}' -- can't determine where to add a new "
            f"year column safely."
        )

    new_col = last_year_col + 1
    year_cell = sheet.cells(YEAR_HEADER_ROW, new_col)
    year_cell.value = _fiscal_label(year)
    year_cell.number_format = "General"
    return new_col


def _find_income_town_row(sheet, town_name: str, expected_postcode: str) -> int:
    """Find the town-name row, confirmed against the postcode row
    directly below it as a sanity check (not just trusting a name
    match alone) -- the postcode is a cheap, independent cross-check
    that we've found the right block, not a different town that
    happens to share a name fragment.
    """
    used = sheet.used_range
    max_row = used.last_cell.row
    col_a = sheet.range((1, 1), (max_row, 1)).value
    if not isinstance(col_a, list):
        col_a = [col_a]

    matches = []
    for offset, val in enumerate(col_a):
        row = 1 + offset
        if val == town_name:
            matches.append(row)

    if len(matches) == 0:
        raise ValueError(
            f"Could not find town '{town_name}' in sheet '{sheet.name}'. "
            f"Check spelling/casing against the sheet exactly -- this "
            f"function does not guess or fuzzy-match, and does not create "
            f"a new town block for you."
        )
    if len(matches) > 1:
        raise ValueError(
            f"AMBIGUOUS: '{town_name}' appears at rows {matches} in sheet "
            f"'{sheet.name}'. Check for a genuine duplicate."
        )

    town_row = matches[0]
    postcode_cell = sheet.cells(town_row + 1, 1).value
    postcode_str = str(int(postcode_cell)) if isinstance(postcode_cell, (int, float)) else str(postcode_cell)
    if postcode_str != str(expected_postcode):
        raise ValueError(
            f"Found '{town_name}' at row {town_row}, but the postcode row "
            f"below it ({postcode_str!r}) doesn't match towns.toml's "
            f"postcode ({expected_postcode!r}). Not proceeding -- this "
            f"could be the wrong block or a stale/incorrect postcode."
        )
    return town_row


# Confirmed real, explicit label variants (2026-09-23, direct inspection
# of the reference workbook) -- Narrabri's block uses different exact
# wording from every other town for three rows. Same shape of issue as
# Rainfall's Narrabri quirk (plain "Summer"/"Winter" there). Listed
# explicitly per known case, not a general fuzzy-match weakening -- the
# row-finder still requires an exact match against one of a specific,
# documented set of strings, never a "close enough" guess.
ALTERNATE_LABELS = {
    "Average Taxable Income or Loss (all individuals)":     ["Average taxable income (all)"],
    "Average Taxable Income or Loss (taxable individuals)": ["Average taxable income (taxable)"],
    "Total wage and salary earnings":                        ["Total wage & salary earnings"],
}


def _find_income_indicator_row(sheet, town_row: int, indicator_label: str) -> int:
    """Search the bounded window belonging to this town's block
    (town_row+1 through town_row+TOWN_BLOCK_SIZE-1, i.e. up to but not
    including the next town's name row) for indicator_label, or one of
    its known explicit alternates (see ALTERNATE_LABELS). Bounded
    rather than scanning the whole sheet -- avoids ever matching into
    a different town's block or the benchmark section below all towns.
    """
    candidates = [indicator_label] + ALTERNATE_LABELS.get(indicator_label, [])
    for row in range(town_row + 1, town_row + TOWN_BLOCK_SIZE):
        if sheet.cells(row, 1).value in candidates:
            return row

    raise ValueError(
        f"Could not find indicator '{indicator_label}' (or its known "
        f"alternates {ALTERNATE_LABELS.get(indicator_label, [])}) within "
        f"rows {town_row + 1}-{town_row + TOWN_BLOCK_SIZE - 1} (town block "
        f"starting row {town_row}) in sheet '{sheet.name}'. Check the "
        f"label matches exactly -- this function does not guess or "
        f"fuzzy-match, and does not create a new row for you."
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
        if isinstance(label, str) and "/" in label and isinstance(val, (int, float)):
            year = int(label.split("/")[0])
            series[year] = val
    return series


def _describe_write(label: str, report, coord: str, extra_note: str = "") -> tuple[str, bool, bool]:
    """Shared three-way outcome description for a write attempt --
    written cleanly, written but visually flagged (bold red, see
    should_write_with_flag), or genuinely not written at all (a
    cell-level block, never overridden regardless of formatting).
    Returns (result_line, counts_as_written, counts_as_flagged).
    """
    suffix = f" ({extra_note})" if extra_note else ""
    if report.safe_to_write and not report.should_write_with_flag:
        # (should_write_with_flag is always False here since safe_to_write
        # is True, but spelled out for clarity -- this is the plain case)
        return f"{label}: WRITTEN{suffix} -> {coord}", True, False
    if report.should_write_with_flag:
        return (
            f"{label}: WRITTEN but FLAGGED FOR REVIEW (bold red in sheet){suffix} "
            f"-> {coord} — {report.summary_line()}",
            True, True,
        )
    return f"{label}: FLAGGED, not written — {report.summary_line()}", False, True


def _write_income_row(sheet, row: int, year: int, value, source_series: dict | None):
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
        # Explicitly reset formatting to plain -- a cell that was once
        # flagged (bold red, see below) and is now writing cleanly
        # (confirmed correct, or the concerning pattern resolved) should
        # not stay visually flagged forever.
        # BUG FIXED 2026-09-23: xlwings' real Font.color setter requires
        # an actual RGB tuple -- None crashes (rgb_to_int tries to
        # subscript it), confirmed via a real run against real Excel
        # (a mock couldn't catch this, since it doesn't reproduce
        # xlwings' own internal validation). Black (0,0,0) is the
        # standard default text colour, not a real "no colour" option.
        cell.font.bold = False
        cell.font.color = (0, 0, 0)
    elif report.should_write_with_flag:
        # Per Steve's request: a series-shape concern (not a cell-level
        # block) writes the value anyway, but in bold red -- easier to
        # spot and review directly in the spreadsheet than cross-
        # referencing a separate log. Cell-level blocks (a formula, an
        # unexpected existing value) never reach here -- see
        # should_write_with_flag's own docstring for why that boundary
        # matters.
        cell.value = value
        cell.number_format = "General"
        cell.font.bold = True
        cell.font.color = (255, 0, 0)

    return report, cell.address


def _fy_to_calendar_year(fy_label: str) -> int:
    """'2022-23' -> 2023 (the calendar year the fiscal year ENDS in) --
    matches the workbook's own convention, confirmed via base.py's
    _fiscal_label: 2025 -> '2024/25', i.e. columns are labelled by the
    calendar year they end in.
    """
    start = int(fy_label.replace("\u2013", "-").split("-")[0])
    return start + 1


# Confirmed real benchmark-section structure (2026-09-23): a "State"
# section header, then "Benchmark", then each state's own 3-row block
# (state-name row, then TWO indicator rows directly -- no postcode row
# like town blocks have). QLD and NSW use genuinely different exact
# wording for their own two rows, confirmed directly, not assumed:
STATE_LABELS = {
    "QLD": {
        "avg_income_all":     "Queensland Average Taxable Income or loss (all individuals) -ATO",
        "avg_income_taxable": "Average Taxable Income or Loss (taxable individuals) -ATO",
    },
    "NSW": {
        "avg_income_all":     "Average taxable income (all)",
        "avg_income_taxable": "Average taxable income (taxable)",
    },
}
STATE_HEADER_TEXT = {"QLD": "Queensland", "NSW": "NSW"}
STATE_BLOCK_SIZE = 3  # name row + 2 indicator rows, confirmed -- no postcode row here


def _find_state_row(sheet, state_code: str) -> int:
    """Find the state-name header row. Confirmed real quirk: the
    Queensland row has a trailing space ("Queensland ") -- matched by
    stripped comparison here specifically for the header row (the
    state's own two indicator rows below it still need exact-text
    matches via STATE_LABELS, only the header search is stripped).
    """
    used = sheet.used_range
    max_row = used.last_cell.row
    col_a = sheet.range((1, 1), (max_row, 1)).value
    if not isinstance(col_a, list):
        col_a = [col_a]

    target = STATE_HEADER_TEXT[state_code]
    matches = [
        1 + offset for offset, val in enumerate(col_a)
        if isinstance(val, str) and val.strip() == target
    ]
    if len(matches) == 0:
        raise ValueError(
            f"Could not find state header '{target}' in sheet '{sheet.name}'. "
            f"This function does not guess or create a new block for you."
        )
    if len(matches) > 1:
        raise ValueError(f"AMBIGUOUS: '{target}' appears at rows {matches}.")
    return matches[0]


def _find_state_indicator_row(sheet, state_row: int, label: str) -> int:
    for row in range(state_row + 1, state_row + STATE_BLOCK_SIZE):
        if sheet.cells(row, 1).value == label:
            return row
    raise ValueError(
        f"Could not find '{label}' within rows {state_row + 1}-"
        f"{state_row + STATE_BLOCK_SIZE - 1} (state block starting row "
        f"{state_row}) in sheet '{sheet.name}'."
    )


def update_income(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    t8_files = sorted(cache_dir.glob("*_income.json"))
    t6_files_by_slug = {
        p.stem.replace("_income_t6", ""): p
        for p in cache_dir.glob("*_income_t6.json")
        if p.name != "benchmark_income_t6.json"
        # BUG FIXED 2026-09-23: the glob pattern also matches
        # benchmark_income_t6.json (handled separately, below, via its
        # own distinct "benchmarks" structure rather than a "town"
        # field) -- caused a real crash (KeyError: 'town') the first
        # time a benchmark file actually existed to be swept in.
    }
    # BUG FIXED 2026-09-23: was keyed by filename slug ("chinchilla") but
    # looked up by the display name from Table 8's JSON ("Chinchilla") --
    # guaranteed mismatch, so skip_this_year was silently always False.
    # Table 8 wrote into every "all individuals" cell first every time,
    # and only the audit's 1% tolerance masked it for towns where the two
    # sources' values happened to be close (Dysart, Moranbah, Toowoomba) --
    # everywhere else (Chinchilla, Dalby, Goondiwindi, Miles, Roma, Tara,
    # Wallumbilla, Wandoan) it surfaced as a confusing "cell already
    # contains X" flag against a value THIS SAME RUN had just written.
    # Rebuilt keyed by display name (reading each file's own "town" field),
    # so both loops use one consistent key.
    t6_files = {}
    for path in t6_files_by_slug.values():
        with open(path, encoding="utf-8") as f:
            t6_files[json.load(f)["town"]] = path

    if not t8_files and not t6_files:
        raise FileNotFoundError(
            f"No *_income.json or *_income_t6.json files found in {cache_dir} "
            f"-- run fetch_income.py and fetch_income_table6.py first."
        )

    import tomllib
    towns_toml_path = Path(__file__).parent.parent.parent / "towns.toml"
    with open(towns_toml_path, "rb") as f:
        towns_data = tomllib.load(f)
    postcode_by_name = {t["name"]: t["postcode"] for t in towns_data["towns"].values()}

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

            for t8_path in t8_files:
                with open(t8_path, encoding="utf-8") as f:
                    data = json.load(f)

                town = data["town"]
                series_by_fy = data.get("avg_taxable_income_by_year", {})
                if not series_by_fy:
                    results.append(f"{town} avg taxable income (all): no data, skipped")
                    continue

                expected_postcode = postcode_by_name.get(town)
                if expected_postcode is None:
                    results.append(f"{town}: not found in towns.toml, skipped")
                    continue

                try:
                    town_row = _find_income_town_row(sheet, town, expected_postcode)
                except ValueError as exc:
                    results.append(f"{town}: SKIPPED (row-finding) — {exc}")
                    flagged_count += 1
                    continue

                source_series = {_fy_to_calendar_year(fy): v for fy, v in series_by_fy.items()}
                latest_fy = max(series_by_fy, key=lambda fy: int(fy.split("-")[0]))
                latest_year = _fy_to_calendar_year(latest_fy)
                latest_value = series_by_fy[latest_fy]

                # Table 6B, not Table 8, is the documented/verified-correct
                # source for "all individuals" (see module docstring and
                # Notes_DD.docx) -- if Table 6 covers this same year for
                # this town, it takes over that specific year entirely, so
                # Table 8's own (confirmed different, less accurate) figure
                # for that year never gets written at all, rather than
                # writing it and then immediately overwriting/conflicting
                # with Table 6's value in the same run.
                t6_path = t6_files.get(town)
                skip_this_year = False
                if t6_path:
                    with open(t6_path, encoding="utf-8") as f:
                        t6_data = json.load(f)
                    t6_cal_year = int(t6_data.get("cal_year", 0))
                    if t6_cal_year == latest_year:
                        skip_this_year = True

                if skip_this_year:
                    results.append(
                        f"{town} avg taxable income (all): {latest_year} deferred to "
                        f"Table 6B (verified-correct source for this year) -- see below"
                    )
                else:
                    label = INDICATOR_ROW_LABELS["avg_taxable_income_all"]
                    try:
                        row = _find_income_indicator_row(sheet, town_row, label)
                        report, coord = _write_income_row(sheet, row, latest_year, latest_value, source_series)
                        line, is_written, is_flagged = _describe_write(
                            f"{town} avg taxable income (all)", report, coord,
                            f"{latest_year} = {latest_value:,.0f}, Table 8, no overlapping Table 6 data",
                        )
                        results.append(line)
                        if is_written:
                            written_count += 1
                            any_written = True
                        if is_flagged:
                            flagged_count += 1
                    except ValueError as exc:
                        results.append(f"{town} avg taxable income (all): SKIPPED (row-finding) — {exc}")
                        flagged_count += 1

            for town, t6_path in t6_files.items():
                with open(t6_path, encoding="utf-8") as f:
                    data = json.load(f)

                town_name = data["town"]
                indicators = data.get("indicators", {})
                expected_postcode = postcode_by_name.get(town_name)
                if expected_postcode is None:
                    results.append(f"{town_name}: not found in towns.toml, skipped")
                    continue

                try:
                    town_row = _find_income_town_row(sheet, town_name, expected_postcode)
                except ValueError as exc:
                    results.append(f"{town_name} (Table 6): SKIPPED (row-finding) — {exc}")
                    flagged_count += 3
                    continue

                for key, row_label_key in (
                    ("avg_income_all", "avg_taxable_income_all"),
                    ("avg_income_taxable", "avg_taxable_income_taxable"),
                    ("earners_no", "earners_no"),
                    ("wages_total", "wages_total"),
                    ("abn_npp_income_total", "abn_npp_income_total"),
                    ("abn_npp_income_no", "abn_npp_income_no"),
                    ("abn_pp_income_total", "abn_pp_income_total"),
                    ("abn_pp_income_no", "abn_pp_income_no"),
                    ("abn_total_income_total", "abn_total_income_total"),
                    ("abn_total_income_no", "abn_total_income_no"),
                ):
                    values_by_year = indicators.get(key, {})
                    if not values_by_year:
                        results.append(f"{town_name} {key}: no data, skipped")
                        continue

                    year_str = next(iter(values_by_year))
                    value = values_by_year[year_str]
                    if value is None:
                        results.append(f"{town_name} {key}: value is null, skipped")
                        continue
                    year = int(year_str)

                    label = INDICATOR_ROW_LABELS[row_label_key]
                    try:
                        row = _find_income_indicator_row(sheet, town_row, label)
                        report, coord = _write_income_row(sheet, row, year, value, None)
                        line, is_written, is_flagged = _describe_write(
                            f"{town_name} {key}", report, coord, f"{year} = {value:,.0f}"
                        )
                        results.append(line)
                        if is_written:
                            written_count += 1
                            any_written = True
                        if is_flagged:
                            flagged_count += 1
                    except ValueError as exc:
                        results.append(f"{town_name} {key}: SKIPPED (row-finding) — {exc}")
                        flagged_count += 1

            if any_written:
                wb.save()

            # State benchmarks -- separate block structure, separate cache file
            benchmark_path = cache_dir / "benchmark_income_t6.json"
            if benchmark_path.exists():
                with open(benchmark_path, encoding="utf-8") as f:
                    bdata = json.load(f)
                any_benchmark_written = False
                for state_code, indicators in bdata.get("benchmarks", {}).items():
                    try:
                        state_row = _find_state_row(sheet, state_code)
                    except ValueError as exc:
                        results.append(f"{state_code} benchmark: SKIPPED (row-finding) — {exc}")
                        flagged_count += 2
                        continue

                    for key in ("avg_income_all", "avg_income_taxable"):
                        values_by_year = indicators.get(key, {})
                        if not values_by_year:
                            continue
                        year_str = next(iter(values_by_year))
                        value = values_by_year[year_str]
                        if value is None:
                            continue
                        year = int(year_str)
                        label = STATE_LABELS[state_code][key]
                        try:
                            row = _find_state_indicator_row(sheet, state_row, label)
                            report, coord = _write_income_row(sheet, row, year, value, None)
                            line, is_written, is_flagged = _describe_write(
                                f"{state_code} {key}", report, coord, f"{year} = {value:,.0f}"
                            )
                            results.append(line)
                            if is_written:
                                written_count += 1
                                any_benchmark_written = True
                            if is_flagged:
                                flagged_count += 1
                        except ValueError as exc:
                            results.append(f"{state_code} {key}: SKIPPED (row-finding) — {exc}")
                            flagged_count += 1
                if any_benchmark_written:
                    wb.save()
            else:
                results.append(
                    f"State benchmarks: no {benchmark_path.name} found in {cache_dir} -- "
                    f"run fetch_income_table6.py first"
                )
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
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/ato dir> [--visible]")
        sys.exit(1)

    xlsx_path = Path(args[0])
    cache_dir = Path(args[1])

    results = update_income(xlsx_path, cache_dir, visible=visible)
    for line in results:
        print(line)