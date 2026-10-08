"""
regional-indicators/transform/xlsx_update/base.py

Shared engine for writing indicator values into the indicators workbook.

The workbook is edited through Excel itself (xlwings, COM automation on
Windows / AppleScript on macOS) rather than rebuilt with a file-level
library. openpyxl cannot round-trip this workbook: on save it drops chart
style parts, external links, threaded comments, custom XML parts,
embedded images and printer settings, and the result does not reopen in
Excel. Letting Excel perform every read and write guarantees the file is
only ever produced by the application that owns the format.

Requirements: a local Microsoft Excel installation (Windows or macOS).
The writers cannot run headless or on Linux.

Sheet layout assumed by the helpers in this module:
  row 1      fiscal-year labels, from column C
  row 2      calendar years, from column C
  row 3+     repeating blocks in column A: a region heading row followed
             by its indicator rows
A sheet may be divided into geography sections (LGA / SA2 / UCL), each
introduced by a row whose column A holds the section name. The same
region and indicator names can occur in more than one section, so
lookups that match more than one row raise instead of choosing.

Performance: each cell access over COM is a round trip to the Excel
process, so the finder functions read whole ranges in one call.
"""

from __future__ import annotations

from pathlib import Path

import xlwings as xw

YEAR_HEADER_ROW = 2          # calendar-year header row, e.g. 2001
FISCAL_HEADER_ROW = 1        # fiscal-year header row, e.g. "2000/01"
FIRST_YEAR_COLUMN = 3         # column C
FIRST_DATA_ROW = 3            # rows 1-2 are headers; region blocks start at row 3


def _fiscal_label(calendar_year: int) -> str:
    """Return the fiscal-year label ending in `calendar_year`: 2025 -> '2024/25'."""
    start = calendar_year - 1
    end_short = str(calendar_year)[-2:]
    return f"{start}/{end_short}"


def _find_year_column(sheet: "xw.Sheet", year: int) -> int:
    """Return the column index for `year`, appending a new column if needed.

    A new column is placed immediately after the last used column and
    both header rows are filled in. The header cells are set to the
    "General" number format explicitly, because Excel otherwise copies
    the format of the neighbouring cell (which may be a percentage).
    The header row is read in a single range call.
    """
    used = sheet.used_range
    max_col = used.last_cell.column

    header_row = sheet.range((YEAR_HEADER_ROW, FIRST_YEAR_COLUMN), (YEAR_HEADER_ROW, max_col)).value
    if not isinstance(header_row, list):
        header_row = [header_row]  # a single-cell range returns a scalar

    for offset, cell_year in enumerate(header_row):
        if cell_year == year:
            return FIRST_YEAR_COLUMN + offset

    # Year not present: append a column after the last one in use.
    new_col = max_col + 1
    fiscal_cell = sheet.cells(FISCAL_HEADER_ROW, new_col)
    year_cell = sheet.cells(YEAR_HEADER_ROW, new_col)
    fiscal_cell.value = _fiscal_label(year)
    fiscal_cell.number_format = "General"
    year_cell.value = year
    year_cell.number_format = "General"
    return new_col


SECTION_LABELS = {"LGA", "SA2", "UCL"}


def _find_town_indicator_row(
    sheet: "xw.Sheet", town: str, indicator: str, sub_label: str | None,
    section: str | None = None,
) -> int:
    """Return the row of `indicator` within `town`'s block.

    Scans the whole sheet and raises ValueError if no row, or more than
    one row, matches. Columns A:B are read in a single range call.

    sub_label: column B text, used to distinguish rows that share an
        indicator name within one block.
    section: "LGA", "SA2" or "UCL". Restricts the match to that
        geography section. Sections are delimited by rows whose column
        A holds the section name and whose column B is empty. When
        omitted, every section is searched and an ambiguous match
        raises.
    """
    used = sheet.used_range
    max_row = used.last_cell.row

    ab_values = sheet.range((FIRST_DATA_ROW, 1), (max_row, 2)).value
    if ab_values and not isinstance(ab_values[0], list):
        ab_values = [ab_values]  # a single-row range returns a flat list

    matches: list[int] = []
    in_town_block = False
    current_section: str | None = None

    # The first section heading can sit in the header rows above
    # FIRST_DATA_ROW (the Population sheet has "LGA" in A2, alongside the
    # year headers), where the scan below does not look. Seed the current
    # section from those rows.
    header_a = sheet.range((1, 1), (FIRST_DATA_ROW - 1, 1)).value
    if not isinstance(header_a, list):
        header_a = [header_a]
    for head_val in header_a:
        if head_val in SECTION_LABELS:
            current_section = head_val

    for offset, (col_a, col_b) in enumerate(ab_values):
        row = FIRST_DATA_ROW + offset

        if col_a and col_b is None:
            if col_a in SECTION_LABELS:
                current_section = col_a
                in_town_block = False  # a section heading is never a region row
                continue
            in_town_block = col_a == town
            continue

        if in_town_block and col_a == indicator:
            if section is not None and current_section != section:
                continue
            if sub_label is None or col_b == sub_label:
                matches.append(row)

    if len(matches) == 1:
        return matches[0]

    if len(matches) == 0:
        raise ValueError(
            f"Could not find indicator '{indicator}'"
            f"{f' (sub-label {sub_label!r})' if sub_label else ''}"
            f"{f' (section {section!r})' if section else ''} "
            f"under town '{town}' in sheet '{sheet.name}'. "
            f"Check spelling/casing against the sheet exactly -- "
            f"this function does not guess or fuzzy-match, and does not "
            f"create a new row for you."
        )

    raise ValueError(
        f"AMBIGUOUS: {len(matches)} rows match indicator '{indicator}'"
        f"{f' (sub-label {sub_label!r})' if sub_label else ''} "
        f"under town '{town}' in sheet '{sheet.name}' (rows {matches}). "
        f"Pass section='LGA'/'SA2'/'UCL' (or sub_label) to disambiguate, "
        f"or check which row is correct before proceeding."
    )


def read_existing_series(sheet: "xw.Sheet", row: int, exclude_col: int | None = None) -> dict[int, float]:
    """Return the existing {year: value} pairs in `row`.

    Years come from the calendar-year header row. `exclude_col` (the
    column about to be written) and non-numeric cells are skipped.
    """
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


OVERRIDES_PATH = Path(__file__).parent.parent.parent / "verified_overrides.toml"
_overrides_cache: list[dict] | None = None


def _load_overrides() -> list[dict]:
    """Load and cache verified_overrides.toml.

    Each entry records a (town, indicator, sub_label, year, value)
    combination that has been verified independently and may be written
    even though an audit flags it. A missing file means no overrides.
    """
    global _overrides_cache
    if _overrides_cache is not None:
        return _overrides_cache

    if not OVERRIDES_PATH.exists():
        _overrides_cache = []
        return _overrides_cache

    import tomllib
    with open(OVERRIDES_PATH, "rb") as f:
        data = tomllib.load(f)
    _overrides_cache = data.get("override", [])
    return _overrides_cache


def _check_override(town: str, indicator: str, sub_label: str | None, year: int, value) -> str | None:
    """Return the override reason for an exact match, else None.

    All of town, indicator, sub_label, year and value must match. If the
    fetched value changes, the override no longer applies and the audit
    flag is reported again.
    """
    for entry in _load_overrides():
        if (
            entry.get("town") == town
            and entry.get("indicator") == indicator
            and entry.get("sub_label") == sub_label
            and entry.get("year") == year
            and entry.get("value") == value
        ):
            return entry.get("reason", "verified override, no reason recorded")
    return None


def write_one(
    sheet, town: str, indicator: str, sub_label: str | None, year: int, value,
    source_series: dict | None = None, section: str | None = None,
):
    """Audit and write one value. Returns (WriteAuditReport, cell_address).

    The value is written only when the audits pass, or when an exact
    entry in verified_overrides.toml allows it.

    source_series: full {year: value} history from the source. When
        given, existing workbook values are compared with it directly
        (audit_historical_series), which replaces the shape-based
        outlier check.
    section: "LGA", "SA2" or "UCL"; see _find_town_indicator_row.
    """
    from audit import audit_cell, audit_series, audit_historical_series, WriteAuditReport

    col = _find_year_column(sheet, year)
    row = _find_town_indicator_row(sheet, town, indicator, sub_label, section=section)
    cell = sheet.cells(row, col)

    cell_result = audit_cell(cell, value)
    existing_series = read_existing_series(sheet, row, exclude_col=col)
    series_result = audit_series(existing_series, year, value)

    historical_result = None
    if source_series:
        historical_result = audit_historical_series(existing_series, source_series)

    report = WriteAuditReport(cell_result, series_result, historical_result)
    if not report.safe_to_write:
        report.override_reason = _check_override(town, indicator, sub_label, year, value)

    if report.safe_to_write:
        cell.value = value
        cell.number_format = "General"

    return report, cell.address


def deep_audit_context(lga: str, lga_series: dict, flagged_years: list) -> str:
    """Describe whether an LGA series corroborates a flagged UCL value.

    A genuine regional change normally appears at both the urban-centre
    and LGA level; a change confined to the urban-centre series is more
    likely a data-entry error. The returned text is context for a
    reviewer, not a decision.
    """
    if not lga_series or not flagged_years:
        return f"  [deep-audit] No LGA series available for {lga} to cross-check against."

    lines = [f"  [deep-audit] Cross-checking against {lga} LGA data:"]
    for year in flagged_years:
        prev_v = lga_series.get(str(year - 1)) or lga_series.get(year - 1)
        cur_v = lga_series.get(str(year)) or lga_series.get(year)
        if prev_v is None or cur_v is None or prev_v == 0:
            lines.append(f"    {year}: LGA data unavailable for this year, can't cross-check.")
            continue
        pct_change = (cur_v - prev_v) / prev_v * 100
        verdict = (
            "corroborates a real regional swing"
            if abs(pct_change) >= 15
            else "shows little matching movement -- weaker case for a genuine change"
        )
        lines.append(
            f"    {year}: {lga} LGA went {prev_v:,} -> {cur_v:,} ({pct_change:+.0f}%) -- {verdict}"
        )
    return "\n".join(lines)


def update_indicator_value(
    xlsx_path: str,
    sheet_name: str,
    town: str,
    indicator: str,
    year: int,
    value,
    sub_label: str | None = None,
    visible: bool = False,
) -> str:
    """Write a single value, opening and closing its own Excel session.

    Convenience wrapper for one-off writes. Batch writers should open
    one Excel session and call write_one() for each value instead.
    """
    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(xlsx_path)
        try:
            sheet = wb.sheets[sheet_name]
            report, coord = write_one(sheet, town, indicator, sub_label, year, value)
            if report.safe_to_write:
                wb.save()
            return coord
        finally:
            wb.close()
    finally:
        app.quit()