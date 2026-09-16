"""
regional-indicators/transform/xlsx_update/base.py -- generic engine for
idempotently writing indicator values into Indicators_Data-Charts.xlsx,
driving Excel itself via COM automation (xlwings) rather than
reconstructing the file with a third-party library.

WHY THIS, NOT openpyxl:
openpyxl was confirmed to corrupt this specific workbook on save badly
enough that even Excel's own "recover as much as we can" repair could
not open the result. Diffing the raw OOXML parts (every xlsx is a zip of
XML files) showed openpyxl silently drops, on every save: every chart's
style/colour XML (all 231 charts), external links, threaded comments and
their author metadata (downgraded to legacy comments instead), custom
XML parts, an embedded image, and printer settings. This is a documented
category of openpyxl limitation, not a bug specific to this file, and
isn't fixable by patching around each dropped part individually. The old
version is kept as base_openpyxl_DEPRECATED.py for reference only -- do
not use it against the real workbook.

The only way to guarantee zero data loss is to have Excel itself do the
reading and writing -- the file is then never "reconstructed" by
anything other than the same program that always produces it.

REQUIRES a real Excel installation on the machine running this (Windows
or Mac with Excel). Will NOT work in a headless Linux environment -- an
inherent trade-off of this fix, not an oversight.

Sheet layout assumed (confirmed against the real workbook): row 1 =
fiscal-year labels from column C, row 2 = plain calendar years from
column C, then repeating <town name row> / <indicator rows> blocks down
column A. A town name can legitimately repeat across separate
geography-level sections (LGA / SA2 / UCL) on the same sheet --
ambiguous (town, indicator) matches raise rather than guess.

Performance note: cell-by-cell reads over COM are slow (each .value
access is a round-trip to the Excel process). The row/column finder
functions below do ONE bulk range read each rather than looping
cell-by-cell, which matters at this sheet's real size (~170 rows).
"""

from __future__ import annotations

import xlwings as xw

YEAR_HEADER_ROW = 2          # plain calendar year, e.g. 2001
FISCAL_HEADER_ROW = 1        # fiscal-year string, e.g. "2000/01"
FIRST_YEAR_COLUMN = 3         # column C
FIRST_DATA_ROW = 3            # rows 1-2 are headers; town blocks start row 3


def _fiscal_label(calendar_year: int) -> str:
    """2025 -> '2024/25', matching the existing header convention (fiscal
    year ending in the given calendar year)."""
    start = calendar_year - 1
    end_short = str(calendar_year)[-2:]
    return f"{start}/{end_short}"


def _find_year_column(sheet: "xw.Sheet", year: int) -> int:
    """Return the column index for `year`, creating a new column with both
    header rows filled in if it doesn't exist yet. Reads row 2's header
    values in a single bulk range read, not cell-by-cell. Explicitly sets
    number_format="General" on any newly-created header cells -- Excel
    was confirmed to silently inherit a percentage format from an
    adjacent cell on real data."""
    used = sheet.used_range
    max_col = used.last_cell.column

    header_row = sheet.range((YEAR_HEADER_ROW, FIRST_YEAR_COLUMN), (YEAR_HEADER_ROW, max_col)).value
    if not isinstance(header_row, list):
        header_row = [header_row]  # a single-cell range returns a scalar, not a list

    for offset, cell_year in enumerate(header_row):
        if cell_year == year:
            return FIRST_YEAR_COLUMN + offset

    # Not found -- append a new column right after the last one in use.
    new_col = max_col + 1
    fiscal_cell = sheet.cells(FISCAL_HEADER_ROW, new_col)
    year_cell = sheet.cells(YEAR_HEADER_ROW, new_col)
    fiscal_cell.value = _fiscal_label(year)
    fiscal_cell.number_format = "General"
    year_cell.value = year
    year_cell.number_format = "General"
    return new_col


def _find_town_indicator_row(
    sheet: "xw.Sheet", town: str, indicator: str, sub_label: str | None
) -> int:
    """Locate the row for `indicator` (optionally disambiguated by
    `sub_label`) inside `town`'s block, scanning the WHOLE sheet and
    raising if more than one match is found (this sheet genuinely has
    same-name, same-indicator collisions across its LGA/SA2/UCL sections
    -- confirmed on real data: Goondiwindi's 'Population (ERP)' differs
    by 42% between its LGA and SA2 rows). Reads columns A:B in a single
    bulk range read.
    """
    used = sheet.used_range
    max_row = used.last_cell.row

    ab_values = sheet.range((FIRST_DATA_ROW, 1), (max_row, 2)).value
    if ab_values and not isinstance(ab_values[0], list):
        ab_values = [ab_values]  # a single-row range returns a flat list, not nested

    matches: list[int] = []
    in_town_block = False
    for offset, (col_a, col_b) in enumerate(ab_values):
        row = FIRST_DATA_ROW + offset

        if col_a and col_b is None:
            in_town_block = col_a == town
            continue

        if in_town_block and col_a == indicator:
            if sub_label is None or col_b == sub_label:
                matches.append(row)

    if len(matches) == 1:
        return matches[0]

    if len(matches) == 0:
        raise ValueError(
            f"Could not find indicator '{indicator}'"
            f"{f' (sub-label {sub_label!r})' if sub_label else ''} "
            f"under town '{town}' in sheet '{sheet.name}'. "
            f"Check spelling/casing against the sheet exactly -- "
            f"this function does not guess or fuzzy-match, and does not "
            f"create a new row for you."
        )

    raise ValueError(
        f"AMBIGUOUS: {len(matches)} rows match indicator '{indicator}'"
        f"{f' (sub-label {sub_label!r})' if sub_label else ''} "
        f"under town '{town}' in sheet '{sheet.name}' (rows {matches}). "
        f"Pass sub_label to disambiguate, or check which row is correct "
        f"before proceeding."
    )


def read_existing_series(sheet: "xw.Sheet", row: int, exclude_col: int | None = None) -> dict[int, float]:
    """Read every existing (year -> value) pair in `row`, using row 2's
    headers to identify the year for each populated column. Skips
    exclude_col (the target column being written to -- irrelevant to
    "existing" history) and any column that isn't a plain number
    (blank cells, or a stray note/formula result that isn't numeric --
    those aren't part of a usable series for statistical comparison).
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


def write_one(sheet, town: str, indicator: str, sub_label: str | None, year: int, value, source_series: dict | None = None):
    """Shared audited-write helper: run both pre-write audits, write
    only if clean. Returns (WriteAuditReport, cell_address).

    source_series, when provided (a full freshly re-fetched history for
    this town/indicator), enables the ground-truth historical
    cross-check in audit.py's audit_historical_series -- comparing
    every existing year against the real source value, rather than
    guessing from the existing series' shape alone. Confirmed on real
    data: this clears a false positive the shape-based check alone
    produced (Isaac's 2012 NRW figure looked like an isolated miscopy
    from its shape, but exactly matches the real source).

    Used by both update_population_nrw.py (UCL-level, no source_series
    -- the UCL/FTE source only gives the latest year) and
    update_population_nrw_lga.py (LGA-level, source_series available --
    the LGA source gives full multi-year history).
    """
    from audit import audit_cell, audit_series, audit_historical_series, WriteAuditReport

    col = _find_year_column(sheet, year)
    row = _find_town_indicator_row(sheet, town, indicator, sub_label)
    cell = sheet.cells(row, col)

    cell_result = audit_cell(cell, value)
    existing_series = read_existing_series(sheet, row, exclude_col=col)
    series_result = audit_series(existing_series, year, value)

    historical_result = None
    if source_series:
        historical_result = audit_historical_series(existing_series, source_series)

    report = WriteAuditReport(cell_result, series_result, historical_result)

    if report.safe_to_write:
        cell.value = value
        cell.number_format = "General"

    return report, cell.address


def deep_audit_context(lga: str, lga_series: dict, flagged_years: list) -> str:
    """For a flagged UCL entry, check whether the corresponding LGA's
    own history shows a similar swing in the same year(s) -- a real
    regional workforce event should show up at both levels; a
    UCL-only blip is more likely a genuine data-entry error specific
    to that cell. This is corroborating context, not a verdict --
    still a human call, just a better-informed one. Confirmed on real
    data: correctly shows Isaac LGA's +26% swing in 2012 corroborating
    Moranbah's flagged UCL figure that year.
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
    """Open xlsx_path in a real Excel instance, write `value` into the
    cell for (town, indicator, year) on `sheet_name`, save, and close.
    Single-call convenience wrapper (opens/closes its own Excel
    instance) -- for a batch run over many towns, use write_one()
    directly inside one shared Excel session instead (see
    update_population_ucl.py / update_population_nrw.py for the
    pattern), since starting Excel per call is needlessly slow.
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