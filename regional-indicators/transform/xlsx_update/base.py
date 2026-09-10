"""
xlsx_update/base.py -- generic engine for idempotently writing indicator
values into Indicators_Data-Charts.xlsx, driving Excel itself via COM
automation (xlwings) rather than reconstructing the file with a
third-party library.

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

*** NOT YET TESTED against a live Excel instance. *** This was written
carefully, following documented xlwings patterns, but could not be run
in the environment it was developed in (no Windows/Excel available
there). Treat the first real run as a genuine test:
  - Run against a throwaway copy first, never the real workbook.
  - Consider your first call with visible=True (see update_indicator_value's
    `visible` parameter) so you can watch Excel actually do it, rather
    than trusting a background process blind.
  - Check the saved file opens cleanly afterwards with no repair prompt,
    same as you would for any first real run of new automation.

Sheet layout assumed (same as before, confirmed against the real
workbook): row 1 = fiscal-year labels from column C, row 2 = plain
calendar years from column C, then repeating <town name row> /
<indicator rows> blocks down column A. A town name can legitimately
repeat across separate geography-level sections (LGA / SA2 / UCL) on the
same sheet -- ambiguous (town, indicator) matches raise rather than
guess, same principle as the openpyxl version.

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
    values in a single bulk range read, not cell-by-cell."""
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
    cell for (town, indicator, year) on `sheet_name`, save, and close --
    using Excel's own save, so every part of the file Excel itself
    manages (chart styling, external links, threaded comments, etc.) is
    preserved exactly as Excel would always produce it.

    Idempotent: calling this again with the same arguments writes the
    same cell, not a new one.

    Explicitly sets the written cell's number format to General --
    confirmed necessary on real data: Excel silently inherited a
    percentage format from an adjacent cell for at least one newly
    written value. Relying on whatever format Excel happens to carry
    over is not safe.

    visible=True runs Excel on-screen so you can watch it happen --
    recommended for your first real test of this, before trusting it to
    run invisibly in the background.

    Opens and closes its own Excel instance per call. If you're calling
    this many times in a loop (e.g. once per town), that's a lot of
    Excel start/stop overhead -- see update_population_ucl.py for the
    batched version that opens Excel once for the whole run instead.
    """
    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False  # suppress "keep current format?" etc. prompts
    try:
        wb = app.books.open(xlsx_path)
        try:
            sheet = wb.sheets[sheet_name]
            col = _find_year_column(sheet, year)
            row = _find_town_indicator_row(sheet, town, indicator, sub_label)
            cell = sheet.cells(row, col)
            cell.value = value
            cell.number_format = "General"
            coord = cell.address
            wb.save()
            return coord
        finally:
            wb.close()
    finally:
        app.quit()