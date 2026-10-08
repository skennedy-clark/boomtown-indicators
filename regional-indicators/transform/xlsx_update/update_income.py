"""
regional-indicators/transform/xlsx_update/update_income.py

Writes ATO taxation statistics into the Income sheet.

Input:  cache/ato/<slug>_income.json      (ATO Table 8, fetch_income.py)
        cache/ato/<slug>_income_t6.json   (ATO Table 6, fetch_income_table6.py)
        cache/ato/benchmark_income_t6.json (state benchmarks)
Target: the Income sheet.

Sheet layout:
  Row 1  fiscal-year labels ("2000/01", ...) from column C. There is no
         calendar-year header row; row 2 holds column headings.
  Each town is a block of twelve rows: the town name, its postcode
  (a number in column A), then ten indicator rows in a fixed order.
  A "State" section follows the towns, with a three-row block per state
  (name row and two indicator rows, no postcode row).

This sheet has its own row finders. base.py's block scanner treats a
row with an empty column B as the end of a block, which the postcode
row would trigger. The postcode is instead used to confirm that the
right town block has been found.

Indicator sources:
  Average Taxable Income or Loss (all individuals)
      Table 6B, taxable income or loss ($) divided by number of
      individuals, for the year Table 6 covers. Table 8's published
      average is used only for earlier years that Table 6 does not
      cover; where both cover a year, Table 6 is written and Table 8's
      figure for that year is not.
  Average Taxable Income or Loss (taxable individuals)
      Table 6A, restricted to taxable individuals.
  No. wage and salary earners / Total wage and salary earnings
      Table 6B.
  Individual ABN ... (six rows)
      Table 6B business income fields. NPP is non-primary production,
      PP is primary production.

Table 8 provides a multi-year series, which is passed to the historical
audit for the years it supplies. Table 6 provides one year per release,
so those writes use the cell and series audits only.

A series-level audit concern does not block the write: the value is
written and marked bold red for review. A cell-level block (a formula
or unexpected content) prevents the write.

The workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_income.py <workbook.xlsx> <cache/ato dir> [--visible]
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
FIRST_YEAR_COLUMN = 3    # column C
YEAR_HEADER_ROW = 1      # fiscal-year labels; this sheet has
                          # no calendar-year header row
FIRST_DATA_ROW = 3        # rows 1-2 are headers; town blocks start at row 3
TOWN_BLOCK_SIZE = 12      # name row, postcode row and 10 indicator rows

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
    """Return the column for the financial year ending in `year`, appending
    one if needed.

    Row 1 holds fiscal-year labels. The scan locates the rightmost label
    in row 1 itself and does not rely on the sheet's used range, which
    extends beyond the year columns. A new column receives its fiscal
    label.
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
            # `year` is the calendar year in which the financial year ends, so a
            # label "2022/23" corresponds to year 2023.
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
    """Return the row of a town's name, verified against the postcode in
    the row below it.
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


# Narrabri's block words three row labels differently from the other
# towns. The accepted alternatives are listed explicitly; matching
# remains exact.
ALTERNATE_LABELS = {
    "Average Taxable Income or Loss (all individuals)":     ["Average taxable income (all)"],
    "Average Taxable Income or Loss (taxable individuals)": ["Average taxable income (taxable)"],
    "Total wage and salary earnings":                        ["Total wage & salary earnings"],
}


def _find_income_indicator_row(sheet, town_row: int, indicator_label: str) -> int:
    """Return the row of `indicator_label`, or one of its listed
    alternatives, within the town's block.

    The search is limited to the rows of that block, so a label is never
    matched in another town's block or in the State section.
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
    """Describe the outcome of a write attempt.

    Returns (result_line, counts_as_written, counts_as_flagged) for the
    three cases: written, written and flagged for review, or not
    written.
    """
    suffix = f" ({extra_note})" if extra_note else ""
    if report.safe_to_write and not report.should_write_with_flag:
        # safe_to_write is True here, so should_write_with_flag is False.
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
        # A clean write is set to regular black, clearing any earlier flag.
        # xlwings requires an explicit RGB tuple for Font.color.
        cell.font.bold = False
        cell.font.color = (0, 0, 0)
    elif report.should_write_with_flag:
        # A series-level concern: write the value and mark it bold red for
        # review. Cell-level blocks never reach this branch.
        cell.value = value
        cell.number_format = "General"
        cell.font.bold = True
        cell.font.color = (255, 0, 0)

    return report, cell.address


def _fy_to_calendar_year(fy_label: str) -> int:
    """Convert a fiscal-year string to the calendar year it ends in:
    '2022-23' -> 2023.
    """
    start = int(fy_label.replace("\u2013", "-").split("-")[0])
    return start + 1


# State section: a "State" heading, then "Benchmark", then a three-row
# block per state (name row and two indicator rows; no postcode row).
# Queensland and NSW word their two rows differently:
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
STATE_BLOCK_SIZE = 3  # name row + 2 indicator rows


def _find_state_row(sheet, state_code: str) -> int:
    """Return the row of a state's name in the State section.

    The name is compared with surrounding whitespace removed; the
    indicator rows below it are matched exactly via STATE_LABELS.
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
        # The glob also matches benchmark_income_t6.json, which has a different
        # structure and is processed separately below.
    }
    # Keyed by the town name recorded in each file, the same key the
    # Table 8 loop uses.
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
    # Regions with no postcode (LGA-level benchmarks) have no Income block.
    postcode_by_name = {
        t["name"]: t["postcode"] for t in towns_data["towns"].values() if t.get("postcode")
    }

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

                # Table 6 is the source for "all individuals" in the year it covers.
                # Where it covers this year for this town, Table 8's figure for the
                # same year is not written.
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

            # State benchmarks: separate block structure and cache file.
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