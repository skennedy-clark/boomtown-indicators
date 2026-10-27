"""
regional-indicators/transform/xlsx_update/update_population_projections.py

Writes population projections into the Population sheet, each under its
own year, and first moves any projection the sheet holds under another
year's column to its own.

Input:  cache/population/population_projections.json, from
        fetchers/fetch_population_projections.py.
Target: the Population sheet.

Rows written
  Queensland, QGSO 2018 edition (the projection sections of the sheet):
    "Projected Population - High|Low|Medium series: Persons Total
     (Persons), 2018ed"            in each LGA's block, from the LGA table;
    "Population projection: district 2018e (Medium Series)"
                                   in each SA2's block, from the SA2 table.
    Years 2016, 2021, ..., 2041 (2016 is the base estimate).
  Queensland resource regions, QGSO non-resident population projections
  2022 to 2026:
    "Population projection: NRW 2022e (Series A|B|C|D)" in each LGA's
    block, years 2022-2026 (the 2021 estimate they start from is the
    observed non-resident row's figure, not a projection). A series QGSO
    does not publish for the LGA (Central Highlands: Series B) is left
    empty and reported. QGSO has since withdrawn these projections ("no
    longer current and should not be used"); they are kept as the
    historical vintage the sheet shows.
  New South Wales, NSW Planning projections: the "Population (ERP)" row
    of the LGA and SA2 blocks of New South Wales towns (towns.toml), in
    columns after the last data year. The years are those the other New
    South Wales rows show (Narrabri: 2031, 2036, 2041); with none to
    follow, the five-year points 2031, 2036, 2041.

  A block is matched to a source area by name, compared loosely ("Isaac
  (R)" = "Isaac", "Broadsound - Nebo" = "Broadsound-Nebo"). An SA2 renamed
  since 2016 is found by its code, through the town in towns.toml whose
  SA2 it is (Roma Surrounds was "Roma Region" in the 2018 edition).

Moving projections to their own years
  The workbooks up to 2026 held the 2018 edition's 2016, 2021 and 2026
  figures under the columns headed 2011, 2017 and 2025, and the
  non-resident projections for 2022-2026 under 2019, 2021, 2023, 2024
  and 2026. A row is moved only when its figures are, in column order,
  exactly the source's figures in year order (within rounding), so a
  row is never rearranged on a guess; each move is listed. A row whose
  figures do not match the source in that way is left as it is and
  reported. New South Wales rows are already under their own years and
  are not moved. The charts that read these rows span the years
  concerned, so they need no change.

Rounding follows the rows already filled (the LGA rows hold whole
numbers, the SA2 rows the source's decimals).

Only empty cells are written. A cell that already holds a different
figure is kept and reported.

The workbook is edited through Excel (xlwings); see base.py.

Usage:
    python update_population_projections.py <workbook.xlsx> <cache/population dir> [--visible]
"""

from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent.parent))

from fetchers.fetch_population_projections import loose_name  # noqa: E402

SHEET_NAME = "Population"
CACHE_FILE = "population_projections.json"
SECTIONS = {"LGA": "lga", "SA2": "sa2", "UCL": "ucl"}

QLD_LGA_ROW = re.compile(r"^Projected Population - (High|Low|Medium) series\b.*2018ed", re.IGNORECASE)
QLD_SA2_ROW = re.compile(r"^Population projection: district 2018e \(Medium Series\)", re.IGNORECASE)
NRW_ROW = re.compile(r"^Population projection: NRW 2022e \(Series ([A-D])\)", re.IGNORECASE)
ERP_ROW = "Population (ERP)"
NSW_DEFAULT_YEARS = (2031, 2036, 2041)
MATCH_TOLERANCE = 0.5


def _number(value) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        from decimal import Decimal
        if isinstance(value, Decimal):
            return float(value)
    except ImportError:
        pass
    return None


def _empty(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def header_years(grid: list[list]) -> dict[int, int]:
    """{column: year} (1-based) from the first row holding at least three
    years in increasing order (the sheet's calendar-year header)."""
    for row in grid:
        cells = {c: _number(v) for c, v in enumerate(row, start=1) if c > 1 and _number(v) is not None}
        values = list(cells.values())
        if (len(values) >= 3 and all(v.is_integer() and 1990 <= v <= 2100 for v in values)
                and all(b > a for a, b in zip(values, values[1:]))):
            return {c: int(v) for c, v in cells.items()}
    return {}


@dataclass
class Target:
    row: int
    key: tuple              # row type: ("qld_lga", "medium"), ("qld_sa2", "medium"), ("nsw", "lga")
    area: str               # block label
    series: dict[int, float]
    years: tuple = ()       # the years to show; empty = every year of the series

    def shown(self) -> dict[int, float]:
        return {y: v for y, v in self.series.items() if not self.years or y in self.years}


def find_targets(grid: list[list], data: dict, towns: list) -> tuple[list[Target], list[str]]:
    """The rows to fill and the source series for each."""
    qld = data.get("qld_2018") or {}
    nsw = data.get("nsw") or {}
    qld_lga = {loose_name(k): v for k, v in (qld.get("lga") or {}).items()}
    qld_sa2 = {loose_name(k): v for k, v in (qld.get("sa2") or {}).items()}
    qld_sa2_code = {v.get("code"): v for v in (qld.get("sa2") or {}).values()}
    sa2_code_of = {loose_name(t.sa2_name): str(getattr(t, "qgso_sa2", "") or t.sa2_code or "")
                   for t in towns if getattr(t, "sa2_name", "")}
    nsw_areas = {"lga": {loose_name(re.sub(r"\s*\([^)]*\)\s*$", "", t.lga)) for t in towns if t.state == "NSW"},
                 "sa2": {loose_name(t.sa2_name) for t in towns if t.state == "NSW"}}
    nsw_series = {level: {loose_name(k): v for k, v in (nsw.get(level) or {}).items()} for level in ("lga", "sa2")}
    nrw = {loose_name(k): v for k, v in ((data.get("nrw_2022") or {}).get("lga") or {}).items()}

    targets, notes = [], []
    section = block = None
    for r, row in enumerate(grid, start=1):
        a = row[0] if row else None
        label = a.strip() if isinstance(a, str) else ""
        if label in SECTIONS:
            section, block = SECTIONS[label], None
            continue
        b = row[1] if len(row) > 1 else None
        is_data_row = bool(QLD_LGA_ROW.match(label) or QLD_SA2_ROW.match(label) or NRW_ROW.match(label)
                           or label == ERP_ROW)
        if label and not is_data_row and _empty(b):
            # A block heading: a name in column A and nothing in column B.
            # Some headings carry growth-rate formulas across the row. A
            # data row still empty (a town just added) looks the same, so
            # the known row labels are excluded.
            block = label
            continue
        if not block or not label:
            continue
        name = loose_name(block)

        def as_years(series: dict) -> dict[int, float]:
            return {int(y): float(v) for y, v in series.items() if v is not None}

        match = QLD_LGA_ROW.match(label)
        if match:
            area = qld_lga.get(name)
            series = (area or {}).get(match.group(1).lower())
            if series:
                targets.append(Target(r, ("qld_lga", match.group(1).lower()), block, as_years(series)))
            else:
                notes.append(f"{block} (LGA): not in the QGSO 2018 projections; row {r} left as it is")
            continue
        if QLD_SA2_ROW.match(label):
            area = qld_sa2.get(name) or qld_sa2_code.get(sa2_code_of.get(name))
            if area:
                targets.append(Target(r, ("qld_sa2", "medium"), block, as_years(area["medium"])))
            else:
                notes.append(f"{block} (SA2): not in the QGSO 2018 projections; row {r} left as it is")
            continue
        match = NRW_ROW.match(label)
        if match:
            table = nrw.get(name)
            series_name = f"Series {match.group(1).upper()}"
            if table and series_name in table["series"]:
                values = as_years(table["series"][series_name])
                targets.append(Target(r, ("nrw", series_name), block, values,
                                      tuple(y for y in values if y > int(table["estimate_year"]))))
            elif table:
                notes.append(f"{block} (LGA): {series_name} not published for this LGA ({table['table']}); "
                             f"row {r} left empty. {table.get('not_published', '')}".strip())
            else:
                notes.append(f"{block} (LGA): not in the QGSO non-resident projections; row {r} left as it is")
            continue
        if label == ERP_ROW and section in ("lga", "sa2") and name in nsw_areas[section]:
            series = nsw_series[section].get(name)
            if series:
                targets.append(Target(r, ("nsw", section), block, as_years(series)))
            else:
                notes.append(f"{block} ({section.upper()}): not in the NSW projections")
    return targets, notes


def own_columns(years_by_col: dict[int, int]) -> dict[int, int]:
    """{year: column} from the year header."""
    return {y: c for c, y in years_by_col.items()}


def plan_moves(grid: list[list], targets: list["Target"], years_by_col: dict[int, int]):
    """([(row, from column, to column)], report lines): projections held
    under another year's column, to be moved to their own. See the module
    docstring for the rule."""
    column_of = own_columns(years_by_col)
    moves, lines = [], []
    for t in targets:
        if t.key[0] == "nsw":
            continue
        row = grid[t.row - 1]
        cells = [(c, _number(row[c - 1])) for c in sorted(years_by_col)
                 if c - 1 < len(row) and _number(row[c - 1]) is not None]
        expected = sorted(t.shown().items())
        if not cells:
            continue
        matches = (len(cells) == len(expected)
                   and all(abs(v - src) <= MATCH_TOLERANCE for (_, v), (_, src) in zip(cells, expected)))
        if not matches:
            if any(years_by_col[c] not in t.series for c, _ in cells):
                lines.append(f"  {t.area} (row {t.row}): NOT MOVED -- its figures do not match the source "
                             f"year by year; check by hand")
            continue
        planned = [(c, column_of.get(year), year) for (c, _), (year, _) in zip(cells, expected)
                   if years_by_col[c] != year]
        if not planned:
            continue
        if any(to is None for _, to, _ in planned):
            lines.append(f"  {t.area} (row {t.row}): NOT MOVED -- no column for "
                         f"{', '.join(str(y) for _, to, y in planned if to is None)}")
            continue
        sources = {c for c, _ in cells}
        blocked = [to for _, to, _ in planned if to not in sources and not _empty(row[to - 1] if to - 1 < len(row) else None)]
        if blocked:
            lines.append(f"  {t.area} (row {t.row}): NOT MOVED -- the column for "
                         f"{', '.join(str(years_by_col[b]) for b in blocked)} already holds another figure")
            continue
        moves += [(t.row, frm, to) for frm, to, _ in planned]
        lines.append(f"  {t.area} (row {t.row}): MOVED to own year: "
                     + ", ".join(f"{year} (was under {years_by_col[frm]})" for frm, _, year in planned))
    return moves, lines


def apply_moves(grid: list[list], moves) -> None:
    """Make the moves in the grid of values (all values read first, so a
    figure moved into a column another figure leaves is not lost)."""
    by_row = defaultdict(list)
    for r, frm, to in moves:
        by_row[r].append((frm, to))
    for r, pairs in by_row.items():
        row = grid[r - 1]
        width = max(len(row), max(to for _, to in pairs))
        row.extend([None] * (width - len(row)))
        values = {frm: row[frm - 1] for frm, _ in pairs}
        for frm, _ in pairs:
            row[frm - 1] = None
        for frm, to in pairs:
            row[to - 1] = values[frm]


def plan_writes(grid: list[list], data: dict, towns: list, last_data_year: int):
    """([(row, column, value)], [(row, from, to)], report lines) for the
    Population sheet, given as a grid of cell values (rows of 1-based
    columns). The writes are planned on the grid after the moves; the
    grid passed in is changed accordingly."""
    years_by_col = header_years(grid)
    column_of = own_columns(years_by_col)
    targets, notes = find_targets(grid, data, towns)
    lines = list(notes)

    moves, move_lines = plan_moves(grid, targets, years_by_col)
    apply_moves(grid, moves)
    lines += move_lines

    by_key: dict[tuple, list[Target]] = defaultdict(list)
    for t in targets:
        by_key[t.key].append(t)

    # New South Wales: the years the filled rows show, after the last data year.
    for key, group in by_key.items():
        if key[0] != "nsw":
            continue
        shown = {y for t in group for y, v in t.series.items()
                 if y > last_data_year and y in column_of
                 and _number(_cell(grid, t.row, column_of[y])) is not None}
        years = tuple(sorted(shown)) or tuple(y for y in NSW_DEFAULT_YEARS if y > last_data_year)
        for t in group:
            t.years = years

    writes = []
    for key, group in sorted(by_key.items()):
        filled = [_number(_cell(grid, t.row, c)) for t in group for c in years_by_col
                  if _number(_cell(grid, t.row, c)) is not None]
        whole = all(v.is_integer() for v in filled) if filled else key[0] != "qld_sa2"
        for t in group:
            done, kept, no_column = [], [], []
            for year, value in sorted(t.shown().items()):
                c = column_of.get(year)
                if c is None:
                    no_column.append(year)
                    continue
                current = _cell(grid, t.row, c)
                if _empty(current):
                    writes.append((t.row, c, round(value) if whole else value))
                    done.append(year)
                elif _number(current) is None or abs(_number(current) - value) > MATCH_TOLERANCE:
                    kept.append(f"{year}: sheet {current}, source {round(value, 1)}")
            if done:
                lines.append(f"  {t.area} (row {t.row}): WRITTEN {', '.join(map(str, done))}")
            if kept:
                lines.append(f"  {t.area} (row {t.row}): kept, differs from the source: {'; '.join(kept)}")
            if no_column:
                lines.append(f"  {t.area} (row {t.row}): no column for {', '.join(map(str, no_column))}")
    return writes, moves, lines


def _cell(grid, r: int, c: int):
    row = grid[r - 1]
    return row[c - 1] if c - 1 < len(row) else None


def write_projections(sheet, data: dict, towns: list, last_data_year: int) -> tuple[list[str], int]:
    """Plan and make the moves and writes on an xlwings sheet (or a
    stand-in with the same interface, tests/fake_xlwings_sheet.py).
    Returns (report lines, number of cells changed)."""
    last = sheet.used_range.last_cell
    grid = sheet.range((1, 1), (last.row, last.column)).value
    grid = [list(r) if isinstance(r, (list, tuple)) else [r] for r in grid]
    original = [list(r) for r in grid]
    writes, moves, lines = plan_writes(grid, data, towns, last_data_year)
    by_row = defaultdict(list)
    for r, frm, to in moves:
        by_row[r].append((frm, to))
    for r, pairs in by_row.items():
        formats = {frm: sheet.cells(r, frm).number_format for frm, _ in pairs}
        for frm, _ in pairs:
            sheet.cells(r, frm).value = None
        for frm, to in pairs:
            sheet.cells(r, to).value = original[r - 1][frm - 1]
            sheet.cells(r, to).number_format = formats[frm]
    for r, c, value in writes:
        if _empty(sheet.cells(r, c).value):
            sheet.cells(r, c).value = value
    if moves:
        lines.insert(0, f"{len(moves)} projections moved to their own year's column.")
    return lines, len(writes) + len(moves)


def update_population_projections(xlsx_path: Path, cache_dir: Path, visible: bool = False) -> list[str]:
    import xlwings as xw
    from config import YEAR_END, get_config

    cache_file = cache_dir / CACHE_FILE
    if not cache_file.exists():
        raise FileNotFoundError(f"{cache_file} not found -- run "
                                f"python run_update.py --only population_projections first.")
    with open(cache_file, encoding="utf-8") as f:
        data = json.load(f)
    towns = list(get_config().towns)

    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(str(xlsx_path))
        try:
            lines, written = write_projections(wb.sheets[SHEET_NAME], data, towns, YEAR_END)
            if written:
                wb.save()
        finally:
            wb.close()
    finally:
        app.quit()
    lines += ["", f"Summary: {written} cells written or moved."]
    return lines


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--visible"]
    if len(args) != 2:
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/population dir> [--visible]")
        sys.exit(1)
    for line in update_population_projections(Path(args[0]), Path(args[1]), visible="--visible" in sys.argv):
        print(line)
