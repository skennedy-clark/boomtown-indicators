"""
regional-indicators/transform/xlsx_update/update_population_projections.py

Writes population projections into the Population sheet.

Input:  cache/population/population_projections.json, from
        fetchers/fetch_population_projections.py.
Target: the Population sheet.

Rows written
  Queensland, QGSO 2018 edition (the projection sections of the sheet):
    "Projected Population - High|Low|Medium series: Persons Total
     (Persons), 2018ed"            in each LGA's block, from the LGA table;
    "Population projection: district 2018e (Medium Series)"
                                   in each SA2's block, from the SA2 table.
  New South Wales, NSW Planning projections: the "Population (ERP)" row
    of the LGA and SA2 blocks of New South Wales towns (towns.toml), in
    columns after the last data year, as Narrabri's rows hold them.

  A block is matched to a source area by name, compared loosely ("Isaac
  (R)" = "Isaac", "Broadsound - Nebo" = "Broadsound-Nebo"). An SA2 renamed
  since 2016 is found by its code, through the town in towns.toml whose
  SA2 it is (Roma Surrounds was "Roma Region" in the 2018 edition).

Columns
  The sheet does not put each projection under its own year: the 2018
  edition's 2016, 2021 and 2026 figures sit in the columns headed 2011,
  2017 and 2025. The columns are therefore taken from the rows already
  filled: for each row type and projection year, the column in which the
  other areas' rows hold that year's figure. Only the years those rows
  show are written (for New South Wales, 2031, 2036 and 2041). Where no
  row of the type has figures yet, each year goes under its own year,
  limited to the columns the other projection rows use (2031, 2036,
  2041) where those are among its years.
  Whether figures are rounded is also taken from the filled rows (the
  LGA rows hold whole numbers, the SA2 rows the source's decimals).

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
from collections import Counter, defaultdict
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
ERP_ROW = "Population (ERP)"
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

    targets, notes = [], []
    section = block = None
    for r, row in enumerate(grid, start=1):
        a = row[0] if row else None
        label = a.strip() if isinstance(a, str) else ""
        if label in SECTIONS:
            section, block = SECTIONS[label], None
            continue
        b = row[1] if len(row) > 1 else None
        is_data_row = bool(QLD_LGA_ROW.match(label) or QLD_SA2_ROW.match(label) or label == ERP_ROW)
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
        if label == ERP_ROW and section in ("lga", "sa2") and name in nsw_areas[section]:
            series = nsw_series[section].get(name)
            if series:
                targets.append(Target(r, ("nsw", section), block, as_years(series)))
            else:
                notes.append(f"{block} ({section.upper()}): not in the NSW projections")
    return targets, notes


def column_map(grid, targets: list[Target], years_by_col: dict[int, int],
               candidates: set[int]) -> tuple[dict[int, int], bool, bool]:
    """For one row type: ({projection year: column}, whether the filled
    rows hold whole numbers, whether the map came from filled rows)."""
    votes: dict[int, Counter] = defaultdict(Counter)
    whole = []
    for t in targets:
        row = grid[t.row - 1]
        for year, value in t.series.items():
            for c in candidates:
                cell = _number(row[c - 1]) if c - 1 < len(row) else None
                if cell is not None and abs(cell - value) <= MATCH_TOLERANCE:
                    votes[year][c] += 1
                    whole.append(cell.is_integer())
    if votes:
        mapping = {}
        for year, counter in votes.items():
            best = max(counter.values())
            tied = [c for c, n in counter.items() if n == best]
            mapping[year] = next((c for c in tied if years_by_col.get(c) == year), min(tied))
        return mapping, bool(whole) and all(whole), True
    by_year = {y: c for c, y in years_by_col.items() if c in candidates}
    years = {y for t in targets for y in t.series}
    return {y: by_year[y] for y in years if y in by_year}, True, False


def plan_writes(grid: list[list], data: dict, towns: list, last_data_year: int):
    """([(row, column, value)], report lines) for the Population sheet,
    given as a grid of cell values (rows of 1-based columns)."""
    years_by_col = header_years(grid)
    targets, notes = find_targets(grid, data, towns)
    lines = list(notes)
    writes = []
    by_key: dict[tuple, list[Target]] = defaultdict(list)
    for t in targets:
        by_key[t.key].append(t)
    def candidates_for(key):
        if key[0] == "nsw":
            return {c for c, y in years_by_col.items() if y > last_data_year}
        return set(years_by_col)

    maps = {key: column_map(grid, group, years_by_col, candidates_for(key)) for key, group in by_key.items()}
    # A row type with no filled rows yet shows the future years that the
    # other row types show (the sheet's projection columns: 2031, 2036,
    # 2041), each under its own year; failing those, every year.
    shown_cols = {c for mapping, _, from_rows in maps.values() if from_rows for c in mapping.values()}
    for key, group in sorted(by_key.items()):
        mapping, whole, from_rows = maps[key]
        if not from_rows:
            future = {y: c for y, c in mapping.items() if c in shown_cols and years_by_col.get(c) == y}
            mapping = future or mapping
        shown = ", ".join(f"{y}->{years_by_col.get(c, c)}" for y, c in sorted(mapping.items()))
        lines.append(f"{key[0]} {key[1]}: years -> columns headed {shown}"
                     + ("" if from_rows else " (no filled rows of this type to follow)"))
        for t in group:
            row = grid[t.row - 1]
            done, kept = [], []
            for year, c in sorted(mapping.items()):
                if year not in t.series:
                    continue
                value = round(t.series[year]) if whole else t.series[year]
                current = row[c - 1] if c - 1 < len(row) else None
                if _empty(current):
                    writes.append((t.row, c, value))
                    done.append(year)
                elif _number(current) is None or abs(_number(current) - t.series[year]) > MATCH_TOLERANCE:
                    kept.append(f"{year}: sheet {current}, source {round(t.series[year], 1)}")
            if done:
                lines.append(f"  {t.area} (row {t.row}): WRITTEN {', '.join(map(str, done))}")
            if kept:
                lines.append(f"  {t.area} (row {t.row}): kept, differs from the source: {'; '.join(kept)}")
    return writes, lines


def write_projections(sheet, data: dict, towns: list, last_data_year: int) -> tuple[list[str], int]:
    """Plan and make the writes on an xlwings sheet (or a stand-in with the
    same interface, tests/fake_xlwings_sheet.py)."""
    last = sheet.used_range.last_cell
    grid = sheet.range((1, 1), (last.row, last.column)).value
    grid = [list(r) if isinstance(r, (list, tuple)) else [r] for r in grid]
    writes, lines = plan_writes(grid, data, towns, last_data_year)
    for r, c, value in writes:
        if _empty(sheet.cells(r, c).value):
            sheet.cells(r, c).value = value
    return lines, len(writes)


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
    lines += ["", f"Summary: {written} cells written."]
    return lines


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--visible"]
    if len(args) != 2:
        print(f"Usage: python {sys.argv[0]} <xlsx_path> <cache/population dir> [--visible]")
        sys.exit(1)
    for line in update_population_projections(Path(args[0]), Path(args[1]), visible="--visible" in sys.argv):
        print(line)
