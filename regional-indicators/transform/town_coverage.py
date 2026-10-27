"""
regional-indicators/transform/town_coverage.py

Lists what is still missing for a town in the indicators workbook: the
blocks it should have and does not, and the years left empty in the
blocks it has. Intended for checking a town added with add_town.py once
the writers have run, but it works for any town in towns.toml.

How "should have" is decided (no list of expected rows is kept)
  The workbook itself is the reference. A town's peers are the other
  towns of the same state that are in towns.toml.
  - Missing block: a section (Population SA2, Crime, Rainfall, ...)
    holds a block for a peer's area but none for the town's area at the
    same level (town, SA2 or LGA).
  - Empty year: in a block the town has, a cell is empty in a year for
    which most of the peers' blocks in the same section have a figure
    in the same row (rows are matched by label, as add_town matches
    them). A formula counts as a figure.
  Blocks of the same area shared by two towns (Mitchell and Roma are
  both in Maranoa LGA) are checked once, under each town.

  A gap can be one that no source fills (a peer's figures were entered
  by hand from a source that has since gone, or the source starts later
  for the town); the report says which peers have the figures, so it
  can be judged.

Also reported: whether the town has a chart page and lines in
web_export_map.csv.

Usage:
    python transform/town_coverage.py <workbook.xlsx>                   # every town with a page
    python transform/town_coverage.py <workbook.xlsx> --town Singleton --town Muswellbrook

Reads the workbook with openpyxl; changes nothing and needs no Excel.
"""

from __future__ import annotations

import argparse
import csv
import sys
from dataclasses import dataclass, field
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))

from add_town import (WEB_MAP, Unit, _match_rows, area_name, is_number, label_suffix, norm,
                      read_sheet_rows, read_units, unit_level)


# ── Reading years ──────────────────────────────────────────────────────────────

def year_rows(rows: list[tuple]) -> dict[int, dict[int, int]]:
    """{row: {column index: year}} for the sheet's year header rows: rows
    whose numbers are a run of at least three consecutive years."""
    found = {}
    for r, row in enumerate(rows, start=1):
        cells = {c: v for c, v in enumerate(row) if c > 0 and is_number(v)}
        values = list(cells.values())
        if (len(values) >= 3 and all(float(v).is_integer() and 1990 <= v <= 2100 for v in values)
                and all(b - a == 1 for a, b in zip(values, values[1:]))):
            found[r] = {c: int(v) for c, v in cells.items()}
    return found


def years_for(headers: dict[int, dict[int, int]], row: int) -> dict[int, int]:
    """The year of each column for a row: from the nearest year header at
    or above it."""
    above = [r for r in headers if r <= row]
    return headers[max(above)] if above else {}


def filled(value) -> bool:
    if is_number(value):
        return True
    return isinstance(value, str) and value.startswith("=")


def cell(rows, r: int, c: int):
    row = rows[r - 1]
    return row[c] if c < len(row) else None


def is_year_header(headers, r: int) -> bool:
    return r in headers


# ── The check ─────────────────────────────────────────────────────────────────

@dataclass
class Gap:
    sheet: str
    section: str
    block: str
    row_label: str
    years: list[int]
    peers: list[str]


@dataclass
class Coverage:
    town: str
    page: bool
    web_lines: int
    missing: list[str] = field(default_factory=list)
    gaps: list[Gap] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def unit_area(unit: Unit, towns) -> dict[str, list]:
    """The towns (from `towns`) whose area at the unit's level the unit is."""
    level = unit_level(unit)
    if level is None:
        return []
    return [t for t in towns if area_name(t, level) and norm(area_name(t, level)) == norm(unit.label)]


def row_label(rows, r: int, unit: Unit) -> str:
    row = rows[r - 1]
    a = row[0] if row else None
    b = row[1] if len(row) > 1 else None
    parts = [str(x).strip() for x in (a, b) if isinstance(x, str) and x.strip() and not x.startswith("=")]
    return " / ".join(parts) or f"row {r}"


def check_town(sheet_rows: dict[str, list[tuple]], town, towns, pages: list[str],
               web_rows: list[dict]) -> Coverage:
    result = Coverage(town.name, town.name in pages,
                      sum(1 for line in web_rows if norm(line.get("folder")) == norm(town.name)))
    peers_all = [t for t in towns if t.name != town.name and t.state == town.state]

    for sheet, rows in sheet_rows.items():
        units = read_units(sheet, rows)
        headers = year_rows(rows)
        sections = {}
        for unit in units:
            sections.setdefault(unit.section_no, []).append(unit)

        for section_units in sections.values():
            own = [u for u in section_units if town in unit_area(u, towns)]
            peer_units = [(u, [t for t in unit_area(u, towns) if t in peers_all]) for u in section_units]
            peer_units = [(u, ts) for u, ts in peer_units if ts and town not in unit_area(u, towns)]
            if not peer_units:
                continue
            section = section_units[0].section or ""
            where = f"{sheet} {section}".strip()
            if not own:
                level = unit_level(peer_units[0][0])
                name = area_name(town, level) if level else ""
                if not name:
                    result.notes.append(f"{where}: the town has no {level} name in towns.toml; not checked")
                else:
                    peer_names = sorted({u.label.strip() for u, _ in peer_units})
                    result.missing.append(f"{where}: no block for {name} "
                                          f"(the section has {', '.join(peer_names)})")
                continue

            for unit in own:
                peers = [u for u, _ in peer_units if u.kind == unit.kind]
                result.gaps += _block_gaps(sheet, section, rows, headers, unit, peers)
    return result


def peer_has(peer: Unit, peer_row: int, value) -> bool:
    """Whether a peer's cell holds a figure. On a block's name row only
    numbers count: some name rows carry working formulas (a growth rate,
    a reference to a cell) that are not data, while the New South Wales
    crime blocks carry the total rate there as a number."""
    if peer_row == peer.first and peer.kind == "block":
        return is_number(value)
    return filled(value)


def _block_gaps(sheet, section, rows, headers, unit: Unit, peers: list[Unit]) -> list[Gap]:
    # Each peer's rows matched to the town's rows (peer row -> town row).
    by_row: dict[int, list[tuple[Unit, int]]] = {}
    for peer in peers:
        for peer_row, own_row in _match_rows(rows, peer, unit).items():
            if own_row is not None:
                by_row.setdefault(own_row, []).append((peer, peer_row))

    gaps = []
    for r in range(unit.first, unit.last + 1):
        if is_year_header(headers, r) or r not in by_row:
            continue
        # Only peers with figures in the row count: a peer's row that is
        # empty throughout (another town just added) says nothing.
        matched = [(peer, peer_row) for peer, peer_row in by_row[r]
                   if any(peer_has(peer, peer_row, v) for v in rows[peer_row - 1][1:])]
        own_years = years_for(headers, r)
        missing, sources = [], set()
        for c, year in own_years.items():
            if filled(cell(rows, r, c)):
                continue
            having = []
            for peer, peer_row in matched:
                peer_years = years_for(headers, peer_row)
                peer_col = next((pc for pc, y in peer_years.items() if y == year), None)
                if peer_col is not None and peer_has(peer, peer_row, cell(rows, peer_row, peer_col)):
                    having.append(peer.label.strip())
            if having and len(having) * 2 >= len(matched):
                missing.append(year)
                sources.update(having)
        if missing:
            gaps.append(Gap(sheet, section, unit.label.strip(), row_label(rows, r, unit), missing, sorted(sources)))
    return gaps


# ── Report ────────────────────────────────────────────────────────────────────

def year_ranges(years: list[int]) -> str:
    years = sorted(years)
    out, start = [], None
    for i, y in enumerate(years):
        if start is None:
            start = y
        if i + 1 == len(years) or years[i + 1] != y + 1:
            out.append(str(start) if start == y else f"{start}-{y}")
            start = None
    return ", ".join(out)


def report(result: Coverage) -> list[str]:
    lines = [f"{result.town}"]
    lines.append(f"  chart page: {'yes' if result.page else 'NO'};  web_export_map.csv lines: {result.web_lines}")
    if not result.missing and not result.gaps:
        lines.append("  complete: no missing blocks and no empty years that the peers have")
    if result.missing:
        lines.append(f"  Missing blocks ({len(result.missing)}):")
        lines += [f"    - {m}" for m in result.missing]
    if result.gaps:
        lines.append(f"  Empty years ({len(result.gaps)} rows):")
        for g in result.gaps:
            where = f"{g.sheet} {g.section}".strip()
            lines.append(f"    - {where} | {g.block} | {g.row_label}: {year_ranges(g.years)}"
                         f"  (filled for {', '.join(g.peers)})")
    for note in result.notes:
        lines.append(f"  note: {note}")
    return lines


def read_web_map(path: Path = WEB_MAP) -> list[dict]:
    if not path.exists():
        return []
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="List the blocks and years still missing for a town.")
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--town", action="append", help="town to check (repeatable); default: every town with a page")
    args = parser.parse_args(argv)
    if not args.workbook.is_file():
        print(f"Workbook not found: {args.workbook}")
        return 2

    from config import get_config
    towns = list(get_config().towns)

    import openpyxl
    book = openpyxl.load_workbook(args.workbook, read_only=True)
    pages = book.sheetnames
    book.close()

    if args.town:
        chosen = []
        for name in args.town:
            match = [t for t in towns if norm(t.name) == norm(name)]
            if not match:
                print(f"{name} is not in towns.toml.")
                return 2
            chosen += match
    else:
        chosen = [t for t in towns if t.name in pages]

    sheet_rows = read_sheet_rows(args.workbook)
    web_rows = read_web_map()
    for town in chosen:
        print("\n".join(report(check_town(sheet_rows, town, towns, pages, web_rows))))
        print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
