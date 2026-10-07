"""
regional-indicators/transform/web_export.py

Builds the website CSV folder from the indicators workbook: one folder
per town, one CSV file per indicator.

The workbook is the system of record. It holds every series back to
2000, including years that cannot be re-fetched, and it is the source
for the charts and booklets, so the website is generated from it and
not from the fetcher cache.

The workbook is opened read-only with openpyxl and is never saved, so
the limitations described in xlsx_update/base.py do not apply and Excel
is not required. Values are read as last calculated and saved by Excel;
a mapped row whose cells have no saved values is reported.

Mapping: regional-indicators/web_export_map.csv, one line per output
file:
    folder     town folder under the output directory
    file       CSV name, without the ".csv" extension
    sheet      workbook sheet holding the series
    section    section of that sheet ("LGA", "SA2", "UCL", "State",
               "Rainfall", "Education", "Fuel", "Narrabri (LGA)"), or
               blank if the sheet is not divided into sections
    block      heading the row sits under (usually a town or region
               name), or blank if the row itself is labelled with the
               region
    row_label  column A text of the row to export, or "@next" for the
               row immediately below the block heading
    sub_label  column B text, where needed to distinguish two rows
    note       free text, ignored
Towns and indicators are defined entirely by this file.

Row resolution:
  1. The section starts at the first row whose column A equals the
     section name and ends at the next section heading.
  2. The block is the row in that section whose column A equals the
     block name.
  3. The row is the first one below the block heading whose column A
     equals row_label (and whose column B equals sub_label, if given).
     The scan stops at the next block heading named in the map, so a
     missing row is reported and never taken from another block.
Text is compared with surrounding whitespace removed. Rows are located
by label, not by row number.

Years: the nearest year header row above a series applies. A fiscal
label ("2023/24") is treated as the calendar year in which it ends
(2024). Every file has the same header, from FIRST_YEAR to the latest
year present in any exported series, with empty fields where a series
has no value. Values under years later than the cut-off (by default the
current calendar year; see --last-year) are excluded and reported, since
some rows hold projections in later columns.

Output format:
    2000,2001,...,2025<CR><LF>
    ,3545,3569,...,6565<CR><LF>
Two lines, unquoted, CRLF line endings. See format_number for number
formatting.

--compare-with <folder> compares the series to be exported with an
earlier export and reports, per file, the years whose values differ.
The earlier folder is read before any file is written. Earlier files
may be quoted, contain thousands separators or have trailing blank
lines.

Usage:
    python web_export.py <workbook.xlsx> [--out "3 Web Content"]
                         [--map web_export_map.csv]
                         [--compare-with <earlier export folder>]
                         [--last-year 2025] [--dry-run]
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl

PROJECT_DIR  = Path(__file__).resolve().parent.parent           # regional-indicators/
DEFAULT_MAP  = PROJECT_DIR / "web_export_map.csv"
DEFAULT_OUT  = PROJECT_DIR.parent / "3 Web Content"
FIRST_YEAR   = 2000
NEXT_ROW     = "@next"
SECTION_LABELS = {"LGA", "SA2", "UCL", "State", "Rainfall", "Education", "Fuel", "Narrabri (LGA)"}
MIN_YEARS_IN_HEADER = 8
FISCAL_RE = re.compile(r"^(\d{4})/(\d{2})$")


# ── Map ────────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class MapEntry:
    folder: str
    file: str
    sheet: str
    section: str
    block: str
    row_label: str
    sub_label: str
    line: int

    @property
    def name(self) -> str:
        return f"{self.folder}/{self.file}.csv"


def load_map(path: Path) -> list[MapEntry]:
    entries: list[MapEntry] = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        needed = {"folder", "file", "sheet", "section", "block", "row_label", "sub_label"}
        missing = needed - set(reader.fieldnames or [])
        if missing:
            raise ValueError(f"{path.name} is missing column(s): {sorted(missing)}")
        for line, row in enumerate(reader, start=2):
            if not (row["folder"] or "").strip():
                continue                                    # blank line
            entries.append(MapEntry(
                folder=row["folder"].strip(), file=row["file"].strip(),
                sheet=row["sheet"].strip(), section=(row["section"] or "").strip(),
                block=(row["block"] or "").strip(), row_label=(row["row_label"] or "").strip(),
                sub_label=(row["sub_label"] or "").strip(), line=line,
            ))
    seen: dict[tuple[str, str], int] = {}
    for e in entries:
        key = (e.folder.lower(), e.file.lower())
        if key in seen:
            raise ValueError(
                f"{path.name} lines {seen[key]} and {e.line} both produce {e.name} -- "
                f"each file can only come from one row."
            )
        seen[key] = e.line
    return entries


# ── Workbook access ────────────────────────────────────────────────────────

def _text(value) -> str:
    return str(value).strip() if value is not None else ""


def _is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _as_year(value) -> int | None:
    """Return the year a header cell represents, or None.

    2024 -> 2024; "2023/24" -> 2024.
    """
    if _is_number(value) and float(value).is_integer() and 1990 <= value <= 2100:
        return int(value)
    if isinstance(value, str):
        m = FISCAL_RE.match(value.strip())
        if m:
            return int(m.group(1)) + 1
    return None


class SheetGrid:
    """In-memory copy of a sheet's values with its year header rows indexed."""

    def __init__(self, ws):
        self.name = ws.title
        self.rows: list[tuple] = [tuple(r) for r in ws.iter_rows(values_only=True)]
        self.col_a = [_text(r[0]) if r else "" for r in self.rows]
        self.col_b = [_text(r[1]) if len(r) > 1 else "" for r in self.rows]
        # year header rows: {row index: {year: column index}}
        self.headers: dict[int, dict[int, int]] = {}
        for index, row in enumerate(self.rows):
            years = [(col, _as_year(v)) for col, v in enumerate(row)]
            years = [(col, y) for col, y in years if y is not None]
            # Keep the first run of strictly increasing years. Some sheets repeat
            # year labels further right, above derived columns.
            run: list[tuple[int, int]] = []
            for col, y in years:
                if run and y <= run[-1][1]:
                    break
                run.append((col, y))
            if len(run) >= MIN_YEARS_IN_HEADER:
                self.headers[index] = {y: col for col, y in run}

    def header_for(self, index: int) -> dict[int, int] | None:
        above = [h for h in self.headers if h <= index]
        return self.headers[max(above)] if above else None

    def series(self, index: int) -> tuple[dict[int, float], list[int]]:
        """Return ({year: value}, [years whose cell holds non-numeric text])
        for the row at `index`.
        """
        header = self.header_for(index)
        if header is None:
            raise LookupError(f"no year header row above row {index + 1} on sheet '{self.name}'")
        row = self.rows[index]
        values, odd = {}, []
        for year, col in header.items():
            value = row[col] if col < len(row) else None
            if _is_number(value):
                values[year] = value
            elif value is not None and _text(value) != "":
                odd.append(year)
        return values, odd


def find_row(grid: SheetGrid, entry: MapEntry, known_blocks: set[str]) -> int:
    """Return the 0-based row index for a map entry.

    Raises LookupError, with an explanation, if the row cannot be
    identified uniquely.
    """
    start, end = 0, len(grid.rows)
    if entry.section:
        starts = [i for i, a in enumerate(grid.col_a) if a == entry.section]
        if not starts:
            raise LookupError(f"no section '{entry.section}' on sheet '{grid.name}'")
        start = starts[0] + 1
        end = next((i for i in range(start, len(grid.rows)) if grid.col_a[i] in SECTION_LABELS),
                   len(grid.rows))
    else:
        end = next((i for i, a in enumerate(grid.col_a) if a in SECTION_LABELS), len(grid.rows))

    def matches(i: int) -> bool:
        return grid.col_a[i] == entry.row_label and (
            not entry.sub_label or grid.col_b[i] == entry.sub_label
        )

    where = f"sheet '{grid.name}'" + (f", section '{entry.section}'" if entry.section else "")

    if not entry.block:
        hits = [i for i in range(start, end) if matches(i)]
        if len(hits) == 1:
            return hits[0]
        if not hits:
            raise LookupError(f"no row labelled '{entry.row_label}' in {where}")
        raise LookupError(
            f"{len(hits)} rows labelled '{entry.row_label}' in {where} "
            f"(rows {[h + 1 for h in hits]}) -- give a block or sub_label"
        )

    blocks = [i for i in range(start, end) if grid.col_a[i] == entry.block]
    if not blocks:
        raise LookupError(f"no block headed '{entry.block}' in {where}")
    if len(blocks) > 1:
        raise LookupError(
            f"{len(blocks)} rows headed '{entry.block}' in {where} (rows {[b + 1 for b in blocks]})"
        )
    block_row = blocks[0]

    if entry.row_label == NEXT_ROW:
        if block_row + 1 >= end:
            raise LookupError(f"nothing below block '{entry.block}' in {where}")
        return block_row + 1

    stop_at = known_blocks - {entry.block}
    for i in range(block_row + 1, end):
        if matches(i):
            return i
        if grid.col_a[i] in stop_at and grid.col_a[i] != entry.row_label:
            break
    raise LookupError(
        f"no row labelled '{entry.row_label}'"
        + (f" / '{entry.sub_label}'" if entry.sub_label else "")
        + f" under block '{entry.block}' in {where}"
    )


# ── Formatting and files ───────────────────────────────────────────────────

MAX_DIGITS = 7


def format_number(value) -> str:
    """Format a number for the CSV files.

    Whole numbers are written without a decimal point. Other numbers
    are written with at most MAX_DIGITS digits in total, counting at
    least one digit before the decimal point (62.2454, 86.99918,
    440381.3, 0.682128), rounded half-up, without trailing zeros and
    never in exponent notation.

    This reproduces Excel's General number format at default column
    width, which is how values appeared in earlier exports.
    """
    if value is None:
        return ""
    number = Decimal(repr(float(value)))
    if number == number.to_integral_value():
        return str(int(number))
    integer_digits = max(1, len(str(abs(int(number)))))
    decimals = max(0, MAX_DIGITS - integer_digits)
    rounded = number.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
    text = format(rounded, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return "0" if text in ("-0", "") else text


def csv_text(years: list[int], values: dict[int, float]) -> str:
    header = ",".join(str(y) for y in years)
    data = ",".join(format_number(values.get(y)) for y in years)
    return f"{header}\r\n{data}\r\n"


def read_existing_csv(path: Path) -> dict[int, float]:
    """Read a CSV from an earlier export.

    Accepts quoted or unquoted values, thousands separators, padding
    and trailing blank lines. Returns {year: value} from the first data
    line.
    """
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = [r for r in csv.reader(f) if any(c.strip() for c in r)]
    if len(rows) < 2:
        return {}
    out = {}
    for year_text, cell in zip(rows[0], rows[1]):
        year_text, cell = year_text.strip(), cell.strip().replace(",", "").replace("$", "")
        if year_text.isdigit() and cell not in ("", "-"):
            try:
                out[int(year_text)] = float(cell)
            except ValueError:
                pass
    return out


def _same(a: float, b: float) -> bool:
    """True if two values are equal as formatted by format_number, or
    differ by no more than the precision of an earlier export.
    """
    return format_number(a) == format_number(b) or abs(a - b) <= max(0.0051, abs(b) * 1e-6)


# ── Export ─────────────────────────────────────────────────────────────────

def build_series(workbook_path: Path, entries: list[MapEntry]):
    """Resolve every map entry against the workbook.

    Returns ({entry: {year: value}}, [problems], [notes]).
    """
    wb = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        grids: dict[str, SheetGrid] = {}
        for sheet in sorted({e.sheet for e in entries}):
            if sheet in wb.sheetnames:
                grids[sheet] = SheetGrid(wb[sheet])
    finally:
        wb.close()

    known_blocks: dict[tuple[str, str], set[str]] = {}
    for e in entries:
        if e.block:
            known_blocks.setdefault((e.sheet, e.section), set()).add(e.block)

    series: dict[MapEntry, dict[int, float]] = {}
    problems: list[str] = []
    notes: list[str] = []
    for e in entries:
        grid = grids.get(e.sheet)
        if grid is None:
            problems.append(f"{e.name}: workbook has no sheet '{e.sheet}' (map line {e.line})")
            continue
        try:
            index = find_row(grid, e, known_blocks.get((e.sheet, e.section), set()))
            values, odd = grid.series(index)
        except LookupError as exc:
            problems.append(f"{e.name}: {exc} (map line {e.line})")
            continue
        if not values:
            problems.append(
                f"{e.name}: {e.sheet} row {index + 1} has no numbers under any year "
                f"(map line {e.line}) -- formulas without saved results? Open and save "
                f"the workbook in Excel."
            )
            continue
        if odd:
            notes.append(f"{e.name}: {e.sheet} row {index + 1} has text, not a number, for {odd}")
        series[e] = values
    return series, problems, notes


def export(workbook_path: Path, out_dir: Path, map_path: Path = DEFAULT_MAP,
           compare_with: Path | None = None, dry_run: bool = False,
           last_year: int | None = None) -> tuple[list[str], int]:
    """Write the export folder. Returns (report lines, number of problems)."""
    entries = load_map(map_path)
    series, problems, notes = build_series(workbook_path, entries)

    report = [f"Workbook : {workbook_path}", f"Map      : {map_path.name} ({len(entries)} files)"]
    if not series:
        report += ["", "Nothing could be exported:"] + [f"  {p}" for p in problems]
        return report, len(problems)

    # Values under years later than the cut-off are not exported. Some rows
    # hold projections in later columns, and observed data cannot be dated
    # later than the current year, so the default cut-off is the current
    # calendar year.
    cutoff = last_year if last_year is not None else datetime.now().year
    dropped = {
        e.name: sorted(y for y in v if y > cutoff) for e, v in series.items()
        if any(y > cutoff for y in v)
    }
    series = {e: {y: x for y, x in v.items() if y <= cutoff} for e, v in series.items()}
    for e in [e for e, v in series.items() if not v]:
        problems.append(f"{e.name}: no figures up to {cutoff} (map line {e.line})")
        del series[e]
    for name, yrs in sorted(dropped.items()):
        notes.append(f"{name}: left out figures dated after {cutoff} (projections): {yrs}")

    last_year = max(max(v) for v in series.values())
    years = list(range(FIRST_YEAR, last_year + 1))
    report.append(f"Years    : {FIRST_YEAR}-{last_year} in every file")

    written = 0
    for e, values in series.items():
        if not dry_run:
            folder = out_dir / e.folder
            folder.mkdir(parents=True, exist_ok=True)
            with open(folder / f"{e.file}.csv", "w", newline="", encoding="utf-8") as f:
                f.write(csv_text(years, values))
        written += 1
    verb = "Would write" if dry_run else "Wrote"
    folders = len({e.folder for e in series})
    report.append(f"{verb}    : {written} files in {folders} town folders under {out_dir}")

    with_latest = sum(1 for v in series.values() if last_year in v)
    report.append(f"           {with_latest} have a {last_year} figure, {written - with_latest} do not")
    by_file: dict[str, list[str]] = {}
    for e, values in series.items():
        if last_year not in values:
            by_file.setdefault(e.file, []).append(f"{e.folder} (to {max(values)})")
    for file, towns in sorted(by_file.items()):
        report.append(f"             no {last_year}: {file} — {', '.join(sorted(towns))}")

    if notes:
        report += ["", "Notes:"] + [f"  {n}" for n in notes]
    if problems:
        report += ["", f"PROBLEMS ({len(problems)}) -- these files were NOT written:"]
        report += [f"  {p}" for p in problems]

    if compare_with is not None:
        report += [""] + compare(series, compare_with)

    return report, len(problems)


def compare(series: dict[MapEntry, dict[int, float]], old_dir: Path) -> list[str]:
    """Compare the series to be exported with an earlier export folder."""
    lines = [f"Comparison with {old_dir}:"]
    identical = revised = extended = missing_old = 0
    detail: list[str] = []
    for e, values in sorted(series.items(), key=lambda kv: kv[0].name):
        old_path = old_dir / e.folder / f"{e.file}.csv"
        if not old_path.exists():
            missing_old += 1
            detail.append(f"  NEW   {e.name}: not in the earlier export")
            continue
        old = read_existing_csv(old_path)
        diffs = [y for y in sorted(old) if y in values and not _same(values[y], old[y])]
        dropped = [y for y in sorted(old) if y not in values]
        added = [y for y in sorted(values) if y not in old and y >= FIRST_YEAR]
        if diffs or dropped:
            revised += 1
            parts = []
            if diffs:
                shown = ", ".join(
                    f"{y}: {format_number(old[y])} -> {format_number(values[y])}" for y in diffs[:4]
                )
                parts.append(f"{len(diffs)} year(s) differ ({shown}{', ...' if len(diffs) > 4 else ''})")
            if dropped:
                parts.append(f"earlier file had figures for {dropped} that the workbook lacks")
            detail.append(f"  DIFF  {e.name}: " + "; ".join(parts))
        elif added:
            extended += 1
        else:
            identical += 1
    old_files = {
        (p.parent.name.lower(), p.stem.lower())
        for p in old_dir.glob("*/*.csv")
    }
    ours = {(e.folder.lower(), e.file.lower()) for e in series}
    gone = sorted(old_files - ours)
    lines.append(
        f"  {identical} identical, {extended} same history plus new year(s), "
        f"{revised} with differences, {missing_old} new, {len(gone)} in the earlier export only"
    )
    lines += detail
    lines += [f"  GONE  {folder}/{file}.csv: in the earlier export, not in this one" for folder, file in gone]
    return lines


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Build the website CSV folder from the indicators workbook."
    )
    parser.add_argument("workbook", type=Path, help="the finished indicators workbook (.xlsx)")
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT,
                        help=f'output folder (default: "{DEFAULT_OUT.name}" beside regional-indicators)')
    parser.add_argument("--map", type=Path, default=DEFAULT_MAP, dest="map_path",
                        help="map file (default: regional-indicators/web_export_map.csv)")
    parser.add_argument("--compare-with", type=Path, default=None,
                        help="an earlier export folder to compare against")
    parser.add_argument("--dry-run", action="store_true", help="report only, write nothing")
    parser.add_argument("--last-year", type=int, default=None,
                        help="last year to export (default: this calendar year; "
                             "later columns hold projections)")
    args = parser.parse_args(argv)

    if not args.workbook.exists():
        print(f"Workbook not found: {args.workbook}")
        return 2
    report, problems = export(
        args.workbook, args.out, args.map_path, args.compare_with, args.dry_run,
        args.last_year,
    )
    for line in report:
        print(line)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
