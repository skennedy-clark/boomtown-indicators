"""
regional-indicators/transform/chart_audit.py

Audits the data ranges of every chart series in the indicators workbook.

Each chart series plots one row of a data sheet (Population, Employment,
Housing, Income, Business, Crime, Exogenous) across a range of year
columns. After the annual update a series needs its range extended when
the row holds a figure for a year beyond the end of the range. This
script lists every series, the years its range covers and the years the
row has data for, and classifies it. It changes nothing.

The workbook is read directly (chart definitions from the .xlsx package,
cell values with openpyxl), so Excel is not required.

Status of a series:
    OK              the range ends at the last year of the row's
                    continuous run of data
    EXTEND          the columns immediately after the range hold data;
                    a new range is proposed
    PROJECTION      as EXTEND, but the range already ends after the
                    cut-off year, or the proposed range would (a row of
                    projections); listed for review, no change proposed
    ENDS_PAST_DATA  the last cell of the range is empty
    MULTI_AREA      the range is made of several separate areas. When
                    all of them lie on one row, a single range is
                    proposed: from the start of the first area to the
                    end of the continuous run of data that begins there
    UNRESOLVED      the range could not be interpreted (not a single
                    row, no year header found above it, or a sheet that
                    is not in the workbook)

How years are found: the year header for a data row is the nearest row
above it that holds a run of at least five consecutive years, either as
calendar years (2025) or as financial-year labels ("2024/25", read as
2025). Sections with their own header row (for example Fuel and
Education on the Exogenous sheet) are therefore read against that header
and not against the top of the sheet.

How the proposed range is found: from the last cell of the range, the
run of numeric cells continues to the right until the first empty cell.
Data that resumes after a gap is mentioned in the note and is not
included. The category (axis label) range, when it lies on the same
columns as the values, is moved to the same end column.

With --reference <workbook>, the same audit is run on a second workbook
and each series is shown with the range it has there. Series are paired
by page, chart position on the page and series order.

Output: a CSV file (default chart_audit.csv), one line per series, and
a summary on the console.

Usage:
    python chart_audit.py <workbook.xlsx> [--out chart_audit.csv]
                          [--last-year 2025]
                          [--reference <other workbook.xlsx>]
"""

from __future__ import annotations

import argparse
import csv
import posixpath
import re
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

NS = {
    "c":   "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m":   "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
DRAWING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"

MIN_YEARS_IN_HEADER = 5
HEADER_SEARCH_ROWS = 400

STATUSES = ("OK", "EXTEND", "PROJECTION", "ENDS_PAST_DATA", "MULTI_AREA", "UNRESOLVED")

CSV_COLUMNS = [
    "page", "chart_at", "chart_title", "series_no", "series_name",
    "data_sheet", "row", "row_label", "sub_label", "block",
    "values", "categories", "range_from_year", "range_to_year",
    "data_to_year", "status", "proposed_values", "proposed_categories",
    "reference_values", "note",
]


# ── Workbook package ───────────────────────────────────────────────────────────

@dataclass
class Area:
    """One rectangular area of a range reference."""
    sheet: str
    col1: int
    row1: int
    col2: int
    row2: int

    @property
    def single_row(self) -> bool:
        return self.row1 == self.row2

    def with_end_column(self, col: int) -> str:
        return format_area(self.sheet, self.col1, self.row1, col, self.row2)


@dataclass
class Series:
    page: str
    chart_at: str               # top-left cell of the chart on its page
    chart_order: tuple          # (row, column) of that cell, for sorting
    chart_title: str
    series_no: int
    series_name: str
    values: str                 # reference as stored in the chart
    categories: str
    # filled in by assess()
    data_sheet: str = ""
    row: int | None = None
    row_label: str = ""
    sub_label: str = ""
    block: str = ""
    range_from_year: int | None = None
    range_to_year: int | None = None
    data_to_year: int | None = None
    status: str = ""
    proposed_values: str = ""
    proposed_categories: str = ""
    reference_values: str = ""
    notes: list[str] = field(default_factory=list)

    def as_row(self) -> dict:
        row = {name: getattr(self, name, "") for name in CSV_COLUMNS if name != "note"}
        row["note"] = "; ".join(self.notes)
        return {k: ("" if v is None else v) for k, v in row.items()}


def _relationships(package: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    """Return {relationship id: (type, target part)} for a package part."""
    folder, name = posixpath.split(part)
    rels_part = posixpath.join(folder, "_rels", name + ".rels")
    if rels_part not in package.namelist():
        return {}
    out = {}
    for rel in ET.fromstring(package.read(rels_part)).findall("rel:Relationship", NS):
        target = rel.get("Target")
        if rel.get("TargetMode") == "External":
            continue
        resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
        out[rel.get("Id")] = (rel.get("Type"), resolved)
    return out


def _text(element, path: str) -> str:
    if element is None:
        return ""
    return "".join(node.text or "" for node in element.findall(path, NS)).strip()


def _reference(series_element, *paths: str) -> str:
    for path in paths:
        node = series_element.find(path, NS)
        if node is not None and node.text:
            return node.text.strip()
    return ""


def read_chart_series(workbook_path: Path) -> list[Series]:
    """Return every chart series in the workbook, in page and chart order."""
    series_list: list[Series] = []
    with zipfile.ZipFile(workbook_path) as package:
        workbook_rels = _relationships(package, "xl/workbook.xml")
        workbook = ET.fromstring(package.read("xl/workbook.xml"))
        for sheet in workbook.findall("m:sheets/m:sheet", NS):
            page = sheet.get("name")
            sheet_part = workbook_rels[sheet.get(f"{{{NS['r']}}}id")][1]
            for rel_type, drawing_part in _relationships(package, sheet_part).values():
                if rel_type != DRAWING_REL:
                    continue
                drawing_rels = _relationships(package, drawing_part)
                drawing = ET.fromstring(package.read(drawing_part))
                for anchor in list(drawing):
                    chart_ref = anchor.find(".//c:chart", NS)
                    origin = anchor.find("xdr:from", NS)
                    if chart_ref is None:
                        continue
                    rel_id = chart_ref.get(f"{{{NS['r']}}}id")
                    if rel_id not in drawing_rels or drawing_rels[rel_id][0] != CHART_REL:
                        continue
                    if origin is not None:
                        col = int(origin.findtext("xdr:col", "0", NS)) + 1
                        row = int(origin.findtext("xdr:row", "0", NS)) + 1
                    else:
                        col, row = 1, 1
                    chart = ET.fromstring(package.read(drawing_rels[rel_id][1]))
                    title = _text(chart.find("c:chart/c:title", NS), ".//a:t")
                    for number, ser in enumerate(chart.findall(".//c:ser", NS), start=1):
                        name = _text(ser.find("c:tx", NS), ".//c:v")
                        series_list.append(Series(
                            page=page,
                            chart_at=f"{get_column_letter(col)}{row}",
                            chart_order=(row, col),
                            chart_title=title,
                            series_no=number,
                            series_name=name,
                            values=_reference(ser, "c:val/c:numRef/c:f", "c:yVal/c:numRef/c:f"),
                            categories=_reference(
                                ser, "c:cat/c:numRef/c:f", "c:cat/c:strRef/c:f",
                                "c:cat/c:multiLvlStrRef/c:f", "c:xVal/c:numRef/c:f", "c:xVal/c:strRef/c:f",
                            ),
                        ))
    return series_list


# ── Range references ───────────────────────────────────────────────────────────

_AREA = re.compile(
    r"^(?:'((?:[^']|'')+)'|([^'!,()]+))!"
    r"\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$"
)


def parse_reference(reference: str) -> list[Area] | None:
    """Split a chart reference into its areas, or return None if any
    part is not a plain sheet range (for example a defined name)."""
    text = reference.strip()
    if text.startswith("(") and text.endswith(")"):
        text = text[1:-1]
    if not text:
        return None
    areas = []
    for part in _split_areas(text):
        match = _AREA.match(part.strip())
        if not match:
            return None
        quoted, plain, col1, row1, col2, row2 = match.groups()
        sheet = quoted.replace("''", "'") if quoted is not None else plain
        c1 = column_index_from_string(col1.upper())
        c2 = column_index_from_string(col2.upper()) if col2 else c1
        r1 = int(row1)
        r2 = int(row2) if row2 else r1
        areas.append(Area(sheet, min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)))
    return areas


def _split_areas(text: str) -> list[str]:
    """Split on commas that are outside quoted sheet names."""
    parts, current, quoted = [], [], False
    for char in text:
        if char == "'":
            quoted = not quoted
        if char == "," and not quoted:
            parts.append("".join(current))
            current = []
        else:
            current.append(char)
    parts.append("".join(current))
    return parts


def format_area(sheet: str, col1: int, row1: int, col2: int, row2: int) -> str:
    name = sheet if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", sheet) else "'" + sheet.replace("'", "''") + "'"
    start = f"${get_column_letter(col1)}${row1}"
    end = f"${get_column_letter(col2)}${row2}"
    return f"{name}!{start}" if start == end else f"{name}!{start}:{end}"


# ── Cell values ────────────────────────────────────────────────────────────────

class Cells:
    """Cached values of the data sheets, read once with openpyxl."""

    def __init__(self, workbook_path: Path, sheets: set[str]):
        book = openpyxl.load_workbook(workbook_path, data_only=True, read_only=True)
        self.sheets: dict[str, list[tuple]] = {}
        for name in sheets:
            if name in book.sheetnames:
                self.sheets[name] = [row for row in book[name].iter_rows(values_only=True)]
        book.close()

    def has(self, sheet: str) -> bool:
        return sheet in self.sheets

    def value(self, sheet: str, row: int, col: int):
        rows = self.sheets[sheet]
        if row < 1 or row > len(rows):
            return None
        cells = rows[row - 1]
        return cells[col - 1] if 1 <= col <= len(cells) else None

    def width(self, sheet: str, row: int) -> int:
        rows = self.sheets[sheet]
        return len(rows[row - 1]) if 1 <= row <= len(rows) else 0


def is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


_FINANCIAL_YEAR = re.compile(r"^\s*((?:19|20)\d{2})\s*[/-]\s*(\d{2}|\d{4})\s*$")


def as_year(value) -> int | None:
    """Return the year a header cell stands for, or None.

    Accepts a calendar year (2025) and a financial-year label
    ("2024/25" or "2024-25"), which counts as the year it ends in.
    """
    if is_number(value) and float(value).is_integer() and 1990 <= value <= 2100:
        return int(value)
    if isinstance(value, str):
        match = _FINANCIAL_YEAR.match(value)
        if match:
            start = int(match.group(1))
            end = int(match.group(2))
            if end == (start + 1) % 100 or end == start + 1:
                return start + 1
    return None


def is_year_row(cells: "Cells", sheet: str, row: int, columns) -> bool:
    """True if the row holds a run of at least MIN_YEARS_IN_HEADER
    consecutive years in adjacent cells of `columns`. Requiring a run
    excludes data rows whose values happen to look like years.
    """
    run, previous = 0, None
    for col in columns:
        year = as_year(cells.value(sheet, row, col))
        if year is not None and previous is not None and year == previous + 1:
            run += 1
        else:
            run = 1 if year is not None else 0
        if run >= MIN_YEARS_IN_HEADER:
            return True
        previous = year
    return False


def find_year_header(cells: Cells, sheet: str, row: int, col1: int, col2: int) -> int | None:
    """Return the nearest row above `row` that is a year header.

    A row qualifies if it holds a run of consecutive years (see
    is_year_row), looked for in the range's own columns when the range
    is wide enough and across the whole row otherwise. Where a sheet has
    a row of financial-year labels above a row of calendar years, the
    lower row is the one returned.
    """
    span = col2 - col1 + 1
    for candidate in range(row - 1, max(0, row - 1 - HEADER_SEARCH_ROWS), -1):
        if span >= MIN_YEARS_IN_HEADER:
            columns = range(col1, col2 + 1)
        else:
            columns = range(1, cells.width(sheet, candidate) + 1)
        if is_year_row(cells, sheet, candidate, columns):
            return candidate
    return None


def find_block(cells: Cells, sheet: str, row: int) -> str:
    """Return the heading of the block a data row belongs to: the nearest
    row above with text in column A and no numbers in the row."""
    for candidate in range(row - 1, max(0, row - 60), -1):
        label = cells.value(sheet, candidate, 1)
        if not isinstance(label, str) or not label.strip():
            continue
        width = cells.width(sheet, candidate)
        if not any(is_number(cells.value(sheet, candidate, col)) for col in range(2, width + 1)):
            return label.strip()
    return ""


# ── Assessment ─────────────────────────────────────────────────────────────────

def assess(series: Series, cells: Cells, last_year: int) -> None:
    """Classify one series and fill in its descriptive fields."""
    areas = parse_reference(series.values)
    if not areas:
        series.status = "UNRESOLVED"
        series.notes.append("values reference is not a sheet range")
        return

    area = areas[-1]
    series.data_sheet = area.sheet
    if not cells.has(area.sheet):
        series.status = "UNRESOLVED"
        series.notes.append(f"sheet '{area.sheet}' is not in the workbook")
        return
    if not area.single_row:
        series.status = "UNRESOLVED"
        series.notes.append("values range is not a single row")
        return

    row = area.row1
    series.row = row
    label = cells.value(area.sheet, row, 1)
    sub_label = cells.value(area.sheet, row, 2)
    series.row_label = str(label).strip() if label is not None else ""
    series.sub_label = str(sub_label).strip() if isinstance(sub_label, str) else ""
    series.block = find_block(cells, area.sheet, row)

    header = find_year_header(cells, area.sheet, row, area.col1, area.col2)
    if header is None:
        series.status = "UNRESOLVED"
        series.notes.append("no year header found above the row")
        return

    same_row = all(a.sheet == area.sheet and a.single_row and a.row1 == row for a in areas)
    multi = len(areas) > 1
    if multi and not same_row:
        series.status = "MULTI_AREA"
        series.notes.append(f"{len(areas)} separate areas on different rows or sheets")
        return
    if multi:
        # Assess the first area: the proposal is one range running from
        # its start to the end of the data that continues from its end.
        areas = sorted(areas, key=lambda a: a.col1)
        area = areas[0]

    def year_at(col: int) -> int | None:
        return as_year(cells.value(area.sheet, header, col))

    series.range_from_year = year_at(area.col1)
    series.range_to_year = year_at(areas[-1].col2)

    # Continuous run of data to the right of the range.
    end = area.col2
    width = max(cells.width(area.sheet, row), cells.width(area.sheet, header))
    run_end = end
    while run_end + 1 <= width and is_number(cells.value(area.sheet, row, run_end + 1)) \
            and year_at(run_end + 1) is not None:
        run_end += 1
    later = [
        col for col in range(run_end + 2, width + 1)
        if is_number(cells.value(area.sheet, row, col)) and year_at(col) is not None
    ]
    if later:
        series.notes.append(
            f"more data after a gap, from {year_at(later[0])} "
            f"({get_column_letter(later[0])}{row}), not included"
        )

    end_has_data = is_number(cells.value(area.sheet, row, end))
    if not end_has_data:
        last_filled = next(
            (col for col in range(end, area.col1 - 1, -1) if is_number(cells.value(area.sheet, row, col))), None
        )
        series.data_to_year = year_at(last_filled) if last_filled else None
        status = "ENDS_PAST_DATA"
        if last_filled:
            series.notes.append(
                f"range runs to {year_at(end)}; the row's last figure in the range is "
                f"{series.data_to_year} ({get_column_letter(last_filled)}{row})"
            )
        else:
            series.notes.append("the row has no figures in the range")
    else:
        series.data_to_year = year_at(run_end)
        if run_end == end:
            status = "OK"
        elif (year_at(end) or 0) > last_year or (series.data_to_year or 0) > last_year:
            status = "PROJECTION"
            series.notes.append(
                f"data continues to {series.data_to_year}, beyond the cut-off year {last_year}"
            )
        else:
            status = "EXTEND"

    if multi:
        series.status = "MULTI_AREA"
        described = ", ".join(
            f"{get_column_letter(a.col1)}-{get_column_letter(a.col2)}" if a.col1 != a.col2 else get_column_letter(a.col1)
            for a in areas
        )
        series.notes.insert(0, f"{len(areas)} separate areas (columns {described})")
        if status in ("OK", "EXTEND"):
            series.proposed_values = area.with_end_column(run_end)
            series.proposed_categories = _single_category_range(series, area, run_end)
            outside = [
                a for a in areas[1:]
                if a.col1 > run_end and any(is_number(cells.value(a.sheet, row, c)) for c in range(a.col1, a.col2 + 1))
            ]
            if outside:
                series.notes.append("an area beyond the proposed range holds data")
        return

    series.status = status
    category_note = _check_categories(series, area, header, run_end, status == "EXTEND", cells)
    if status == "EXTEND":
        series.proposed_values = area.with_end_column(run_end)
    if category_note:
        series.notes.append(category_note)


def _single_category_range(series: Series, area: Area, run_end: int) -> str:
    """For a multi-area series: the category range rebuilt as one range
    over the proposed value columns, when its first area starts on the
    same column as the values."""
    cat_areas = parse_reference(series.categories) if series.categories else None
    if not cat_areas:
        return ""
    cat_areas.sort(key=lambda a: a.col1)
    first = cat_areas[0]
    if not first.single_row or first.col1 != area.col1:
        return ""
    if any(a.sheet != first.sheet or a.row1 != first.row1 or not a.single_row for a in cat_areas):
        return ""
    return first.with_end_column(run_end)


def _check_categories(series: Series, area: Area, header: int, run_end: int, propose: bool, cells: Cells) -> str:
    """Compare the category range with the values range, and propose a
    new category range when the values range is being extended."""
    if not series.categories:
        return "no category range"
    cat_areas = parse_reference(series.categories)
    if not cat_areas:
        return "category reference is not a sheet range"
    if len(cat_areas) > 1:
        return "category range has several areas"
    cat = cat_areas[0]
    if cat.col1 != area.col1 or cat.col2 < area.col2:
        return (
            f"category range covers columns {get_column_letter(cat.col1)}-{get_column_letter(cat.col2)}, "
            f"values cover {get_column_letter(area.col1)}-{get_column_letter(area.col2)}"
        )
    if cat.col2 > area.col2:
        # A wider axis is deliberate (room for projections): keep it,
        # unless the extended values would run past it.
        if propose and run_end > cat.col2:
            series.proposed_categories = cat.with_end_column(run_end)
        return f"axis runs to column {get_column_letter(cat.col2)}, beyond the values"
    if propose:
        series.proposed_categories = cat.with_end_column(run_end)
    if cells.has(cat.sheet) and cat.single_row and not is_year_row(
            cells, cat.sheet, cat.row1, range(cat.col1, cat.col2 + 1)):
        return f"category row {cat.row1} does not hold years"
    return ""


def audit(workbook_path: Path, last_year: int) -> list[Series]:
    series_list = read_chart_series(workbook_path)
    sheets = set()
    for series in series_list:
        for reference in (series.values, series.categories):
            for area in parse_reference(reference) or []:
                sheets.add(area.sheet)
    cells = Cells(workbook_path, sheets)
    for series in series_list:
        assess(series, cells, last_year)
    series_list.sort(key=lambda s: (s.page, s.chart_order, s.series_no))
    return series_list


def add_reference(series_list: list[Series], reference_list: list[Series]) -> int:
    """Record, for each series, the values range of the series in the
    same position in the reference workbook. Returns the number paired."""
    by_position = {(s.page, s.chart_at, s.series_no): s for s in reference_list}
    paired = 0
    for series in series_list:
        other = by_position.get((series.page, series.chart_at, series.series_no))
        if other is not None:
            series.reference_values = other.values
            paired += 1
    return paired


# ── Reporting ──────────────────────────────────────────────────────────────────

def summary_lines(series_list: list[Series], last_year: int) -> list[str]:
    charts = {(s.page, s.chart_at) for s in series_list}
    pages = sorted({s.page for s in series_list})
    lines = [
        f"{len(charts)} charts, {len(series_list)} series on {len(pages)} pages. "
        f"Cut-off year: {last_year}.",
        "",
    ]
    totals = Counter(s.status for s in series_list)
    lines.append("Status          Series")
    for status in STATUSES:
        lines.append(f"  {status:<14}{totals.get(status, 0):>6}")

    lines += ["", "By data sheet:"]
    lines.append("  " + f"{'':<12}" + "".join(f"{status:>16}" for status in STATUSES))
    for sheet in sorted({s.data_sheet or "(none)" for s in series_list}):
        counts = Counter(s.status for s in series_list if (s.data_sheet or "(none)") == sheet)
        lines.append("  " + f"{sheet:<12}" + "".join(f"{counts.get(status, 0):>16}" for status in STATUSES))

    lines += ["", "By page:"]
    lines.append("  " + f"{'':<14}" + "".join(f"{status:>16}" for status in STATUSES))
    for page in pages:
        counts = Counter(s.status for s in series_list if s.page == page)
        lines.append("  " + f"{page[:13]:<14}" + "".join(f"{counts.get(status, 0):>16}" for status in STATUSES))

    extend = [s for s in series_list if s.status == "EXTEND"]
    if extend:
        moves = Counter((s.data_sheet, s.range_to_year, s.data_to_year) for s in extend)
        lines += ["", "Extensions proposed (data sheet: range ends -> data ends):"]
        for (sheet, to_year, data_year), count in sorted(moves.items(), key=lambda kv: (kv[0][0], -kv[1])):
            lines.append(f"  {sheet:<12}{to_year} -> {data_year}   {count} series")
    return lines


def write_csv(series_list: list[Series], path: Path) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for series in series_list:
            writer.writerow(series.as_row())


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="List every chart series and whether its range needs extending.")
    parser.add_argument("workbook", type=Path, help="the indicators workbook (.xlsx)")
    parser.add_argument("--out", type=Path, default=Path("chart_audit.csv"),
                        help="CSV file to write (default: chart_audit.csv)")
    parser.add_argument("--last-year", type=int, default=datetime.now().year,
                        help="cut-off year: data after it is treated as projections "
                             "(default: the current year)")
    parser.add_argument("--reference", type=Path, default=None,
                        help="a second workbook whose ranges are shown alongside")
    args = parser.parse_args(argv)

    if not args.workbook.exists():
        print(f"Workbook not found: {args.workbook}")
        return 2

    series_list = audit(args.workbook, args.last_year)
    print(f"Workbook: {args.workbook}")
    for line in summary_lines(series_list, args.last_year):
        print(line)

    if args.reference is not None:
        if not args.reference.exists():
            print(f"Reference workbook not found: {args.reference}")
            return 2
        reference_list = audit(args.reference, args.last_year)
        paired = add_reference(series_list, reference_list)
        same = sum(1 for s in series_list if s.reference_values and s.reference_values == s.values)
        proposed = sum(1 for s in series_list if s.proposed_values and s.proposed_values == s.reference_values)
        print("")
        print(f"Reference: {args.reference}")
        print(f"  {paired} of {len(series_list)} series paired by page, chart position and series order")
        print(f"  {same} have the same range in both workbooks")
        print(f"  {proposed} proposed ranges equal the reference workbook's range")

    write_csv(series_list, args.out)
    print("")
    print(f"Wrote {len(series_list)} lines to {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
