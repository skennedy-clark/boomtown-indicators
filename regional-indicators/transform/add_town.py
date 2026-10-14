"""
regional-indicators/transform/add_town.py

Adds a town to the indicators workbook by copying an existing town
("the template town"): its rows on the data sheets, its chart page, and
its lines in web_export_map.csv. The new rows are left empty; the
writers fill them from the fetchers' cache on the next run (a row with
no figures yet receives the source's full history).

The town must already be in towns.toml.

How the workbook is read
  Each data sheet (Population, Employment, Business, Income, Housing,
  Crime, Exogenous) is divided into sections ("LGA", "SA2", "UCL",
  "Rainfall", ...) and, within a section, into blocks: a heading row
  naming an area, followed by its indicator rows. Where a section has
  no blocks (Employment), each row is labelled with its area and is
  treated as a block of one row.

  A heading is a row with text in column A, nothing in column B, no
  numbers in the rest of the row, and a label that is never used on a
  row with numbers (which would make it an indicator label such as
  "Historic Average"). Column header rows ("QPS Division", "Chart
  label") have text in column B and are not headings.

  The area a block stands for depends on its section: LGA blocks are
  named after the LGA, SA2 blocks after the SA2 and UCL blocks after
  the town. Sheets without those sections use one level throughout
  (SHEET_LEVELS): Business is by SA2; Income, Crime and Exogenous are
  by town.

What is added, for each block of the template town
  - the new town's block already exists in that section (two towns in
    one SA2 or LGA): it is used as it is;
  - the template's area is also the new town's area (Mitchell and Roma
    are both in Maranoa LGA): the rows are shared;
  - otherwise a copy of the template's block is inserted, in
    alphabetical order among the blocks of the same shape, with the
    area name (and, on the Income sheet, the postcode) changed and the
    figures cleared. Formulas are kept.

  The template's chart page is copied and named after the new town.
  Every chart series that read the template's rows is pointed at the
  new town's rows; series that read other rows (Queensland, Brisbane)
  are left as they are. The town name is replaced in the page's cells,
  chart titles and text boxes, and every replacement is listed.

  web_export_map.csv: the template's lines are copied for the new town.

What has to be checked by hand (listed in the report)
  - text that names the template town in other ways (station names
    such as "Roma Airport", commentary);
  - the CSG notice line on the charts, which is placed by hand;
  - blocks for sources that do not cover the new town stay empty.

Usage:
    python add_town.py <workbook.xlsx> --town Emerald --like Moranbah --dry-run
    python add_town.py <workbook.xlsx> --town Emerald --like Moranbah [--visible]

The workbook is changed in place, through Excel (xlwings); a copy of
the original is saved next to it first, named "<name> (before adding
<town>).xlsx". --dry-run reads the workbook with openpyxl, lists what
would be done and changes nothing (no Excel needed).
"""

from __future__ import annotations

import argparse
import csv
import re
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "xlsx_update"))

DATA_SHEETS = ("Population", "Employment", "Business", "Income", "Housing", "Crime", "Exogenous")

# Rows whose column A holds one of these start a section.
SECTION_HEADINGS = {"LGA", "SA2", "UCL", "State", "Benchmark", "Rainfall", "Education", "Fuel",
                    "Non-primary production", "Primary production"}

# The area level of the blocks in a section, and of sheets not divided into
# area sections. "town" = the town's name, "sa2" = its SA2 name, "lga" = its LGA.
SECTION_LEVELS = {"LGA": "lga", "SA2": "sa2", "UCL": "town"}
SHEET_LEVELS = {"Business": "sa2", "Income": "town", "Crime": "town", "Exogenous": "town"}
NO_LEVEL_SECTIONS = {"State", "Benchmark"}

WEB_MAP = HERE.parent / "web_export_map.csv"


# ── Names ──────────────────────────────────────────────────────────────────────

def norm(text) -> str:
    """Compare labels loosely: case, surrounding space, a trailing "LGA" and a
    state or type suffix such as "(Qld)" or "(R)" are ignored."""
    text = str(text or "").casefold().strip()
    text = re.sub(r"\((qld|nsw|vic|sa|wa|tas|nt|act|[a-z])\)", "", text)
    text = re.sub(r"\blga\b", "", text)
    return re.sub(r"\s+", " ", text).strip()


def lga_label(name: str) -> str:
    """"Central Highlands (Qld)" -> "Central Highlands"."""
    return re.sub(r"\s*\([^)]*\)\s*$", "", name or "").strip()


def area_name(town, level: str) -> str:
    if level == "sa2":
        return town.sa2_name
    if level == "lga":
        return lga_label(town.lga)
    return town.name


def replace_word(text: str, old: str, new: str) -> str:
    """Replace `old` as a whole word (or phrase) in `text`."""
    if not old or old == new:
        return text
    return re.sub(r"(?<![\w])" + re.escape(old) + r"(?![\w])", new, text)


# ── Reading the data sheets ────────────────────────────────────────────────────

def is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


@dataclass
class Unit:
    """A block (heading row and its rows) or a single labelled row."""
    sheet: str
    section: str | None
    section_no: int
    label: str
    first: int
    last: int
    kind: str                   # "block" or "row"
    signature: tuple            # the rows' labels, to recognise blocks of the same shape

    @property
    def size(self) -> int:
        return self.last - self.first + 1


def read_units(sheet_name: str, rows: list[tuple]) -> list[Unit]:
    """Divide a sheet into units. `rows` are the cell values (formulas as
    text) of rows 1..n, as returned by openpyxl's iter_rows(values_only=True)."""
    def a(r):
        value = rows[r - 1][0] if rows[r - 1] else None
        return value.strip() if isinstance(value, str) else value

    def numbers(r):
        return any(is_number(v) for v in rows[r - 1][1:])

    def blank(r):
        return all(v is None or (isinstance(v, str) and not v.strip()) for v in rows[r - 1])

    n = len(rows)
    indicator_labels = {norm(a(r)) for r in range(1, n + 1)
                        if isinstance(a(r), str) and a(r) and numbers(r) and a(r) not in SECTION_HEADINGS}

    def b(r):
        value = rows[r - 1][1] if len(rows[r - 1]) > 1 else None
        return value.strip() if isinstance(value, str) else value

    def heading(r):
        label = a(r)
        if not isinstance(label, str) or not label:
            return False
        if label in SECTION_HEADINGS:
            return True
        label_b = b(r)
        b_is_text = isinstance(label_b, str) and label_b != "" and not label_b.startswith("=")
        return not b_is_text and not numbers(r) and norm(label) not in indicator_labels

    # Sections, and the block headings within them.
    sections: list[tuple[str | None, int, int]] = []   # (name, first row after heading, last row)
    current, start = None, 1
    for r in range(1, n + 1):
        if isinstance(a(r), str) and a(r) in SECTION_HEADINGS:
            sections.append((current, start, r - 1))
            current, start = a(r), r + 1
    sections.append((current, start, n))

    units: list[Unit] = []
    for section_no, (section, s_first, s_last) in enumerate(sections):
        if s_first > s_last:
            continue
        heads = [r for r in range(s_first, s_last + 1) if heading(r)]
        # A heading straight after another heading is a sub-heading of the same block.
        starts = [r for r in heads if (r - 1) not in heads or r == s_first]
        if starts:
            for i, first in enumerate(starts):
                end = (starts[i + 1] - 1) if i + 1 < len(starts) else s_last
                while end > first and blank(end):
                    end -= 1
                signature = tuple(norm(a(r)) for r in range(first + 1, end + 1)
                                  if isinstance(a(r), str) and a(r) and not heading(r))
                units.append(Unit(sheet_name, section, section_no, a(first), first, end, "block", signature))
        else:
            for r in range(s_first, s_last + 1):
                if isinstance(a(r), str) and a(r) and numbers(r):
                    units.append(Unit(sheet_name, section, section_no, a(r), r, r, "row", (norm(b(r)),)))
    return units


def unit_level(unit: Unit) -> str | None:
    if unit.section in NO_LEVEL_SECTIONS:
        return None
    return SECTION_LEVELS.get(unit.section or "") or SHEET_LEVELS.get(unit.sheet)


# ── Planning ───────────────────────────────────────────────────────────────────

@dataclass
class Insert:
    sheet: str
    at: int                     # row (original numbering) before which the rows go in
    rows: int                   # number of rows inserted
    offset: int                 # first row of the copied block within the inserted rows
    template: Unit
    label: str                  # the new block's area name


@dataclass
class Plan:
    town: object
    template: object
    inserts: list[Insert] = field(default_factory=list)
    reused: list[tuple[Unit, Unit]] = field(default_factory=list)     # (template unit, existing unit)
    shared: list[Unit] = field(default_factory=list)
    row_map: dict[tuple[str, int], int] = field(default_factory=dict)  # final rows: template -> new
    notes: list[str] = field(default_factory=list)

    def shift(self, sheet: str, row: int) -> int:
        """Where an original row ends up once every insert is made."""
        return row + sum(i.rows for i in self.inserts if i.sheet == sheet and i.at <= row)

    def new_first(self, insert: Insert) -> int:
        """Final row of the first row of an inserted block."""
        before = sum(i.rows for i in self.inserts if i.sheet == insert.sheet and i.at < insert.at)
        return insert.at + before + insert.offset


def _text_label(value) -> str:
    if not isinstance(value, str) or value.startswith("=") or _numeric_text(value):
        return ""
    return norm(value)


def _match_rows(rows, template_unit: Unit, existing: Unit) -> dict[int, int | None]:
    """Template row -> the existing block's row with the same labels: column
    A and B, or column A alone where that is unique in the existing block.
    Column B labels often name the area, so they are compared with the
    area names removed."""
    def labels(r, area):
        row = rows[r - 1]
        a = _text_label(row[0]) if row else ""
        b = _text_label(row[1]) if len(row) > 1 else ""
        return a, b.replace(norm(area), "").strip()

    by_both, by_a = {}, {}
    for r in range(existing.first, existing.last + 1):
        a, b = labels(r, existing.label)
        by_both.setdefault((a, b), []).append(r)
        by_a.setdefault(a, []).append(r)
    found = {}
    for r in range(template_unit.first, template_unit.last + 1):
        a, b = labels(r, template_unit.label)
        if r == template_unit.first:
            found[r] = existing.first
        elif len(by_both.get((a, b), [])) == 1:
            found[r] = by_both[(a, b)][0]
        elif a and len(by_a.get(a, [])) == 1:
            found[r] = by_a[a][0]
        else:
            found[r] = None
    return found


def make_plan(sheet_rows: dict[str, list[tuple]], town, template) -> Plan:
    plan = Plan(town, template)
    units_by_sheet = {name: read_units(name, rows) for name, rows in sheet_rows.items()}

    for sheet, units in units_by_sheet.items():
        for unit in units:
            level = unit_level(unit)
            if level is None or norm(unit.label) != norm(area_name(template, level)):
                continue
            new_label = area_name(town, level)
            if not new_label:
                plan.notes.append(f"{sheet} {unit.section or ''} {unit.label}: the new town has no {level} name; skipped")
                continue
            if norm(new_label) == norm(unit.label):
                plan.shared.append(unit)
                continue
            same_section = [u for u in units if u.section_no == unit.section_no]
            existing = [u for u in same_section if norm(u.label) == norm(new_label) and u.kind == unit.kind]
            if existing:
                plan.reused.append((unit, existing[0]))
                continue
            plan.inserts.append(_place(unit, new_label, same_section))
            rows = sheet_rows[sheet]
            plan.notes += outside_references(rows, unit)
            for r in range(unit.first + 1, unit.last + 1):
                label = rows[r - 1][0]
                if isinstance(label, str) and norm(template.name) in norm(label):
                    plan.notes.append(f"{sheet} {unit.section or ''}: row {label!r} will read "
                                      f"{replace_word(label.strip(), template.name, town.name)!r}; check it "
                                      f"(for rainfall, the station name)")

    # Rows that the template's charts read, mapped to the new town's rows.
    for insert in plan.inserts:
        first = plan.new_first(insert)
        for r in range(insert.template.first, insert.template.last + 1):
            plan.row_map[(insert.sheet, plan.shift(insert.sheet, r))] = first + (r - insert.template.first)
    for template_unit, existing in plan.reused:
        rows = sheet_rows[template_unit.sheet]
        for r, target in _match_rows(rows, template_unit, existing).items():
            if target is None:
                if _text_label(rows[r - 1][0]):
                    plan.notes.append(f"{template_unit.sheet} row {r} ({rows[r - 1][0]}): no matching row "
                                      f"under {existing.label}; charts reading it are left unchanged")
                continue
            plan.row_map[(template_unit.sheet, plan.shift(template_unit.sheet, r))] = plan.shift(template_unit.sheet, target)
    return plan


CELL_REFERENCE = re.compile(r"(?<![A-Za-z0-9_!$'.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])")


def outside_references(rows, unit: Unit) -> list[str]:
    """Notes for formulas in a block that read other sheets, or rows outside
    the block by relative reference: copied to the new block, those would
    read rows that do not belong to the new town."""
    notes = []
    for r in range(unit.first, unit.last + 1):
        for value in rows[r - 1]:
            if not (isinstance(value, str) and value.startswith("=")):
                continue
            if "!" in value:
                notes.append(f"{unit.sheet} row {r}: formula {value} reads another sheet; check it in the copy")
                break
            # Absolute rows ($41) keep pointing at the same cell, which is
            # what they are for (a section heading); relative rows move.
            outside = [int(m.group(4)) for m in CELL_REFERENCE.finditer(value)
                       if not m.group(3) and not unit.first <= int(m.group(4)) <= unit.last]
            if outside:
                notes.append(f"{unit.sheet} row {r}: formula {value} reads rows outside the {unit.label} block; "
                             f"check it in the copy")
                break
    return notes


def _place(unit: Unit, new_label: str, same_section: list[Unit]) -> Insert:
    """Where a copy of `unit` goes: in alphabetical order among the blocks
    of its section, with the same spacing as the template block."""
    siblings = sorted((u for u in same_section if u.kind == unit.kind), key=lambda u: u.first)
    index = siblings.index(unit)
    gap = 0
    if unit.kind == "block":
        if index + 1 < len(siblings):
            gap = siblings[index + 1].first - unit.last - 1
        elif index > 0:
            gap = unit.first - siblings[index - 1].last - 1
    after = [u for u in siblings if u.label.casefold() > new_label.casefold()]
    if after:
        return Insert(unit.sheet, after[0].first, unit.size + gap, 0, unit, new_label)
    last = siblings[-1]
    return Insert(unit.sheet, last.last + 1, unit.size + gap, gap, unit, new_label)


# ── Chart series ───────────────────────────────────────────────────────────────

REFERENCE = re.compile(
    r"(?P<sheet>'(?:[^']|'')+'|[A-Za-z0-9_.]+)!"
    r"(?P<c1>\$?[A-Z]{1,3})(?P<d1>\$?)(?P<r1>\d+)"
    r"(?::(?P<c2>\$?[A-Z]{1,3})(?P<d2>\$?)(?P<r2>\d+))?"
)


def remap_formula(formula: str, row_map: dict[tuple[str, int], int]) -> str:
    """Point every reference in a SERIES formula at the mapped rows. A
    reference is changed only if each of its rows is mapped."""
    def repl(match):
        sheet = match.group("sheet").strip("'").replace("''", "'")
        r1 = int(match.group("r1"))
        r2 = int(match.group("r2")) if match.group("r2") else None
        n1 = row_map.get((sheet, r1))
        if n1 is None:
            return match.group(0)
        text = f"{match.group('sheet')}!{match.group('c1')}{match.group('d1')}{n1}"
        if r2 is not None:
            n2 = row_map.get((sheet, r2))
            if n2 is None or n2 - n1 != r2 - r1:
                return match.group(0)
            text += f":{match.group('c2')}{match.group('d2')}{n2}"
        return text
    return REFERENCE.sub(repl, formula or "")


# ── web_export_map.csv ─────────────────────────────────────────────────────────

def web_map_lines(path: Path, town, template) -> tuple[list[dict], list[str]]:
    """The map lines to add for `town`, copied from the template's lines.
    Returns (lines, header)."""
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        header, lines = reader.fieldnames, list(reader)
    if any(norm(line["folder"]) == norm(town.name) for line in lines):
        return [], header
    added = []
    for line in lines:
        if norm(line["folder"]) != norm(template.name):
            continue
        new = dict(line, folder=town.name)
        section = line.get("section") or None
        level = None if section in NO_LEVEL_SECTIONS else (SECTION_LEVELS.get(section or "") or SHEET_LEVELS.get(line["sheet"]))
        if level:
            old_area, new_area = area_name(template, level), area_name(town, level)
            if line["block"] and norm(line["block"]) == norm(old_area):
                new["block"] = new_area
            elif not line["block"] and norm(line["row_label"]) == norm(old_area):
                new["row_label"] = new_area
        added.append(new)
    return added, header


def append_web_map(path: Path, lines: list[dict], header: list[str]) -> None:
    """Append lines, keeping the file's own line endings."""
    raw = path.read_bytes()
    ending = "\r\n" if b"\r\n" in raw else "\n"
    with open(path, "a", encoding="utf-8", newline="") as f:
        if raw and not raw.endswith(b"\n"):
            f.write(ending)
        writer = csv.DictWriter(f, fieldnames=header, lineterminator=ending)
        writer.writerows(lines)


# ── Report ─────────────────────────────────────────────────────────────────────

def plan_lines(plan: Plan, template_series: list | None = None) -> list[str]:
    town, template = plan.town, plan.template
    lines = [f"Adding {town.name}, copied from {template.name}.", ""]
    lines.append("New blocks:")
    for i in sorted(plan.inserts, key=lambda i: (DATA_SHEETS.index(i.sheet), i.at)):
        where = f"{i.sheet} {i.template.section or ''}".strip()
        lines.append(f"  {where}: {i.label} ({i.template.size} rows, copied from {i.template.label} "
                     f"rows {i.template.first}-{i.template.last}), inserted at row {i.at}")
    if not plan.inserts:
        lines.append("  none")
    if plan.reused:
        lines.append("Existing blocks used:")
        for template_unit, existing in plan.reused:
            lines.append(f"  {template_unit.sheet} {template_unit.section or ''}: {existing.label} (rows {existing.first}-{existing.last})")
    if plan.shared:
        lines.append("Shared with the template town (same area):")
        for unit in plan.shared:
            lines.append(f"  {unit.sheet} {unit.section or ''}: {unit.label}")
    if template_series is not None:
        mapped = left = 0
        for series in template_series:
            text = f"{series.values},{series.categories}"
            if remap_formula(text, _original_map(plan)) != text:
                mapped += 1
            else:
                left += 1
        lines.append(f"Chart page: {len(template_series)} series on {template.name}; "
                     f"{mapped} will read the new town's rows, {left} stay as they are (benchmarks, shared areas).")
    if plan.notes:
        lines += ["", "Notes:"] + [f"  {n}" for n in plan.notes]
    return lines


def _original_map(plan: Plan) -> dict:
    """The row map keyed by the template's original rows (for the dry run,
    where the charts still read the original rows)."""
    reverse = {}
    for (sheet, final_row), new_row in plan.row_map.items():
        for original in range(max(1, final_row - 400), final_row + 1):
            if plan.shift(sheet, original) == final_row:
                reverse[(sheet, original)] = new_row
                break
    return reverse


# ── Excel ──────────────────────────────────────────────────────────────────────

def apply_plan(workbook: Path, plan: Plan, visible: bool = False) -> list[str]:
    import xlwings as xw
    from update_charts import page_charts

    town, template = plan.town, plan.template
    report: list[str] = []
    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(str(workbook))
        try:
            names = [s.name for s in wb.sheets]
            if town.name in names:
                raise ValueError(f"The workbook already has a sheet named {town.name}.")
            if template.name not in names:
                raise ValueError(f"The workbook has no page named {template.name} to copy.")

            # 1. Copy the template's page first, so that the inserts below
            #    move its chart references along with everything else.
            wb.sheets[template.name].api.Copy(After=wb.sheets[len(names) - 1].api)
            copies = [s for s in wb.sheets if s.name not in names]
            if len(copies) != 1:
                raise RuntimeError(f"Copying the {template.name} page did not give one new sheet.")
            page = copies[0]
            page.name = town.name
            report.append(f"Page {town.name} copied from {template.name}.")

            # 2. Insert the new blocks, bottom-up on each sheet.
            for sheet_name in DATA_SHEETS:
                inserts = sorted((i for i in plan.inserts if i.sheet == sheet_name), key=lambda i: -i.at)
                if not inserts:
                    continue
                sheet = wb.sheets[sheet_name]
                done: list[Insert] = []
                for insert in inserts:
                    sheet.api.Rows(f"{insert.at}:{insert.at + insert.rows - 1}").Insert()
                    done.append(insert)
                    t = insert.template
                    current = t.first + sum(d.rows for d in done if d.at <= t.first)
                    dest = insert.at + insert.offset
                    sheet.api.Rows(f"{current}:{current + t.size - 1}").Copy(
                        sheet.api.Rows(f"{dest}:{dest + t.size - 1}"))
                    _clear_and_relabel(sheet, dest, t.size, template, town, insert.label, t.label)
                    report.append(f"{sheet_name}: {insert.label} added at row {dest} ({t.size} rows).")

            # 3. Point the copied page's charts at the new rows.
            changed = unchanged = failed = 0
            for chart_object in page_charts(page.api):
                chart = chart_object.Chart
                for index in range(1, chart.SeriesCollection().Count + 1):
                    series = chart.SeriesCollection(index)
                    try:
                        old = series.Formula
                    except Exception:
                        failed += 1
                        continue
                    new = remap_formula(old, plan.row_map)
                    if new == old:
                        unchanged += 1
                        continue
                    try:
                        series.Formula = new
                        changed += 1
                    except Exception as exc:
                        failed += 1
                        report.append(f"  chart at {chart_object.TopLeftCell.Address}: could not set {new} ({exc})")
                try:
                    if chart.HasTitle:
                        title = chart.ChartTitle.Text
                        new_title = replace_word(title, template.name, town.name)
                        if new_title != title:
                            chart.ChartTitle.Text = new_title
                            report.append(f"  chart title: {title!r} -> {new_title!r}")
                except Exception:
                    pass
            report.append(f"Charts: {changed} series now read {town.name}'s rows, {unchanged} unchanged, {failed} failed.")

            # 4. The town's name in the page's cells and text boxes.
            report += _rename_on_page(page, template.name, town.name, plan.row_map)

            wb.save()
        finally:
            wb.close()
    finally:
        app.quit()
    return report


def _clear_and_relabel(sheet, first: int, size: int, template, town, new_label: str, old_label: str) -> None:
    """Clear the figures in a copied block and give it the new town's names.
    Formulas are kept; column B text is kept (with names replaced)."""
    last_col = sheet.used_range.last_cell.column
    rng = sheet.range((first, 1), (first + size - 1, last_col))
    formulas = rng.formula
    if size == 1:
        formulas = [formulas] if not isinstance(formulas[0], (list, tuple)) else formulas
    new_rows = []
    for row in formulas:
        row = list(row)
        for c, cell in enumerate(row):
            if isinstance(cell, str) and cell.startswith("="):
                continue
            if c == 0:
                text = str(cell).strip() if cell is not None else ""
                if text and norm(text) == norm(old_label):
                    row[c] = new_label
                elif text and text.split(".")[0] == str(template.postcode).strip() and town.postcode:
                    row[c] = town.postcode
                elif isinstance(cell, str):
                    row[c] = replace_word(replace_word(cell, old_label, new_label), template.name, town.name)
            elif c == 1 and isinstance(cell, str) and cell and not _numeric_text(cell):
                row[c] = replace_word(replace_word(cell, old_label, new_label), template.name, town.name)
            elif cell not in (None, ""):
                row[c] = ""
        new_rows.append(row)
    rng.formula = new_rows


def _numeric_text(text: str) -> bool:
    try:
        float(text)
        return True
    except ValueError:
        return False


def _rename_on_page(page, old: str, new: str, row_map: dict) -> list[str]:
    """Replace the town's name in the page's cells, and point cell formulas
    that read the template's rows at the new town's rows. Cells are written
    one at a time (merged cells refuse a block write)."""
    report = []
    used = page.used_range
    formulas = used.formula
    if not isinstance(formulas, (list, tuple)):
        formulas = [[formulas]]
    elif formulas and not isinstance(formulas[0], (list, tuple)):
        formulas = [formulas]
    remapped = 0
    for r, row in enumerate(formulas):
        for c, cell in enumerate(row):
            if not isinstance(cell, str) or not cell:
                continue
            updated = remap_formula(cell, row_map) if cell.startswith("=") else replace_word(cell, old, new)
            if updated == cell:
                continue
            try:
                page.cells(used.row + r, used.column + c).formula = updated
            except Exception as exc:
                report.append(f"  page cell {used.row + r}:{used.column + c}: could not change ({exc})")
                continue
            if cell.startswith("="):
                remapped += 1
            else:
                report.append(f"  page cell {page.cells(used.row + r, used.column + c).address}: {cell!r} -> {updated!r}")
    if remapped:
        report.append(f"  {remapped} formulas on the page now read {new}'s rows")
    from update_charts import _items
    for shape in _items(page.api.Shapes):
        try:
            if shape.HasChart:
                continue
            if shape.TextFrame2.HasText:
                text = shape.TextFrame2.TextRange.Text
                if old in text:
                    report.append(f"  text box {shape.Name!r} mentions {old}: check it ({text[:60]!r})")
        except Exception:
            continue
    report.append("  Check by hand: station names, commentary, and the CSG notice line on the charts.")
    return report


# ── Main ───────────────────────────────────────────────────────────────────────

def read_sheet_rows(workbook: Path) -> dict[str, list[tuple]]:
    import openpyxl
    wb = openpyxl.load_workbook(workbook, read_only=True, data_only=False)
    try:
        return {name: [tuple(row) for row in wb[name].iter_rows(values_only=True)]
                for name in DATA_SHEETS if name in wb.sheetnames}
    finally:
        wb.close()


def find_town(config, name: str):
    for town in config.towns:
        if norm(town.name) == norm(name):
            return town
    raise SystemExit(f"{name} is not in towns.toml.")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Add a town to the indicators workbook, copied from another town.")
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--town", required=True, help="the town to add (as named in towns.toml)")
    parser.add_argument("--like", required=True, metavar="TOWN", help="the town to copy (its page and rows)")
    parser.add_argument("--dry-run", action="store_true", help="list what would be done; change nothing")
    parser.add_argument("--visible", action="store_true", help="show Excel while the town is added")
    args = parser.parse_args(argv)

    if not args.workbook.is_file():
        print(f"Workbook not found: {args.workbook}")
        return 2

    from config import get_config
    config = get_config()
    town, template = find_town(config, args.town), find_town(config, args.like)

    plan = make_plan(read_sheet_rows(args.workbook), town, template)
    series = None
    try:
        from chart_audit import read_chart_series
        series = [s for s in read_chart_series(args.workbook) if s.page == template.name]
    except Exception as exc:
        plan.notes.append(f"Chart series could not be read ({exc}).")
    lines = plan_lines(plan, series)

    map_lines, header = web_map_lines(WEB_MAP, town, template)
    lines.append(f"web_export_map.csv: {len(map_lines)} lines to add for {town.name}"
                 + ("" if map_lines else " (none: the town is already in the map, or the template has no lines)"))

    if args.dry_run:
        print("\n".join(lines))
        print("\nDry run: nothing changed.")
        return 0

    backup = args.workbook.with_name(f"{args.workbook.stem} (before adding {town.name}){args.workbook.suffix}")
    shutil.copy2(args.workbook, backup)
    lines.append(f"Original saved as {backup.name}.")
    lines += apply_plan(args.workbook, plan, visible=args.visible)
    if map_lines:
        append_web_map(WEB_MAP, map_lines, header)
    print("\n".join(lines))
    return 0


if __name__ == "__main__":
    sys.exit(main())
