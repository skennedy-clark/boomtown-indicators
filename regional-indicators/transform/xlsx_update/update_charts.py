"""
regional-indicators/transform/xlsx_update/update_charts.py

Brings the workbook's charts up to date so that each shows every year
for which its data rows have figures.

What needs doing is taken from the chart audit (transform/chart_audit.py),
which is run on the workbook first. Most charts use Excel's chart
filters: a series has a full range, and some of its years (categories)
are filtered out. Two kinds of change follow:

  - UNHIDE   the full range already covers the new year, but the year is
             filtered out. The year is shown again by clearing the
             filter on that category (ChartGroup.FullCategoryCollection,
             IsFiltered = False). Category filters apply to the whole
             chart, so a year is shown for every series in the chart.
  - EXTEND   the data runs past the end of the full range, or the
             series' category (axis label) range stops before the data.
             The series formula is rewritten with the longer range(s);
             any of the new years that are filtered out are then shown
             as above. (In a chart with column and line groups, each
             group takes its axis labels from its own series, so a short
             category range keeps the new year off that group's axis.)

Series with any other status are left as they are, and years with no
figures stay hidden. Nothing else is changed: chart type, formatting,
titles, axes and series names are not touched, and no chart is added or
removed.

How a series in Excel is matched to the audit, in order of preference:
by its current values and category references, read from the series
formula (=SERIES(name, categories, values, order)), which reports the
full ranges; by its values reference alone; and by the position of its
chart on the page together with the series name. A match is used only
where it identifies exactly one planned change. Charts are identified by
the cell under their top-left corner; charts inside grouped shapes are
included.

If a series formula cannot be read or Excel rejects the new one, the
Values and XValues properties are assigned directly.

After saving, the audit is run again and any series that still needs a
change is listed; there should be none for the pages updated.

The workbook is edited through Excel (xlwings); see base.py. This step
uses the Excel object model directly (chart filters require Excel 2013
or later) and has been written for Excel on Windows.

Usage:
    python update_charts.py <workbook.xlsx> [--page Chinchilla ...]
                            [--last-year 2025] [--dry-run] [--visible]

--page limits the update to the named pages (sheet names). --dry-run
lists the changes without opening Excel.
"""

from __future__ import annotations

import argparse
import sys
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
from chart_audit import CHANGE_STATUSES, as_year, audit, parse_reference      # noqa: E402


# ── Series formulas ────────────────────────────────────────────────────────────

def normalise(reference: str) -> str:
    """Reduce a reference to a form that compares equal however Excel
    or the workbook file writes it: no leading '=', outer parentheses,
    '$', spaces or quotes around plain sheet names; upper case."""
    text = (reference or "").strip()
    if text.startswith("="):
        text = text[1:]
    if text.startswith("(") and text.endswith(")"):
        text = text[1:-1]
    out, quoted = [], False
    for char in text:
        if char == "'":
            quoted = not quoted
            continue
        if not quoted and char in "$ ":
            continue
        out.append(char)
    return "".join(out).upper()


def split_series_formula(formula: str) -> list[str] | None:
    """Split =SERIES(a, b, c, d) into its arguments, or return None if
    the text is not a SERIES formula. Commas inside parentheses, quoted
    sheet names and string literals do not split."""
    text = (formula or "").strip()
    if not text.upper().startswith("=SERIES(") or not text.endswith(")"):
        return None
    body = text[len("=SERIES("):-1]
    parts, current, depth, quote = [], [], 0, ""
    for char in body:
        if quote:
            current.append(char)
            if char == quote:
                quote = ""
            continue
        if char in "'\"":
            quote = char
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        elif char == "," and depth == 0:
            parts.append("".join(current))
            current = []
            continue
        current.append(char)
    parts.append("".join(current))
    return parts if len(parts) >= 4 else None


def rebuild_series_formula(parts: list[str], values: str, categories: str | None) -> str:
    """Return the SERIES formula with new values and, if given, new
    categories; the name, order and any further arguments are kept."""
    new = list(parts)
    new[2] = values
    if categories:
        new[1] = categories
    return "=SERIES(" + ",".join(new) + ")"


# ── Plan ───────────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Change:
    """One series whose full range is to be extended."""
    page: str
    chart_at: str
    series_no: int
    series_name: str
    old_values: str
    old_categories: str
    values: str                     # new values reference
    categories: str | None          # new category reference, or None to keep the current one

    @property
    def label(self) -> str:
        return f"{self.page} {self.chart_at} [{self.series_name or self.series_no}]"


@dataclass
class Plan:
    changes: list[Change]
    unhide: dict[tuple[str, str], set[int]]          # (page, chart position) -> years to show

    @property
    def pages(self) -> set[str]:
        return {c.page for c in self.changes} | {page for page, _ in self.unhide}

    def __bool__(self) -> bool:
        return bool(self.changes or self.unhide)


def build_plan(series_list, pages: set[str] | None = None) -> Plan:
    """Collect the range extensions and the years to show, optionally
    limited to the given pages."""
    changes: list[Change] = []
    unhide: dict[tuple[str, str], set[int]] = {}
    for series in series_list:
        if series.status not in CHANGE_STATUSES:
            continue
        if pages is not None and series.page not in pages:
            continue
        if series.status == "EXTEND" and series.proposed_values:
            changes.append(Change(
                page=series.page, chart_at=series.chart_at, series_no=series.series_no,
                series_name=series.series_name or "",
                old_values=series.values, old_categories=series.categories or "",
                values=series.proposed_values, categories=series.proposed_categories or None,
            ))
        years = {y for y in (series.unhide_years or []) if y is not None}
        if years:
            unhide.setdefault((series.page, series.chart_at), set()).update(years)
    return Plan(changes, unhide)


def plan_lines(plan: Plan) -> list[str]:
    lines = []
    charts = sorted({(c.page, c.chart_at) for c in plan.changes} | set(plan.unhide))
    for page, chart_at in charts:
        years = sorted(plan.unhide.get((page, chart_at), ()))
        if years:
            lines.append(f"{page} {chart_at}: show {', '.join(str(y) for y in years)}")
        for change in (c for c in plan.changes if (c.page, c.chart_at) == (page, chart_at)):
            lines.append(
                f"{change.label}: extend {change.old_values} -> {change.values}"
                + (f"   axis: {change.old_categories} -> {change.categories}" if change.categories else "")
            )
    return lines


class PageIndex:
    """Lookups from what Excel reports about a series to its Change.

    Three keys are tried in turn: the values and category references
    together; the values reference alone; and the chart's position with
    the series name (or, for an unnamed series, its position in the
    chart). A key is used only if it identifies exactly one change.
    """

    def __init__(self, page: str, changes: list[Change]):
        self.changes = [c for c in changes if c.page == page]
        self._by_refs = self._unique((normalise(c.old_values), normalise(c.old_categories)) for c in self.changes)
        self._by_values = self._unique(normalise(c.old_values) for c in self.changes)
        self._by_name = self._unique((c.chart_at, c.series_name.strip()) for c in self.changes)
        self._by_position = self._unique((c.chart_at, c.series_no) for c in self.changes)

    def _unique(self, keys) -> dict:
        found: dict = {}
        for change, key in zip(self.changes, keys):
            found.setdefault(key, []).append(change)
        return {key: changes[0] for key, changes in found.items() if len(changes) == 1}

    def find(self, parts: list[str] | None, chart_at: str, name: str, position: int) -> tuple[Change | None, str]:
        if parts is not None:
            values, categories = normalise(parts[2]), normalise(parts[1])
            if (values, categories) in self._by_refs:
                return self._by_refs[(values, categories)], "references"
            if values in self._by_values:
                return self._by_values[values], "values reference"
        if name and (chart_at, name.strip()) in self._by_name:
            return self._by_name[(chart_at, name.strip())], "chart position and series name"
        if not name and (chart_at, position) in self._by_position:
            return self._by_position[(chart_at, position)], "chart position and series order"
        return None, ""


def category_year(name) -> int | None:
    """The year a category label stands for: 2025, "2025" or "2024/25"."""
    year = as_year(name)
    if year is None and isinstance(name, str):
        try:
            year = as_year(float(name.strip()))
        except ValueError:
            year = None
    return year


# ── Excel ──────────────────────────────────────────────────────────────────────

MSO_GROUP = 6          # MsoShapeType.msoGroup


def _items(collection):
    """Yield the members of an Excel collection (1-based Item/Count)."""
    for index in range(1, collection.Count + 1):
        yield collection.Item(index)


def page_charts(sheet_api) -> list:
    """Return every chart-bearing shape on a sheet, including charts
    inside grouped shapes. Falls back to the ChartObjects collection if
    the sheet's shapes cannot be read."""
    found = []

    def visit(shape):
        try:
            if shape.Type == MSO_GROUP:
                for member in _items(shape.GroupItems):
                    visit(member)
                return
            if shape.HasChart:
                found.append(shape)
        except Exception:
            pass

    try:
        for shape in _items(sheet_api.Shapes):
            visit(shape)
    except Exception:
        found = []
    if not found:
        found = list(_items(sheet_api.ChartObjects()))
    return found


def _read(getter, default=""):
    try:
        value = getter()
        return default if value is None else value
    except Exception:
        return default


def _where(chart_object) -> str:
    return str(_read(lambda: chart_object.TopLeftCell.Address, "chart")).replace("$", "")


def _set_range(series, parts: list[str] | None, change: Change, resolve) -> str:
    """Give a series its new range. Returns "" on success, or the reason
    it could not be set."""
    reason = "the series formula could not be read"
    if parts is not None:
        try:
            series.Formula = rebuild_series_formula(parts, change.values, change.categories)
            return ""
        except Exception as exc:
            reason = f"Excel rejected the formula ({exc})"
    if resolve is None:
        return reason
    try:
        series.Values = resolve(change.values)
        if change.categories:
            series.XValues = resolve(change.categories)
        return ""
    except Exception as exc:
        return f"{reason}; assigning the ranges directly also failed ({exc})"


def show_years(chart, years: set[int]) -> tuple[set[int], str]:
    """Clear the chart filter on the categories for `years`. Returns
    (years found among the chart's categories, problem or "")."""
    found: set[int] = set()
    try:
        for group in _items(chart.ChartGroups()):
            for category in _items(group.FullCategoryCollection()):
                year = category_year(_read(lambda: category.Name, ""))
                if year in years:
                    found.add(year)
                    if category.IsFiltered:
                        category.IsFiltered = False
    except Exception as exc:
        return found, f"the chart's category filter could not be changed ({exc})"
    return found, ""


def update_page(page: str, chart_objects, plan: Plan, resolve=None) -> tuple[list[str], Counter]:
    """Apply the plan to the charts of one page.

    `chart_objects` is an iterable of Excel chart shapes (or objects with
    the same interface; see tests/test_update_charts.py). `resolve` turns
    a reference into an Excel range object and is used only when a
    series formula cannot be rewritten. Returns (result lines, counts).
    """
    results: list[str] = []
    counts: Counter = Counter()
    index = PageIndex(page, plan.changes)
    applied: set[Change] = set()
    seen: dict[str, list[str]] = {}

    for chart_object in chart_objects:
        where = _where(chart_object)
        counts["charts"] += 1
        chart = chart_object.Chart
        for position, series in enumerate(_items(chart.SeriesCollection()), start=1):
            counts["series"] += 1
            formula = _read(lambda: series.Formula, None)
            name = str(_read(lambda: series.Name, ""))
            parts = split_series_formula(formula) if isinstance(formula, str) else None
            seen.setdefault(where, []).append(formula if isinstance(formula, str) else f"(formula not readable; name '{name}')")

            change, matched_by = index.find(parts, where, name, position)
            if change is None or change in applied:
                continue
            if parts is not None and normalise(parts[2]) == normalise(change.values) and (
                    not change.categories or normalise(parts[1]) == normalise(change.categories)):
                applied.add(change)                           # already has the new ranges
                continue
            problem = _set_range(series, parts, change, resolve)
            if problem:
                counts["failed"] += 1
                results.append(f"{change.label}: FAILED, {problem}")
                continue
            applied.add(change)
            counts["extended"] += 1
            note = "" if matched_by == "references" else f"   (matched by {matched_by})"
            if normalise(change.old_values) == normalise(change.values):
                results.append(f"{change.label}: axis extended {change.old_categories} -> {change.categories}{note}")
            else:
                results.append(f"{change.label}: extended {change.old_values} -> {change.values}{note}")

        years = plan.unhide.get((page, where))
        if years:
            found, problem = show_years(chart, years)
            if problem:
                counts["failed"] += 1
                results.append(f"{page} {where}: FAILED, {problem}")
            elif found:
                counts["charts_shown"] += 1
                results.append(f"{page} {where}: showing {', '.join(str(y) for y in sorted(found))}")
            missing_years = sorted(years - found)
            if missing_years and not problem:
                counts["years_not_on_axis"] += 1
                results.append(
                    f"{page} {where}: no category for {', '.join(str(y) for y in missing_years)} "
                    f"on the chart's axis"
                )

    missed = [c for c in index.changes if c not in applied]
    counts["not_found"] = len(missed)
    unseen_charts = sorted(chart_at for (p, chart_at) in plan.unhide if p == page and chart_at not in seen)
    counts["not_found"] += len(unseen_charts)
    if missed or unseen_charts:
        results.append(
            f"{page}: planned changes matched nothing in Excel "
            f"({counts['charts']} charts and {counts['series']} series were read on this page)."
        )
        for chart_at in unseen_charts:
            results.append(f"  NOT FOUND {page} {chart_at}: no chart at this position "
                           f"(charts were read at: {', '.join(sorted(seen)) or 'none'})")
        for change in missed:
            results.append(f"  NOT FOUND {change.label}: expected {change.old_values}")
            for formula in seen.get(change.chart_at, []):
                results.append(f"      Excel has: {formula}")
    return results, counts


def update_charts(xlsx_path: Path, pages: set[str] | None, last_year: int,
                  dry_run: bool = False, visible: bool = False) -> tuple[list[str], int]:
    """Audit the workbook, apply the changes and audit again.
    Returns (result lines, number of problems)."""
    series_list = audit(xlsx_path, last_year)
    known_pages = {s.page for s in series_list}
    if pages is not None:
        unknown = sorted(pages - known_pages)
        if unknown:
            raise ValueError(
                f"No charts on page(s) {', '.join(unknown)}. "
                f"Pages with charts: {', '.join(sorted(known_pages))}"
            )
    plan = build_plan(series_list, pages)
    in_scope = [s for s in series_list if pages is None or s.page in pages]
    statuses = Counter(s.status for s in in_scope)
    results = [
        f"{len(in_scope)} series in {len({(s.page, s.chart_at) for s in in_scope})} charts on "
        f"{len({s.page for s in in_scope})} page(s): "
        f"{statuses.get('UNHIDE', 0)} need filtered years shown, {statuses.get('EXTEND', 0)} need a longer range; "
        f"{len(plan.unhide)} charts to change."
    ]
    left_alone = {k: v for k, v in statuses.items() if k not in CHANGE_STATUSES and k != "OK"}
    if left_alone:
        results.append("Left as they are: " + ", ".join(f"{n} {status}" for status, n in sorted(left_alone.items())))

    if dry_run:
        results += plan_lines(plan)
        results += ["", f"Summary: {len(plan.unhide)} charts would be changed (dry run, nothing written)."]
        return results, 0
    if not plan:
        results += ["", "Summary: 0 charts changed, 0 flagged for review."]
        return results, 0

    import xlwings as xw

    totals: Counter = Counter()
    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    try:
        wb = app.books.open(str(xlsx_path))
        try:
            def resolve(reference: str):
                areas = parse_reference(reference)
                if not areas or len(areas) != 1:
                    raise ValueError(f"not a single range: {reference}")
                area = areas[0]
                return wb.sheets[area.sheet].range((area.row1, area.col1), (area.row2, area.col2)).api

            for page in sorted(plan.pages):
                lines, counts = update_page(page, page_charts(wb.sheets[page].api), plan, resolve)
                results += lines
                totals += counts
            if totals["extended"] or totals["charts_shown"]:
                wb.save()
        finally:
            wb.close()
    finally:
        app.quit()

    remaining = []
    if totals["extended"] or totals["charts_shown"]:
        remaining = [
            s for s in audit(xlsx_path, last_year)
            if (pages is None or s.page in pages) and s.status in CHANGE_STATUSES
        ]
        for s in remaining:
            results.append(
                f"STILL OUT OF DATE after saving: {s.page} {s.chart_at} [{s.series_name or s.series_no}] "
                f"{s.status}: shows {s.shown_values}, data to {s.data_to_year}"
            )
    flagged = totals["failed"] + totals["not_found"] + len(remaining)
    results += [
        "",
        f"Summary: {totals['charts_shown']} charts now show the new year(s), "
        f"{totals['extended']} series ranges extended; {flagged} flagged for review "
        f"({totals['failed']} could not be changed, {totals['not_found']} not found in Excel, "
        f"{len(remaining)} series still out of date after saving).",
    ]
    return results, flagged


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Bring chart series up to date: show filtered years and extend ranges.")
    parser.add_argument("workbook", type=Path, help="the indicators workbook (.xlsx)")
    parser.add_argument("--page", nargs="+", default=None, metavar="PAGE",
                        help="only these pages (sheet names); default: every page with charts")
    parser.add_argument("--last-year", type=int, default=datetime.now().year,
                        help="cut-off year: data after it is treated as projections "
                             "(default: the current year)")
    parser.add_argument("--dry-run", action="store_true", help="list the changes; do not open Excel")
    parser.add_argument("--visible", action="store_true", help="show Excel while the charts are updated")
    args = parser.parse_args(argv)

    if not args.workbook.exists():
        print(f"Workbook not found: {args.workbook}")
        return 2
    try:
        results, _ = update_charts(
            args.workbook, set(args.page) if args.page else None, args.last_year,
            dry_run=args.dry_run, visible=args.visible,
        )
    except ValueError as exc:
        print(exc)
        return 2
    for line in results:
        print(line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
