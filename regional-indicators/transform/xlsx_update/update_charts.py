"""
regional-indicators/transform/xlsx_update/update_charts.py

Extends the data ranges of the workbook's chart series to cover the
latest year.

The ranges to change are taken from the chart audit
(transform/chart_audit.py), which is run on the workbook first:

  - EXTEND      the row has data beyond the end of the series' range;
                the range is extended to the last year of the
                continuous run of data, and the category (axis label)
                range with it.
  - MULTI_AREA  the series is made of several separate areas on one
                row; it is replaced by a single range from the start of
                the first area to the end of the data.

Series with any other status are left as they are. Nothing but the
ranges of a series is changed: chart type, formatting, titles, axes and
series names are not touched, and no chart is added or removed.

How a series in Excel is matched to the audit, in order of preference:
by its current values and category references, read from the series
formula (=SERIES(name, categories, values, order)); by its values
reference alone; and by the position of its chart on the page together
with the series name. A match is used only where it identifies exactly
one planned change. Charts inside grouped shapes are included.

The new range is set by rewriting the series formula. If the formula
cannot be read or Excel rejects it, the Values and XValues properties
are assigned directly.

A planned change that matches no series in Excel is reported together
with the series formulas Excel returns for the chart at that position.

After saving, the audit is run again and the number of series still
needing a change is reported; it should be zero for the pages updated.

The workbook is edited through Excel (xlwings); see base.py. This step
uses the Excel object model directly and has been written for Excel on
Windows.

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
from chart_audit import audit, parse_reference                     # noqa: E402

STATUSES_CHANGED = ("EXTEND", "MULTI_AREA")


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
    """One series whose range is to be replaced."""
    page: str
    chart_at: str
    series_no: int
    series_name: str
    status: str
    old_values: str
    old_categories: str
    values: str                     # new values reference
    categories: str | None          # new category reference, or None to keep the current one

    @property
    def label(self) -> str:
        return f"{self.page} {self.chart_at} [{self.series_name or self.series_no}]"


def build_plan(series_list, pages: set[str] | None = None) -> list[Change]:
    """Return a Change for every audited series that has a proposed
    range, optionally limited to the given pages."""
    plan = []
    for series in series_list:
        if series.status not in STATUSES_CHANGED or not series.proposed_values:
            continue
        if pages is not None and series.page not in pages:
            continue
        plan.append(Change(
            page=series.page, chart_at=series.chart_at, series_no=series.series_no,
            series_name=series.series_name or "", status=series.status,
            old_values=series.values, old_categories=series.categories or "",
            values=series.proposed_values, categories=series.proposed_categories or None,
        ))
    return plan


def plan_lines(plan: list[Change]) -> list[str]:
    return [
        f"{change.label} {change.status}: {change.old_values} -> {change.values}"
        + (f"   axis: {change.old_categories} -> {change.categories}" if change.categories else "")
        for change in plan
    ]


class PageIndex:
    """Lookups from what Excel reports about a series to its Change.

    Three keys are tried in turn: the values and category references
    together; the values reference alone; and the chart's position with
    the series name (or, where names repeat within a chart, the series'
    position in the chart). A key is used only if it identifies exactly
    one change.
    """

    def __init__(self, page: str, plan: list[Change]):
        self.changes = [c for c in plan if c.page == page]
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


# ── Excel ──────────────────────────────────────────────────────────────────────

MSO_GROUP = 6          # MsoShapeType.msoGroup


def _items(collection):
    """Yield the members of an Excel collection (1-based Item/Count)."""
    for index in range(1, collection.Count + 1):
        yield collection.Item(index)


def page_charts(sheet_api) -> list:
    """Return every chart-bearing shape on a sheet, including charts
    inside grouped shapes, which the ChartObjects collection does not
    list. Falls back to ChartObjects if the sheet's shapes cannot be
    read."""
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
    it could not be set.

    The series formula is rewritten where it can be read. Otherwise, or
    if Excel rejects the formula, the Values and XValues properties are
    assigned from range objects supplied by `resolve`.
    """
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


def update_page(page: str, chart_objects, plan: list[Change], resolve=None) -> tuple[list[str], Counter]:
    """Apply the plan to the charts of one page.

    `chart_objects` is an iterable of Excel chart shapes (or of objects
    with the same Chart.SeriesCollection()/Formula interface; see
    tests/test_update_charts.py). `resolve` turns a reference into an
    Excel range object and is used only when a series formula cannot be
    rewritten. Returns (result lines, counts).
    """
    results: list[str] = []
    counts: Counter = Counter()
    index = PageIndex(page, plan)
    applied: set[Change] = set()
    seen: dict[str, list[str]] = {}

    for chart_object in chart_objects:
        where = _where(chart_object)
        counts["charts"] += 1
        for position, series in enumerate(_items(chart_object.Chart.SeriesCollection()), start=1):
            counts["series"] += 1
            formula = _read(lambda: series.Formula, None)
            name = str(_read(lambda: series.Name, ""))
            parts = split_series_formula(formula) if isinstance(formula, str) else None
            seen.setdefault(where, []).append(formula if isinstance(formula, str) else f"(formula not readable; name '{name}')")

            change, matched_by = index.find(parts, where, name, position)
            if change is None or change in applied:
                counts["unchanged"] += 1
                continue
            if parts is not None and normalise(parts[2]) == normalise(change.values):
                applied.add(change)                           # already has the new range
                counts["unchanged"] += 1
                continue
            problem = _set_range(series, parts, change, resolve)
            if problem:
                counts["failed"] += 1
                results.append(f"{change.label}: FAILED, {problem}")
                continue
            applied.add(change)
            counts["changed"] += 1
            note = "" if matched_by == "references" else f"   (matched by {matched_by})"
            results.append(f"{change.label}: {change.old_values} -> {change.values}{note}")

    missed = [c for c in index.changes if c not in applied]
    counts["not_found"] = len(missed)
    if missed:
        results.append(
            f"{page}: {len(missed)} planned change(s) matched no series in Excel "
            f"({counts['charts']} charts and {counts['series']} series were read on this page)."
        )
        for chart_at in sorted({c.chart_at for c in missed}):
            for change in (c for c in missed if c.chart_at == chart_at):
                results.append(f"  NOT FOUND {change.label}: expected {change.old_values}")
            if chart_at in seen:
                results.append(f"    series Excel reports for the chart at {chart_at}:")
                results += [f"      {formula}" for formula in seen[chart_at]]
            else:
                results.append(
                    f"    Excel reports no chart at {chart_at}; charts were read at: "
                    f"{', '.join(sorted(seen)) or 'none'}"
                )
    return results, counts


def update_charts(xlsx_path: Path, pages: set[str] | None, last_year: int,
                  dry_run: bool = False, visible: bool = False) -> tuple[list[str], int]:
    """Audit the workbook, apply the proposed ranges and audit again.
    Returns (result lines, number of series that could not be changed)."""
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

    results = [
        f"{len(in_scope)} series on {len({s.page for s in in_scope})} page(s); "
        f"{len(plan)} to change "
        f"({sum(1 for c in plan if c.status == 'EXTEND')} extended, "
        f"{sum(1 for c in plan if c.status == 'MULTI_AREA')} joined into one range)."
    ]
    left_alone = Counter(
        s.status for s in in_scope
        if s.status != "OK" and not (s.status in STATUSES_CHANGED and s.proposed_values)
    )
    if left_alone:
        results.append("Left as they are: " + ", ".join(f"{n} {status}" for status, n in sorted(left_alone.items())))

    if dry_run:
        results += plan_lines(plan)
        results += ["", f"Summary: {len(plan)} would be changed (dry run, nothing written)."]
        return results, 0
    if not plan:
        results += ["", "Summary: 0 changed, 0 flagged for review."]
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

            for page in sorted({change.page for change in plan}):
                lines, counts = update_page(page, page_charts(wb.sheets[page].api), plan, resolve)
                results += lines
                totals += counts
            if totals["changed"]:
                wb.save()
        finally:
            wb.close()
    finally:
        app.quit()

    remaining = 0
    if totals["changed"]:
        after = build_plan(audit(xlsx_path, last_year), pages)
        remaining = len(after)
        if remaining and remaining != totals["failed"] + totals["not_found"]:
            results += [f"STILL OUT OF DATE after saving: {change.label} {change.old_values}" for change in after]
    flagged = max(totals["failed"] + totals["not_found"], remaining)
    results += [
        "",
        f"Summary: {totals['changed']} changed, {flagged} flagged for review "
        f"({totals['failed']} could not be set, {totals['not_found']} not found in Excel, "
        f"{remaining} out of date after saving).",
    ]
    return results, flagged


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Extend chart series ranges to the latest year.")
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
