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

How a series in Excel is matched to the audit: by the page it is on and
by its current values and category references, read from the series
formula (=SERIES(name, categories, values, order)). The position of the
chart on the page and the order of its series are not used, so charts
may be moved or reordered without affecting the match. A series whose
references are not in the audit's list of changes is skipped.
If the category reference does not match, the values reference alone is
used, provided it identifies a single change on the page.

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
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
from chart_audit import audit                                      # noqa: E402

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

def build_plan(series_list, pages: set[str] | None = None) -> dict:
    """Return {(page, values, categories): (new values, new categories)}
    for every audited series that has a proposed range. The key holds
    normalised references; the new references are as proposed."""
    plan = {}
    for series in series_list:
        if series.status not in STATUSES_CHANGED or not series.proposed_values:
            continue
        if pages is not None and series.page not in pages:
            continue
        key = (series.page, normalise(series.values), normalise(series.categories))
        plan[key] = (series.proposed_values, series.proposed_categories or None)
    return plan


def plan_lines(series_list, plan: dict) -> list[str]:
    lines = []
    for series in series_list:
        key = (series.page, normalise(series.values), normalise(series.categories))
        if key not in plan:
            continue
        values, categories = plan[key]
        lines.append(
            f"{series.page} {series.chart_at} [{series.series_name or series.series_no}] "
            f"{series.status}: {series.values} -> {values}"
            + (f"   axis: {series.categories} -> {categories}" if categories else "")
        )
    return lines


# ── Excel ──────────────────────────────────────────────────────────────────────

def _items(collection):
    """Yield the members of an Excel collection (1-based Item/Count)."""
    for index in range(1, collection.Count + 1):
        yield collection.Item(index)


def update_page(page: str, chart_objects, plan: dict) -> tuple[list[str], Counter]:
    """Apply the plan to the charts of one page.

    `chart_objects` is an iterable of Excel ChartObject objects (or of
    objects with the same Chart.SeriesCollection()/Formula interface;
    see tests/test_update_charts.py). Returns (result lines, counts).
    """
    results: list[str] = []
    counts: Counter = Counter()
    candidates: dict[str, set] = {}
    for (plan_page, values, _), change in plan.items():
        if plan_page == page:
            candidates.setdefault(values, set()).add(change)
    by_values = {values: next(iter(changes)) for values, changes in candidates.items() if len(changes) == 1}
    for chart_object in chart_objects:
        where = _where(chart_object)
        for series in _items(chart_object.Chart.SeriesCollection()):
            try:
                formula = series.Formula
            except Exception as exc:                       # Excel refuses for some empty series
                counts["unreadable"] += 1
                results.append(f"{page} {where}: series formula could not be read ({exc})")
                continue
            parts = split_series_formula(formula)
            if parts is None:
                counts["unreadable"] += 1
                results.append(f"{page} {where}: not a SERIES formula: {formula}")
                continue
            change = plan.get((page, normalise(parts[2]), normalise(parts[1])))
            if change is None:
                # Excel can report the category reference differently
                # from the file; fall back to the values reference when
                # it identifies one change on the page.
                change = by_values.get(normalise(parts[2]))
            if change is None:
                counts["unchanged"] += 1
                continue
            values, categories = change
            new_formula = rebuild_series_formula(parts, values, categories)
            try:
                series.Formula = new_formula
            except Exception as exc:
                counts["failed"] += 1
                results.append(f"{page} {where}: FAILED to set {new_formula} ({exc})")
                continue
            counts["changed"] += 1
            results.append(f"{page} {where}: {parts[2]} -> {values}")
    return results, counts


def _where(chart_object) -> str:
    try:
        return chart_object.TopLeftCell.Address.replace("$", "")
    except Exception:
        return "chart"


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
    to_change = [s for s in in_scope if s.status in STATUSES_CHANGED and s.proposed_values]

    results = [
        f"{len(in_scope)} series on {len({s.page for s in in_scope})} page(s); "
        f"{len(to_change)} to change "
        f"({sum(1 for s in to_change if s.status == 'EXTEND')} extended, "
        f"{sum(1 for s in to_change if s.status == 'MULTI_AREA')} joined into one range)."
    ]
    left_alone = Counter(s.status for s in in_scope if s not in to_change and s.status != "OK")
    if left_alone:
        results.append("Left as they are: " + ", ".join(f"{n} {status}" for status, n in sorted(left_alone.items())))

    if dry_run:
        results += plan_lines(series_list, plan)
        results += ["", f"Summary: {len(to_change)} would be changed (dry run, nothing written)."]
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
            for page in sorted({key[0] for key in plan}):
                chart_objects = _items(wb.sheets[page].api.ChartObjects())
                lines, counts = update_page(page, chart_objects, plan)
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
        after = audit(xlsx_path, last_year)
        still = [
            s for s in after
            if (pages is None or s.page in pages) and s.status in STATUSES_CHANGED and s.proposed_values
        ]
        remaining = len(still)
        for s in still:
            results.append(
                f"STILL TO CHANGE {s.page} {s.chart_at} [{s.series_name or s.series_no}]: "
                f"{s.values} -> {s.proposed_values}"
            )
    flagged = totals["failed"] + totals["unreadable"] + remaining
    results += [
        "",
        f"Summary: {totals['changed']} changed, {flagged} flagged for review "
        f"({totals['failed']} could not be set, {totals['unreadable']} could not be read, "
        f"{remaining} still out of date after saving).",
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
