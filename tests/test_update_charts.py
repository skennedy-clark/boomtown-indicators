"""
tests/test_update_charts.py

Tests for extending chart series ranges
(regional-indicators/transform/xlsx_update/update_charts.py).

The planning and formula handling are tested directly. Applying the
plan is tested against small stand-ins for Excel's ChartObject and
Series objects, which record the formula assigned to them; no Excel is
needed. The audit that feeds the plan is covered in test_chart_audit.py.
"""

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

ROOT = Path(__file__).parent.parent / "regional-indicators" / "transform"
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "xlsx_update"))


class FakeCollection:
    def __init__(self, items):
        self._items = items
        self.Count = len(items)

    def Item(self, index):
        return self._items[index - 1]


class FakeSeries:
    """Records the formula or ranges assigned to it. `refuse` makes the
    formula unassignable; `unreadable` makes it raise when read."""

    def __init__(self, formula, name="", refuse=False, unreadable=False):
        self._formula = formula
        self._refuse = refuse
        self._unreadable = unreadable
        self.Name = name
        self.Values = None
        self.XValues = None

    @property
    def Formula(self):
        if self._unreadable:
            raise RuntimeError("no formula")
        return self._formula

    @Formula.setter
    def Formula(self, value):
        if self._refuse or self._unreadable:
            raise RuntimeError("Excel refused")
        self._formula = value


def fake_chart(address, *series):
    return SimpleNamespace(
        TopLeftCell=SimpleNamespace(Address=address),
        Chart=SimpleNamespace(SeriesCollection=lambda: FakeCollection(list(series))),
    )


def audited(page, values, categories, status, proposed_values="", proposed_categories="",
            chart_at="B2", series_no=1, name="s"):
    return SimpleNamespace(
        page=page, chart_at=chart_at, series_no=series_no, series_name=name, values=values,
        categories=categories, status=status, proposed_values=proposed_values,
        proposed_categories=proposed_categories,
    )


@pytest.mark.parametrize("formula, expected", [
    ("=SERIES(Crime!$B$12,Crime!$D$1:$Z$1,Crime!$D$12:$Z$12,1)",
     ["Crime!$B$12", "Crime!$D$1:$Z$1", "Crime!$D$12:$Z$12", "1"]),
    ("=SERIES(,(Crime!$D$1:$Z$1,Crime!$AB$1),(Crime!$D$12:$Z$12,Crime!$AB$12),3)",
     ["", "(Crime!$D$1:$Z$1,Crime!$AB$1)", "(Crime!$D$12:$Z$12,Crime!$AB$12)", "3"]),
    ('=SERIES("Sales, total",\'Data, 2025\'!$B$1:$D$1,\'Data, 2025\'!$B$2:$D$2,2)',
     ['"Sales, total"', "'Data, 2025'!$B$1:$D$1", "'Data, 2025'!$B$2:$D$2", "2"]),
    ("=SUM(A1:A2)", None),
])
def test_series_formulas_are_split_into_arguments(formula, expected):
    from update_charts import split_series_formula

    assert split_series_formula(formula) == expected


def test_references_compare_equal_however_they_are_written():
    from update_charts import normalise

    assert normalise("(Crime!$D$12:$Z$12,Crime!$AB$12)") == normalise("=crime!D12:Z12, Crime!AB12")
    assert normalise("'Population'!$C$4:$K$4") == normalise("Population!C4:K4")
    assert normalise("Population!C4:K4") != normalise("Population!C4:L4")


def test_plan_holds_only_series_with_a_proposed_range():
    from update_charts import build_plan

    series = [
        audited("Roma", "Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND", "Housing!$C$5:$AB$5", "Housing!$C$1:$AB$1"),
        audited("Roma", "Income!$C$5:$Y$5", "Income!$C$1:$Y$1", "OK"),
        audited("Roma", "Population!$C$9:$AQ$9", "Population!$C$1:$AQ$1", "ENDS_PAST_DATA"),
        audited("Miles", "Housing!$C$9:$AA$9", "Housing!$C$1:$AA$1", "EXTEND", "Housing!$C$9:$AB$9", "Housing!$C$1:$AB$1"),
    ]
    assert len(build_plan(series)) == 2
    only_roma = build_plan(series, {"Roma"})
    assert [(c.values, c.categories) for c in only_roma] == [("Housing!$C$5:$AB$5", "Housing!$C$1:$AB$1")]


def test_matching_series_get_the_new_range_and_others_are_untouched():
    from update_charts import build_plan, update_page

    plan = build_plan([
        audited("Roma", "Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND",
                "Housing!$C$5:$AB$5", "Housing!$C$1:$AB$1"),
        audited("Roma", "(Crime!$D$12:$Z$12,Crime!$AB$12)", "(Crime!$D$1:$Z$1,Crime!$AB$1)", "MULTI_AREA",
                "Crime!$D$12:$AA$12", "Crime!$D$1:$AA$1", chart_at="K2"),
        audited("Roma", "Population!$C$4:$X$4", "Population!$C$1:$AQ$1", "EXTEND", "Population!$C$4:$AA$4", "",
                chart_at="B20"),
    ])
    extend = FakeSeries("=SERIES(Housing!$B$5,Housing!$C$1:$AA$1,Housing!$C$5:$AA$5,1)")
    stitched = FakeSeries("=SERIES(Crime!$B$12,(Crime!$D$1:$Z$1,Crime!$AB$1),(Crime!$D$12:$Z$12,Crime!$AB$12),2)")
    wide_axis = FakeSeries("=SERIES(Population!$B$4,Population!$C$1:$AQ$1,Population!$C$4:$X$4,1)")
    current = FakeSeries("=SERIES(Income!$B$5,Income!$C$1:$Y$1,Income!$C$5:$Y$5,1)")

    results, counts = update_page("Roma", [
        fake_chart("$B$2", extend, current), fake_chart("$K$2", stitched), fake_chart("$B$20", wide_axis),
    ], plan)

    assert extend.Formula == "=SERIES(Housing!$B$5,Housing!$C$1:$AB$1,Housing!$C$5:$AB$5,1)"
    assert stitched.Formula == "=SERIES(Crime!$B$12,Crime!$D$1:$AA$1,Crime!$D$12:$AA$12,2)"
    assert wide_axis.Formula == "=SERIES(Population!$B$4,Population!$C$1:$AQ$1,Population!$C$4:$AA$4,1)"   # axis kept
    assert current.Formula == "=SERIES(Income!$B$5,Income!$C$1:$Y$1,Income!$C$5:$Y$5,1)"
    assert (counts["changed"], counts["unchanged"], counts["not_found"]) == (3, 1, 0)
    assert (counts["charts"], counts["series"]) == (3, 4)
    assert any("K2" in line and "Crime!$D$12:$AA$12" in line for line in results)


def test_a_series_is_matched_on_its_values_when_excel_reports_other_categories():
    from update_charts import build_plan, update_page

    plan = build_plan([audited("Roma", "Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND",
                               "Housing!$C$5:$AB$5", "Housing!$C$1:$AB$1")])
    series = FakeSeries("=SERIES(Housing!$B$5,,Housing!$C$5:$AA$5,1)")
    results, counts = update_page("Roma", [fake_chart("$B$2", series)], plan)
    assert counts["changed"] == 1
    assert series.Formula == "=SERIES(Housing!$B$5,Housing!$C$1:$AB$1,Housing!$C$5:$AB$5,1)"
    assert "matched by values reference" in results[0]


def test_the_same_series_in_two_charts_is_told_apart_by_chart_position():
    from update_charts import build_plan, update_page

    refs = ("Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND", "Housing!$C$5:$AB$5", "Housing!$C$1:$AB$1")
    plan = build_plan([audited("Roma", *refs, chart_at="B2", name="Sales"),
                       audited("Roma", *refs, chart_at="K2", name="Sales")])
    first = FakeSeries("=SERIES(Housing!$B$5,Housing!$C$1:$AA$1,Housing!$C$5:$AA$5,1)", name="Sales")
    second = FakeSeries("=SERIES(Housing!$B$5,Housing!$C$1:$AA$1,Housing!$C$5:$AA$5,1)", name="Sales")
    _, counts = update_page("Roma", [fake_chart("$B$2", first), fake_chart("$K$2", second)], plan)
    assert (counts["changed"], counts["not_found"]) == (2, 0)
    assert first.Formula.endswith("$AB$5,1)") and second.Formula.endswith("$AB$5,1)")


def test_a_series_whose_formula_cannot_be_read_is_found_by_name_and_given_ranges():
    from update_charts import build_plan, update_page

    plan = build_plan([audited("Roma", "Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND",
                               "Housing!$C$5:$AB$5", "Housing!$C$1:$AB$1", chart_at="B2", name="Sales")])
    series = FakeSeries(None, name="Sales", unreadable=True)
    results, counts = update_page("Roma", [fake_chart("$B$2", series)], plan, resolve=lambda ref: f"<{ref}>")
    assert counts["changed"] == 1
    assert (series.Values, series.XValues) == ("<Housing!$C$5:$AB$5>", "<Housing!$C$1:$AB$1>")
    assert "matched by chart position and series name" in results[0]


def test_a_formula_excel_rejects_falls_back_to_assigning_the_ranges():
    from update_charts import build_plan, update_page

    plan = build_plan([audited("Roma", "Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND", "Housing!$C$5:$AB$5", "")])
    series = FakeSeries("=SERIES(,Housing!$C$1:$AA$1,Housing!$C$5:$AA$5,1)", refuse=True)
    _, counts = update_page("Roma", [fake_chart("$B$2", series)], plan, resolve=lambda ref: f"<{ref}>")
    assert counts["changed"] == 1
    assert series.Values == "<Housing!$C$5:$AB$5>" and series.XValues is None      # axis kept


def test_a_series_on_another_page_is_not_changed_by_this_pages_plan():
    from update_charts import build_plan, update_page

    plan = build_plan([audited("Miles", "Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND",
                               "Housing!$C$5:$AB$5", "Housing!$C$1:$AB$1")])
    series = FakeSeries("=SERIES(Housing!$B$5,Housing!$C$1:$AA$1,Housing!$C$5:$AA$5,1)", name="s")
    _, counts = update_page("Roma", [fake_chart("$B$2", series)], plan)
    assert counts["changed"] == 0 and series.Formula.endswith("$AA$5,1)")


def test_a_range_that_cannot_be_set_is_reported_and_the_rest_continue():
    from update_charts import build_plan, update_page

    plan = build_plan([
        audited("Roma", "Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND", "Housing!$C$5:$AB$5", "", series_no=1),
        audited("Roma", "Housing!$C$6:$AA$6", "Housing!$C$1:$AA$1", "EXTEND", "Housing!$C$6:$AB$6", "", series_no=2),
    ])
    refused = FakeSeries("=SERIES(,Housing!$C$1:$AA$1,Housing!$C$5:$AA$5,1)", refuse=True)
    accepted = FakeSeries("=SERIES(,Housing!$C$1:$AA$1,Housing!$C$6:$AA$6,2)")
    results, counts = update_page("Roma", [fake_chart("$B$2", refused, accepted)], plan)     # no resolver
    assert (counts["failed"], counts["changed"]) == (1, 1)
    assert any("FAILED" in line for line in results)


def test_a_planned_change_with_no_series_in_excel_is_reported_with_what_excel_has():
    from update_charts import build_plan, update_page

    plan = build_plan([
        audited("Roma", "Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND", "Housing!$C$5:$AB$5", "",
                chart_at="B2", name="Sales"),
        audited("Roma", "Crime!$C$5:$Z$5", "Crime!$C$1:$Z$1", "EXTEND", "Crime!$C$5:$AA$5", "",
                chart_at="T40", name="Drugs"),
    ])
    other = FakeSeries("=SERIES(,Sheet9!$A$1:$C$1,Sheet9!$A$2:$C$2,1)", name="Other")
    results, counts = update_page("Roma", [fake_chart("$B$2", other)], plan)
    text = "\n".join(results)
    assert counts["not_found"] == 2 and counts["changed"] == 0
    assert "NOT FOUND Roma B2 [Sales]" in text and "Sheet9!$A$2:$C$2" in text       # what Excel has at B2
    assert "Excel reports no chart at T40" in text


def test_charts_inside_grouped_shapes_are_included():
    from update_charts import MSO_GROUP, page_charts

    chart_a = SimpleNamespace(Type=3, HasChart=True)
    chart_b = SimpleNamespace(Type=3, HasChart=True)
    picture = SimpleNamespace(Type=13, HasChart=False)
    group = SimpleNamespace(Type=MSO_GROUP, GroupItems=FakeCollection([chart_b, picture]))
    sheet = SimpleNamespace(Shapes=FakeCollection([chart_a, picture, group]),
                            ChartObjects=lambda: FakeCollection([chart_a]))
    assert page_charts(sheet) == [chart_a, chart_b]


def test_dry_run_lists_changes_without_opening_excel(tmp_path, monkeypatch):
    import update_charts

    series = [audited("Roma", "Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND",
                      "Housing!$C$5:$AB$5", "Housing!$C$1:$AB$1"),
              audited("Roma", "Income!$C$5:$Y$5", "Income!$C$1:$Y$1", "OK")]
    monkeypatch.setattr(update_charts, "audit", lambda path, last_year: series)
    monkeypatch.setitem(sys.modules, "xlwings", None)                  # importing it would fail
    results, flagged = update_charts.update_charts(tmp_path / "book.xlsx", None, 2025, dry_run=True)
    text = "\n".join(results)
    assert "1 to change" in text and "Housing!$C$5:$AB$5" in text and "dry run" in text
    assert flagged == 0

    with pytest.raises(ValueError, match="No charts on page"):
        update_charts.update_charts(tmp_path / "book.xlsx", {"Nowhere"}, 2025, dry_run=True)
