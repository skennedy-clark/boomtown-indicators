"""
tests/test_update_charts.py

Tests for bringing chart series up to date
(regional-indicators/transform/xlsx_update/update_charts.py).

The planning and formula handling are tested directly. Applying the
plan is tested against small stand-ins for Excel's chart, series,
chart group and category objects, which record what is assigned to
them; no Excel is needed. The audit that feeds the plan is covered in
test_chart_audit.py.
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


def categories(*labels, filtered=()):
    return [SimpleNamespace(Name=label, IsFiltered=label in filtered) for label in labels]


def fake_chart(address, *series, cats=()):
    group = SimpleNamespace(FullCategoryCollection=lambda: FakeCollection(list(cats)))
    return SimpleNamespace(
        TopLeftCell=SimpleNamespace(Address=address),
        Chart=SimpleNamespace(
            SeriesCollection=lambda: FakeCollection(list(series)),
            ChartGroups=lambda: FakeCollection([group]),
        ),
    )


def audited(page, values, categories, status, proposed_values="", proposed_categories="",
            chart_at="B2", series_no=1, name="s", unhide=()):
    return SimpleNamespace(
        page=page, chart_at=chart_at, series_no=series_no, series_name=name, values=values,
        categories=categories, status=status, proposed_values=proposed_values,
        proposed_categories=proposed_categories, unhide_years=list(unhide),
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


@pytest.mark.parametrize("label, year", [(2025, 2025), ("2025", 2025), ("2024/25", 2025), ("2025.0", 2025),
                                         ("", None), ("Total", None)])
def test_category_labels_are_read_as_years(label, year):
    from update_charts import category_year

    assert category_year(label) == year


def test_plan_collects_extensions_and_years_to_show_per_chart():
    from update_charts import build_plan

    series = [
        audited("Roma", "Housing!$C$5:$AE$5", "Housing!$C$1:$AE$1", "UNHIDE", unhide=[2025], name="a"),
        audited("Roma", "Housing!$C$6:$AE$6", "Housing!$C$1:$AE$1", "UNHIDE", unhide=[2024, 2025], name="b"),
        audited("Roma", "Population!$C$4:$X$4", "Population!$C$1:$X$1", "EXTEND", "Population!$C$4:$AA$4",
                "Population!$C$1:$AA$1", chart_at="K2", unhide=[2023, 2024, 2025]),
        audited("Roma", "Income!$C$5:$Y$5", "Income!$C$1:$Y$1", "OK"),
        audited("Miles", "Housing!$C$9:$AE$9", "Housing!$C$1:$AE$1", "UNHIDE", unhide=[2025]),
    ]
    plan = build_plan(series)
    assert plan.unhide == {("Roma", "B2"): {2024, 2025}, ("Roma", "K2"): {2023, 2024, 2025}, ("Miles", "B2"): {2025}}
    assert [(c.chart_at, c.values) for c in plan.changes] == [("K2", "Population!$C$4:$AA$4")]
    assert build_plan(series, {"Miles"}).pages == {"Miles"}


def test_filtered_years_are_shown_and_other_filtered_years_stay_hidden():
    from update_charts import build_plan, update_page

    plan = build_plan([audited("Roma", "Housing!$C$5:$AE$5", "Housing!$C$1:$AE$1", "UNHIDE", unhide=[2025])])
    cats = categories("2023", "2024", "2025", "2026", "2027", filtered=("2025", "2026", "2027"))
    series = FakeSeries("=SERIES(,Housing!$C$1:$AE$1,Housing!$C$5:$AE$5,1)")
    results, counts = update_page("Roma", [fake_chart("$B$2", series, cats=cats)], plan)
    assert [c.IsFiltered for c in cats] == [False, False, False, True, True]
    assert series.Formula == "=SERIES(,Housing!$C$1:$AE$1,Housing!$C$5:$AE$5,1)"      # range untouched
    assert counts["charts_shown"] == 1 and counts["not_found"] == 0
    assert results == ["Roma B2: showing 2025"]


def test_financial_year_categories_are_matched_by_the_year_they_end_in():
    from update_charts import build_plan, update_page

    plan = build_plan([audited("Roma", "Business!$B$9:$T$9", "Business!$B$1:$T$1", "UNHIDE", unhide=[2025])])
    cats = categories("2023/24", "2024/25", filtered=("2024/25",))
    update_page("Roma", [fake_chart("$B$2", FakeSeries("=SERIES(,Business!$B$1:$T$1,Business!$B$9:$T$9,1)"),
                                    cats=cats)], plan)
    assert cats[1].IsFiltered is False


def test_a_range_is_extended_and_its_new_years_shown():
    from update_charts import build_plan, update_page

    plan = build_plan([audited("Roma", "Population!$C$4:$X$4", "Population!$C$1:$AQ$1", "EXTEND",
                               "Population!$C$4:$AA$4", "", chart_at="K2", unhide=[2023, 2024, 2025])])
    series = FakeSeries("=SERIES(Population!$B$4,Population!$C$1:$AQ$1,Population!$C$4:$X$4,1)")
    cats = categories("2022", "2023", "2024", "2025", "2026", filtered=("2024",))
    results, counts = update_page("Roma", [fake_chart("$K$2", series, cats=cats)], plan)
    assert series.Formula == "=SERIES(Population!$B$4,Population!$C$1:$AQ$1,Population!$C$4:$AA$4,1)"   # axis kept
    assert cats[2].IsFiltered is False
    assert (counts["extended"], counts["charts_shown"]) == (1, 1)


def test_a_series_is_matched_on_its_values_when_excel_reports_other_categories():
    from update_charts import build_plan, update_page

    plan = build_plan([audited("Roma", "Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND",
                               "Housing!$C$5:$AB$5", "Housing!$C$1:$AB$1")])
    series = FakeSeries("=SERIES(Housing!$B$5,,Housing!$C$5:$AA$5,1)")
    results, counts = update_page("Roma", [fake_chart("$B$2", series)], plan)
    assert counts["extended"] == 1
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
    assert (counts["extended"], counts["not_found"]) == (2, 0)


def test_a_series_whose_formula_cannot_be_read_is_found_by_name_and_given_ranges():
    from update_charts import build_plan, update_page

    plan = build_plan([audited("Roma", "Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND",
                               "Housing!$C$5:$AB$5", "Housing!$C$1:$AB$1", name="Sales")])
    series = FakeSeries(None, name="Sales", unreadable=True)
    results, counts = update_page("Roma", [fake_chart("$B$2", series)], plan, resolve=lambda ref: f"<{ref}>")
    assert counts["extended"] == 1
    assert (series.Values, series.XValues) == ("<Housing!$C$5:$AB$5>", "<Housing!$C$1:$AB$1>")
    assert "matched by chart position and series name" in results[0]


def test_a_range_that_cannot_be_set_is_reported_and_the_rest_continue():
    from update_charts import build_plan, update_page

    plan = build_plan([
        audited("Roma", "Housing!$C$5:$AA$5", "Housing!$C$1:$AA$1", "EXTEND", "Housing!$C$5:$AB$5", "", series_no=1),
        audited("Roma", "Housing!$C$6:$AA$6", "Housing!$C$1:$AA$1", "EXTEND", "Housing!$C$6:$AB$6", "", series_no=2),
    ])
    refused = FakeSeries("=SERIES(,Housing!$C$1:$AA$1,Housing!$C$5:$AA$5,1)", refuse=True)
    accepted = FakeSeries("=SERIES(,Housing!$C$1:$AA$1,Housing!$C$6:$AA$6,2)")
    results, counts = update_page("Roma", [fake_chart("$B$2", refused, accepted)], plan)     # no resolver
    assert (counts["failed"], counts["extended"]) == (1, 1)
    assert any("FAILED" in line for line in results)


def test_a_year_missing_from_the_axis_and_a_missing_chart_are_reported():
    from update_charts import build_plan, update_page

    plan = build_plan([
        audited("Roma", "Housing!$C$5:$AE$5", "Housing!$C$1:$AE$1", "UNHIDE", chart_at="B2", unhide=[2025]),
        audited("Roma", "Crime!$C$5:$AB$5", "Crime!$C$1:$AB$1", "UNHIDE", chart_at="T40", unhide=[2025]),
    ])
    cats = categories("2023", "2024")
    results, counts = update_page("Roma", [fake_chart("$B$2", FakeSeries("=SERIES(,A!$A$1:$B$1,A!$A$2:$B$2,1)"),
                                                      cats=cats)], plan)
    text = "\n".join(results)
    assert "Roma B2: no category for 2025" in text
    assert "NOT FOUND Roma T40: no chart at this position" in text
    assert counts["not_found"] == 1


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

    series = [audited("Roma", "Housing!$C$5:$AE$5", "Housing!$C$1:$AE$1", "UNHIDE", unhide=[2025]),
              audited("Roma", "Income!$C$5:$Y$5", "Income!$C$1:$Y$1", "OK")]
    monkeypatch.setattr(update_charts, "audit", lambda path, last_year: series)
    monkeypatch.setitem(sys.modules, "xlwings", None)                  # importing it would fail
    results, flagged = update_charts.update_charts(tmp_path / "book.xlsx", None, 2025, dry_run=True)
    text = "\n".join(results)
    assert "Roma B2: show 2025" in text and "1 charts would be changed" in text
    assert flagged == 0

    with pytest.raises(ValueError, match="No charts on page"):
        update_charts.update_charts(tmp_path / "book.xlsx", {"Nowhere"}, 2025, dry_run=True)
