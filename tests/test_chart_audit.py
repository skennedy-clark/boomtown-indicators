"""
tests/test_chart_audit.py

Tests for the chart series audit
(regional-indicators/transform/chart_audit.py).

The workbook is built here with openpyxl: a data sheet with a calendar
year header, a second data sheet with financial-year labels and a
section that has its own header, and a page of charts whose series
cover each situation the audit classifies. No project workbook and no
Excel are needed.
"""

import csv
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform"))


def _chart(page, anchor, series):
    """Add a line chart whose series are (name, values ref, categories ref)."""
    from openpyxl.chart import LineChart
    from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
    from openpyxl.chart.series import Series, SeriesLabel

    chart = LineChart()
    for name, values, categories in series:
        ser = Series(val=NumDataSource(numRef=NumRef(f=values)), tx=SeriesLabel(v=name))
        if categories:
            ser.cat = AxDataSource(strRef=StrRef(f=categories))
        chart.series.append(ser)
    page.add_chart(chart, anchor)


@pytest.fixture
def workbook(tmp_path):
    import openpyxl

    wb = openpyxl.Workbook()

    pop = wb.active
    pop.title = "Population"
    years = list(range(2016, 2031))                                   # C..Q; 2025 is column L
    pop.append([None, None] + [f"{y - 1}/{str(y)[2:]}" for y in years])
    pop.append(["SA2", None] + years)
    pop.append(["Roma", None])
    pop.append(["Population (ERP)", "Residents (SA2)"] + [7000 + i for i in range(10)])           # to 2025
    pop.append(["Projected population", "Projection"] + [None] * 5 + [7100 + i for i in range(10)])  # 2021-2030
    pop.append(["Miles", None])
    pop.append(["Population (ERP)", "Residents (SA2)"] + [1800 + i for i in range(9)])            # to 2024
    pop.append(["Population (UCL)", "Residents (UCL)"] + [1200 + i for i in range(8)] + [None, None, 1250])

    biz = wb.create_sheet("Business")
    biz.append([None] + [f"{y - 1}/{str(y)[2:]}" for y in range(2017, 2026)])    # B..J; 2024/25 is column J
    biz.append(["Roma", 2042, 1991, 2010, 2031, 1998, 2044, 2003, 2019, 2026])    # totals that look like years
    biz.append(["0k-50k"] + [100 + i for i in range(9)])
    biz.append([None])
    biz.append(["Fuel"] + list(range(2019, 2026)))                                # section with its own header
    biz.append(["Roma"])
    biz.append(["Average price"] + [150 + i for i in range(7)])

    page = wb.create_sheet("Roma")
    _chart(page, "B2", [
        ("Residents", "Population!$C$4:$K$4", "Population!$C$1:$K$1"),          # data continues to 2025
        ("Projection", "Population!$H$5:$L$5", "Population!$H$1:$L$1"),         # continues past the cut-off
    ])
    _chart(page, "B20", [
        ("Miles", "Population!$C$7:$K$7", "Population!$C$1:$K$1"),              # up to date
        ("Miles UCL", "Population!$C$8:$K$8", "Population!$C$1:$Q$1"),          # last cells empty
    ])
    _chart(page, "K2", [
        ("0k-50k", "Business!$B$3:$I$3", "Business!$B$1:$I$1"),                 # financial-year header
        ("Fuel", "Business!$B$7:$G$7", "Business!$B$5:$G$5"),                   # own section header
    ])
    _chart(page, "K20", [
        ("Stitched", "(Population!$C$4:$J$4,Population!$N$4)", "(Population!$C$1:$J$1,Population!$N$1)"),
        ("Elsewhere", "Missing!$A$1:$C$1", None),
    ])

    path = tmp_path / "book.xlsx"
    wb.save(path)
    return path


@pytest.fixture
def audited(workbook):
    from chart_audit import audit

    return {s.series_name: s for s in audit(workbook, last_year=2025)}


def test_every_series_is_listed_with_its_page_and_position(workbook):
    from chart_audit import audit

    series = audit(workbook, last_year=2025)
    assert len(series) == 8
    assert {s.page for s in series} == {"Roma"}
    assert [s.chart_at for s in series][:2] == ["B2", "B2"]
    assert [s.series_no for s in series if s.chart_at == "B2"] == [1, 2]


def test_a_range_that_stops_before_the_new_year_is_extended(audited):
    s = audited["Residents"]
    assert s.status == "EXTEND"
    assert (s.range_from_year, s.range_to_year, s.data_to_year) == (2016, 2024, 2025)
    assert s.proposed_values == "Population!$C$4:$L$4"
    assert s.proposed_categories == "Population!$C$1:$L$1"
    assert (s.row_label, s.sub_label, s.block) == ("Population (ERP)", "Residents (SA2)", "Roma")


def test_an_up_to_date_range_is_left_alone(audited):
    s = audited["Miles"]
    assert s.status == "OK" and s.proposed_values == ""
    assert s.block == "Miles"


def test_projection_rows_are_not_extended_past_the_cut_off(audited):
    s = audited["Projection"]
    assert s.status == "PROJECTION" and s.proposed_values == ""
    assert s.data_to_year == 2030


def test_a_range_ending_on_empty_cells_is_reported(audited):
    s = audited["Miles UCL"]
    assert s.status == "ENDS_PAST_DATA"
    assert s.data_to_year == 2023
    assert "2026" in "; ".join(s.notes)                 # the figure after the gap is mentioned


def test_financial_year_labels_are_read_as_the_year_they_end_in(audited):
    """The row of totals directly above holds numbers that look like
    years and must not be taken for the header."""
    s = audited["0k-50k"]
    assert s.status == "EXTEND"
    assert (s.range_from_year, s.range_to_year, s.data_to_year) == (2017, 2024, 2025)
    assert s.proposed_values == "Business!$B$3:$J$3"


def test_a_section_is_read_against_its_own_year_header(audited):
    s = audited["Fuel"]
    assert s.status == "EXTEND"
    assert (s.range_from_year, s.range_to_year, s.data_to_year) == (2019, 2024, 2025)
    assert s.proposed_categories == "Business!$B$5:$H$5"


def test_separate_areas_on_one_row_get_a_single_range_proposed(audited):
    s = audited["Stitched"]
    assert s.status == "MULTI_AREA"
    assert s.proposed_values == "Population!$C$4:$L$4"
    assert s.proposed_categories == "Population!$C$1:$L$1"


def test_a_range_on_a_missing_sheet_is_unresolved(audited):
    s = audited["Elsewhere"]
    assert s.status == "UNRESOLVED" and "Missing" in "; ".join(s.notes)


@pytest.mark.parametrize("reference, expected", [
    ("Population!$C$4:$K$4", [("Population", 3, 4, 11, 4)]),
    ("'Data Sheet'!B2:D2", [("Data Sheet", 2, 2, 4, 2)]),
    ("(Crime!$D$12:$Z$12,Crime!$AB$12)", [("Crime", 4, 12, 26, 12), ("Crime", 28, 12, 28, 12)]),
    ("'It''s, here'!$A$1", [("It's, here", 1, 1, 1, 1)]),
    ("SomeDefinedName", None),
])
def test_references_are_split_into_areas(reference, expected):
    from chart_audit import parse_reference

    areas = parse_reference(reference)
    if expected is None:
        assert areas is None
    else:
        assert [(a.sheet, a.col1, a.row1, a.col2, a.row2) for a in areas] == expected


def test_the_csv_has_one_line_per_series_and_the_workbook_is_unchanged(workbook, tmp_path, capsys):
    from chart_audit import main

    before = workbook.read_bytes()
    out = tmp_path / "audit.csv"
    assert main([str(workbook), "--out", str(out), "--last-year", "2025", "--reference", str(workbook)]) == 0
    with open(out, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    assert len(rows) == 8
    assert {r["status"] for r in rows} == {"OK", "EXTEND", "PROJECTION", "ENDS_PAST_DATA", "MULTI_AREA", "UNRESOLVED"}
    assert all(r["reference_values"] == r["values"] for r in rows)
    assert workbook.read_bytes() == before
    assert "8 of 8 series paired" in capsys.readouterr().out
