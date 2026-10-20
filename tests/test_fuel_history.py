"""
tests/test_fuel_history.py

Tests for regional-indicators/fetchers/fetch_fuel_history.py (RACQ fuel
prices from every annual report) and its use by update_fuel.py to fill a
new row's earlier years.

No network, PDF or Excel is needed: the parser is given page text in the
form pypdfium2 produces for the reports, and the writer runs against a
small sheet read through tests/fake_xlwings_sheet.py.
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform" / "xlsx_update"))

from fetchers.fetch_fuel_history import (
    canonical_names, local_reports, merge_reports, parse_table, report_links, table_text,
)

MONTHS = "1 2 3 4 5 6 7 8 9 10 11 12"


# ── finding the reports ────────────────────────────────────────────────────────

def test_report_links_reads_old_and_new_file_names():
    html = (
        '<a href="/-/media/fuel-reports/racq-annual-fuel-price-report---2015.pdf?rev=a&amp;hash=b">x</a>'
        '<a href="/-/media/fuel-reports/2025/annual-fuel-price-report-2025-v3.pdf?rev=c&amp;hash=d">y</a>'
        '<a href="/-/media/fuel-reports/2016/monthly-fuel-price-report---july-2016.pdf">z</a>'
    )
    links = report_links(html)
    assert sorted(links) == [2015, 2025]
    assert links[2015] == "https://www.racq.com.au/-/media/fuel-reports/racq-annual-fuel-price-report---2015.pdf"


def test_local_reports_by_year(tmp_path):
    for name in ["racq-annual-fuel-price-report-2014.pdf", "annual-fuel-price-report-2025-v3.pdf", "other.pdf"]:
        (tmp_path / name).write_bytes(b"%PDF")
    assert sorted(local_reports(tmp_path)) == [2014, 2025]


# ── reading a report ───────────────────────────────────────────────────────────

def test_table_text_skips_mentions_and_runs_over_pages():
    pages = [
        "Table 5: Record High Monthly Average ULP Prices in Regional Queensland\nTown Record\nRoma 163.2 July 2008",
        "Table 14: Average ULP Prices for Queensland Regional Centres\nJan 2014 Feb 2014 ...\n"
        f"Roma {MONTHS} 157.3 153.4",
        f"Yeppoon {MONTHS} 150.0 nd\nTable 15: Average Diesel Prices for Queensland Regional Centres\n"
        f"Roma {MONTHS} 999.0 999.0",
    ]
    text = table_text(pages)
    assert text.startswith("Average ULP Prices for Queensland Regional Centres")
    assert "Yeppoon" in text and "999.0" not in text


def test_table_text_none_without_the_table():
    assert table_text(["Average Diesel Prices in Queensland\nJan-25"]) is None


def test_parse_table_years_count_back_from_the_report_year():
    text = (f"Average RULP Prices in Queensland\nJan-18 ... 2018 Ave 2017 Ave 2016 Ave\n"
            f"Charleville {MONTHS} nd nd 123.7\nCharters Towers {MONTHS} 148.3 132.1 124.0\n"
            f"Emerald {MONTHS} 152.4 140.8\n")
    locations, set_aside = parse_table(text, 2018)
    assert locations == {"Charleville": {2016: 123.7},
                         "Charters Towers": {2018: 148.3, 2017: 132.1, 2016: 124.0}}
    assert set_aside == 1                       # Emerald's row is short


# ── combining reports ──────────────────────────────────────────────────────────

def test_canonical_names_joins_split_and_abbreviated_names():
    reports = {2017: {"Towers": {2017: 132.1}, "Mt Isa": {2017: 142.7}},
               2018: {"Charters Towers": {2018: 148.3}, "Mount Isa": {2018: 155.9}}}
    canonical_names(reports)
    assert set(reports[2017]) == {"Charters Towers", "Mount Isa"}


def test_merge_takes_each_year_from_the_newest_report():
    reports = {
        2021: {"Charleville": {2016: 123.7, 2015: 131.4}, "Roma": {2021: 138.9, 2016: 123.7}},
        2025: {"Roma": {2025: 169.1, 2016: 123.8}},
        2014: {"Charleville": {2014: 158.9}},
    }
    locations, from_report, set_aside = merge_reports(reports)
    assert set_aside == {}
    assert locations["Roma"] == {2025: 169.1, 2021: 138.9, 2016: 123.8}
    assert from_report["Roma"][2016] == 2025
    assert locations["Charleville"] == {2016: 123.7, 2015: 131.4, 2014: 158.9}


def test_merge_sets_aside_a_report_that_disagrees_with_newer_ones():
    reports = {
        2012: {"Roma": {2012: 149.8, 2011: 146.4, 2010: 132.2}},
        2010: {"Roma": {2010: 130.3, 2009: 130.0}, "Normanton": {2010: 160.0}},   # misread layout
    }
    locations, _, set_aside = merge_reports(reports)
    assert list(set_aside) == [2010]
    assert "Normanton" not in locations and 2009 not in locations["Roma"]


# ── filling a new row ──────────────────────────────────────────────────────────

@pytest.fixture
def new_town_sheet(tmp_path):
    pytest.importorskip("xlwings")
    import openpyxl
    from fake_xlwings_sheet import FakeSheet

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Exogenous"
    for row in [
        (None, 2001),
        ("Fuel", 2014, 2015, 2016, 2017, 2018, 2019, 2020),
        ("Charleville",),                                    # new, not in the latest report
        ("Average RULP Price (cents)",),
        ("Roma",),                                           # existing row
        ("Average RULP Price (cents)", 157.3, 138.4, 120.0, 131.4, 149.1, 134.1),
    ]:
        ws.append(row)
    path = tmp_path / "exo.xlsx"
    wb.save(path)
    return FakeSheet(path, "Exogenous")


def test_new_row_takes_earlier_reports_where_its_source_has_no_figure(new_town_sheet):
    from update_fuel import write_fuel

    data = {"report_year": 2020, "locations": {"Roma": {"2019": 134.1, "2020": 121.3}}}
    open_data = {"report_year": 2020, "locations": {"Charleville": {"2019": 140.9, "2020": 125.0}}}
    history = {"locations": {"Charleville": {"2014": 158.9, "2015": 131.4, "2016": 123.7},
                             "Roma": {"2016": 123.7}}}
    results, written, flagged = write_fuel(new_town_sheet, data, open_data, history)
    assert (written, flagged) == (2, 0)
    charleville = {col: v for (row, col), v in new_town_sheet.writes.items() if row == 4}
    assert charleville == {2: 158.9, 3: 131.4, 4: 123.7, 7: 140.9, 8: 125.0}   # 2017-2018 left blank
    line = next(l for l in results if l.startswith("Charleville"))
    assert "from earlier RACQ reports: 2014-2016" in line
    assert (6, 4) not in new_town_sheet.writes                                  # Roma's existing 2016 kept
