"""
tests/test_history_fetchers.py

Tests for the history fetchers used when a town is added:
regional-indicators/fetchers/fetch_business_history.py (ABS business
counts) and fetch_income_history.py (ATO postcode tables). The tables
are built in memory in the layouts the sources have used.
"""

import openpyxl

from fetchers.fetch_business_history import band_of, parse_release
from fetchers.fetch_income_history import column_map, postcode_rows, year_figures


# ── Business counts ────────────────────────────────────────────────────────────

def test_band_of_combines_finer_bands():
    assert band_of("Zero to less than $50k") == "0k-50k"
    assert band_of("$50k to less than $100k") == "50k-200k"
    assert band_of("$100k to less than $200k") == "50k-200k"
    assert band_of("$200k to less than $500k") == "200k-2m"
    assert band_of("$500k to less than $2m") == "200k-2m"
    assert band_of("$2m or more") == "2m+"
    assert band_of("$10m or more") == "2m+"
    assert band_of("Total") is None
    assert band_of("SA2 Code") is None


def _release(tmp_path, tables):
    """A release file with a Contents sheet and one "Table N" sheet per year."""
    wb = openpyxl.Workbook()
    contents = wb.active
    contents.title = "Contents"
    for n, (year, _) in enumerate(tables, start=1):
        contents.append([n, f"Businesses by Industry Division by SA2 by Turnover Size Ranges, June {year}"])
    for n, (year, rows) in enumerate(tables, start=1):
        ws = wb.create_sheet(f"Table {n}")
        ws.append(["title"])
        ws.append(["Industry", "Industry", "SA2", "SA2", "Zero to less than $50k", "$50k to less than $100k",
                   "$100k to less than $200k", "$200k to less than $2m", "$2m or more", "Total"])
        ws.append(["Code", "Label", "Code", "Label", "no.", "no.", "no.", "no.", "no.", "no."])
        for row in rows:
            ws.append(row)
    path = tmp_path / "release.xlsx"
    wb.save(path)
    return path


def test_parse_release_splits_primary_production_and_combines_bands(tmp_path):
    rows = [
        ["A", "Agriculture", 308011192, "Emerald", 10, 1, 2, 3, 4, 20],
        ["B", "Mining", 308011192, "Emerald", 5, 1, 1, 1, 1, 9],
        ["C", "Manufacturing", 308011192, "Emerald", 5, 0, 1, 1, 1, 8],
        ["C", "Manufacturing", 999999999, "Elsewhere", 50, 50, 50, 50, 50, 250],
    ]
    path = _release(tmp_path, [(2025, rows), (2024, rows[:1])])
    tables = parse_release(path, {"308011192"})
    assert sorted(tables) == [2024, 2025]
    emerald = tables[2025]["308011192"]
    assert emerald["pp"] == {"0k-50k": 10, "50k-200k": 3, "200k-2m": 3, "2m+": 4}
    assert emerald["npp"] == {"0k-50k": 10, "50k-200k": 3, "200k-2m": 2, "2m+": 2}
    assert "999999999" not in tables[2025]


# ── ATO postcode tables ────────────────────────────────────────────────────────

def test_column_map_reads_units_from_the_row_below():
    heading = ["Postcode", "Taxable individuals", "Taxable income or loss", "Mean taxable income",
               "Salary or wages", None]
    units = [None, "no.", "$", "$", "no.", "$"]
    cols = column_map(heading, units)
    assert cols["postcode"] == 0 and cols["taxable_individuals"] == 1
    assert cols["taxable_$"] == 2 and cols["taxable_mean"] == 3
    assert cols["wages_no"] == 4 and cols["wages_$"] == 5


def test_column_map_reads_units_from_the_heading():
    heading = ["State", "Postcode", "Salary or wages\nno.", "Salary or wages\n$",
               "Total business income non primary production\nno.",
               "Total business income non primary production\n$",
               "Total business income - primary production\nno.", "Total business income3\nno."]
    cols = column_map(heading, None)
    assert cols["wages_no"] == 2 and cols["wages_$"] == 3
    assert cols["bus_npp_no"] == 4 and cols["bus_npp_$"] == 5
    assert cols["bus_pp_no"] == 6 and cols["bus_total_no"] == 7


def _row(status, **values):
    return {"sheet": "s", "postcode": "4720", "status": status, **values}


def test_year_figures_prefer_the_all_individuals_table():
    rows = [
        _row("Taxable", taxable_no=100.0, **{"taxable_$": 5_000_000.0}, wages_no=90.0, **{"wages_$": 1.0}),
        _row("Non Taxable", taxable_no=20.0, **{"taxable_$": 10.0}, wages_no=10.0, **{"wages_$": 1.0}),
        _row("All", taxable_no=120.0, **{"taxable_$": 5_000_010.0}, wages_no=100.0, **{"wages_$": 2_000.0}),
    ]
    figures = year_figures(rows)
    assert figures["avg_income_taxable"] == 50_000
    assert figures["earners_no"] == 100 and figures["wages_total"] == 2_000


def test_year_figures_add_up_taxable_status_rows_without_a_total():
    rows = [
        _row("Taxable", taxable_no=100.0, **{"taxable_$": 5_000_000.0}, wages_no=90.0, **{"wages_$": 900.0}),
        _row("Non Taxable", taxable_no=20.0, **{"taxable_$": 10.0}, wages_no=10.0, **{"wages_$": 100.0}),
    ]
    figures = year_figures(rows)
    assert figures["earners_no"] == 100 and figures["wages_total"] == 1_000


def test_year_figures_early_years_are_taxable_individuals_only():
    rows = [_row("All", taxable_individuals=3815.0, **{"taxable_$": 145_479_842.0}, taxable_mean=38144.0,
                 wages_no=3310.0, **{"wages_$": 113_650_296.0})]
    figures = year_figures(rows)
    assert figures == {"avg_income_taxable": 38144, "earners_no": 3310, "wages_total": 113_650_296}


def test_postcode_rows_reads_a_sheet(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Individuals Table 6B"
    ws.append(["title"])
    ws.append(["State", "Postcode", "Salary or wages\nno.", "Salary or wages\n$"])
    ws.append(["QLD", 4720, 8944, 773878340])
    ws.append(["QLD", 4721, 1, 1])
    path = tmp_path / "t6.xlsx"
    wb.save(path)
    rows = postcode_rows(path, {"4720"})
    assert len(rows) == 1 and rows[0]["status"] == "All" and rows[0]["wages_no"] == 8944
