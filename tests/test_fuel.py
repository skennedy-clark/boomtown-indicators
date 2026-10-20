"""
tests/test_fuel.py

Tests for RACQ fuel prices: parsing the RULP table text of the Annual
Fuel Price Report, and writing into the Fuel section of the Exogenous
sheet.

No network, PDF, workbook file or Excel is needed. The parser is given
text in the form pypdfium2 produces for the report (one line per table
row, separated by single spaces). The writer runs against a small sheet
read through tests/fake_xlwings_sheet.py.
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform" / "xlsx_update"))

MONTHS = "170.1 170.2 170.3 170.4 170.5 170.6 170.7 170.8 170.9 171.0 171.1 171.2"

TABLE_TEXT = f"""Average RULP Prices in Queensland
Jan-25 Feb-25 Mar-25 Apr-25 May-25 Jun-25 Jul-25 Aug-25 Sep-25 Oct-25 Nov-25 Dec-25 2025
Ave
2024
Ave
2023
Ave
Brisbane {MONTHS} 185.2 194.5 193.1
Charters Towers {MONTHS} 187.3 191.4 192.3
Dalby {MONTHS} 171.2 175.6 178.4
Whitsunday {MONTHS} 172.1 179.7 nd
Source: RACQ calculations using OPIS data (2025 to 2018)
"""


# ── parser ─────────────────────────────────────────────────────────────────

def test_parse_reads_years_names_and_values():
    from fetchers.fetch_fuel import parse_rulp_table

    years, locations = parse_rulp_table(TABLE_TEXT)
    assert years == [2025, 2024, 2023]
    assert set(locations) == {"Brisbane", "Charters Towers", "Dalby", "Whitsunday"}
    assert locations["Dalby"] == {2025: 171.2, 2024: 175.6, 2023: 178.4}
    assert locations["Charters Towers"][2025] == 187.3          # two-word name
    assert locations["Whitsunday"] == {2025: 172.1, 2024: 179.7}   # 'nd' is omitted


def test_parse_refuses_rows_of_different_lengths():
    from fetchers.fetch_fuel import parse_rulp_table

    broken = TABLE_TEXT.replace("Dalby " + MONTHS + " 171.2 175.6 178.4",
                                "Dalby " + MONTHS + " 171.2 175.6")
    with pytest.raises(ValueError, match="inconsistent"):
        parse_rulp_table(broken)


def test_parse_needs_the_month_header():
    from fetchers.fetch_fuel import parse_rulp_table

    with pytest.raises(ValueError, match="Dec-"):
        parse_rulp_table("Average RULP Prices in Queensland\nDalby 1 2 3\n")


# ── writer ─────────────────────────────────────────────────────────────────

@pytest.fixture
def fuel_sheet(tmp_path):
    pytest.importorskip("xlwings")
    import openpyxl
    from fake_xlwings_sheet import FakeSheet

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Exogenous"
    for row in [
        (None, 2001, 2002, 2003, 2004),                  # sheet row 1, which is not the Fuel header
        ("Rainfall",),
        ("Dalby",),                                      # a Dalby block in another section
        ("Dalby Airport", 500.0, 600.0, 700.0, 800.0),
        (None,),
        ("Fuel", None, 2022, 2023, 2024),                # the Fuel section's own header row; years start at C
        ("Brisbane",),
        ("Average RULP Price (cents)", None, 184.9, 193.1, 194.5),
        ("Dalby ",),                                     # trailing space
        ("Average RULP Price (cents)", None, 181.9, 174.4, 175.6),
        ("Chinchilla",),                                 # not an RACQ location
        ("Average RULP Price (cents)", None, 180.0, 181.0, 182.0),
        (None,),
        ("average distance", 12100),
    ]:
        ws.append(row)
    path = tmp_path / "exo.xlsx"
    wb.save(path)
    return FakeSheet(path, "Exogenous")


DATA = {
    "report_year": 2025,
    "locations": {
        "Brisbane": {"2022": 184.9, "2023": 193.1, "2024": 194.5, "2025": 185.2},
        "Dalby":    {"2022": 181.9, "2023": 178.4, "2024": 175.6, "2025": 171.2},
    },
}


def test_writes_new_year_into_each_block_of_the_fuel_section(fuel_sheet):
    from update_fuel import write_fuel

    results, written, flagged = write_fuel(fuel_sheet, DATA)
    assert (written, flagged) == (2, 1)
    assert fuel_sheet.writes == {
        (6, 6): 2025,      # year appended to the Fuel header row (F6), after 2024
        (8, 6): 185.2,     # Brisbane
        (10, 6): 171.2,    # Dalby: the Fuel block, not the Rainfall one
    }


def test_block_with_no_racq_location_is_reported_and_left_alone(fuel_sheet):
    from update_fuel import write_fuel

    results, _, _ = write_fuel(fuel_sheet, DATA)
    assert any(line.startswith("Chinchilla: SKIPPED") for line in results)
    assert not any(row == 12 for row, _ in fuel_sheet.writes)


def test_earlier_year_that_differs_from_the_report_is_noted_not_changed(fuel_sheet):
    from update_fuel import write_fuel

    results, _, _ = write_fuel(fuel_sheet, DATA)
    dalby = next(line for line in results if line.startswith("Dalby"))
    assert "2023 sheet 174.4 vs report 178.4" in dalby
    assert (10, 4) not in fuel_sheet.writes            # the 2023 cell is not written


# ── Queensland fuel price open data (fetch_fuel_qld.py) ────────────────────

def test_open_data_fills_a_block_that_racq_does_not_name(fuel_sheet):
    from update_fuel import write_fuel

    open_data = {"report_year": 2025, "locations": {"Chinchilla": {"2024": 182.0, "2025": 179.5},
                                                    "Dalby": {"2025": 999.0}}}
    results, written, flagged = write_fuel(fuel_sheet, DATA, open_data)
    assert (written, flagged) == (3, 0)
    assert fuel_sheet.writes[(12, 6)] == 179.5
    assert fuel_sheet.writes[(10, 6)] == 171.2          # RACQ still wins where it names the town
    assert any(line.startswith("Chinchilla: WRITTEN") and "open data" in line for line in results)


OPEN_DATA_CSV = """SiteId,Site_Name,Site_Brand,Sites_Address_Line_1,Site_Suburb,Site_State,Site_Post_Code,Site_Latitude,Site_Longitude,Fuel_Type,Price,TransactionDateutc
1,A,B,x,EMERALD,QLD,4720,0,0,Unleaded,1800,01/01/2024 00:00
1,A,B,x,EMERALD,QLD,4720,0,0,Unleaded,2000,01/07/2024 00:00
1,A,B,x,EMERALD,QLD,4720,0,0,Diesel,2100,01/01/2024 00:00
2,C,D,y,EMERALD,QLD,4720,0,0,Unleaded,1900,01/01/2024 00:00
2,C,D,y,EMERALD,QLD,4720,0,0,Unleaded,9999,01/03/2024 00:00
2,C,D,y,EMERALD,QLD,4720,0,0,Unleaded,1910,01/04/2024 00:00
3,E,F,z,BRISBANE,QLD,4000,0,0,Unleaded,1500,01/01/2024 00:00
4,G,H,w,EMERALD,QLD,4720,0,0,Unleaded,1950,01/12/2024 00:00
"""


def test_open_data_annual_mean_weights_prices_by_days_and_drops_keying_errors():
    from collections import defaultdict
    from fetchers.fetch_fuel_qld import annual_means, read_changes

    changes = defaultdict(list)
    read_changes(OPEN_DATA_CSV, {"4720"}, changes)
    means = annual_means(changes, range(2024, 2025))
    mean, sites = means[("4720", 2024)]
    # Site 1: 182 days at 180.0, 184 at 200.0; site 2: 999.9 ignored, so 91 days
    # at 190.0 then 275 at 191.0; site 4 has one month only and does not count.
    assert sites == 2
    expected = ((182 * 180.0 + 184 * 200.0) / 366 + (91 * 190.0 + 275 * 191.0) / 366) / 2
    assert abs(mean - expected) < 1e-9


def test_open_data_file_month():
    from fetchers.fetch_fuel_qld import file_month
    assert file_month("Queensland Fuel Prices December 2019") == (2019, 12)
    assert file_month(" Queensland Fuel Prices  March 2024") == (2024, 3)
    assert file_month("fuel-prices-2025-06-changes-only.csv") == (2025, 6)
    assert file_month("Fuel price reporting 2019-2022") is None
