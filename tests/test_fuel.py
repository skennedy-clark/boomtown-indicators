"""
tests/test_fuel.py -- RACQ fuel prices: parsing the Annual Fuel Price
Report's RULP table text, and writing into the Exogenous sheet's Fuel
section.

No network, no PDF, no real workbook, no Excel. The parser is fed text
in the shape pypdfium2 gives for the real report (one line per table
row, single-space separated); the writer runs against a tiny sheet
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
    assert locations["Whitsunday"] == {2025: 172.1, 2024: 179.7}   # 'nd' left out


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
        (None, 2001, 2002, 2003, 2004),                  # sheet row 1 -- NOT the Fuel header
        ("Rainfall",),
        ("Dalby",),                                      # a Dalby block in ANOTHER section
        ("Dalby Airport", 500.0, 600.0, 700.0, 800.0),
        (None,),
        ("Fuel", None, 2022, 2023, 2024),                # Fuel's own header row: years start at C
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
        (6, 6): 2025,      # year appended to the FUEL header row (F6), after 2024
        (8, 6): 185.2,     # Brisbane
        (10, 6): 171.2,    # Dalby -- the Fuel block, not the Rainfall one
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
    assert (10, 4) not in fuel_sheet.writes            # the 2023 cell is untouched
