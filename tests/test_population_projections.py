"""
tests/test_population_projections.py

Tests for the population projections fetcher's parsers
(regional-indicators/fetchers/fetch_population_projections.py) and the
writer's planning (regional-indicators/transform/xlsx_update/
update_population_projections.py). No network or Excel is needed: the
sheets are given as rows of cell values.
"""

import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators"))
sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform" / "xlsx_update"))

from fetchers.fetch_population_projections import (  # noqa: E402
    loose_name, nsw_links, parse_nsw, parse_qld_lga, parse_qld_sa2,
)
from update_population_projections import plan_writes  # noqa: E402


# ── Parsers ───────────────────────────────────────────────────────────────────

def test_loose_names():
    assert loose_name("Central Highlands (R) (Qld)") == loose_name("Central Highlands")
    assert loose_name("Broadsound - Nebo") == loose_name("Broadsound-Nebo")
    assert loose_name("Isaac (R)") == "isaac"


def test_qld_lga_table():
    rows = [
        ("Projected population (medium series), by local government area", None, None, None),
        (None, None, None, None),
        ("Local Government Area (a)", "At 30 June", None, None),
        (None, "2016 (b)", 2021, 2026),
        ("— number —", None, None, None),
        ("Isaac (R)", 21563, 20761.79, 21556.4),
        ("Queensland", 4848877, 5193500.1, 5555600.2),
    ]
    assert parse_qld_lga(rows)["Isaac (R)"] == {"2016": 21563, "2021": 20761.79, "2026": 21556.4}


def test_qld_sa2_table_keeps_the_code():
    rows = [
        ("title",),
        (None, None, None, None, None, None, "2016 (b)", 2021, 2026),
        (312, "Mackay - Isaac - Whitsunday", 31201, "Bowen Basin - North", 312011338, "Broadsound - Nebo",
         8724, 7834.92, 7644.39),
        (307, "Darling Downs - Maranoa", 30701, "Darling Downs (West) - Maranoa", 307011177, "Roma Region",
         5872, 5726.42, 5599.4),
    ]
    sa2 = parse_qld_sa2(rows)
    assert sa2["Broadsound - Nebo"] == {"code": "312011338", "medium": {"2016": 8724, "2021": 7834.92, "2026": 7644.39}}
    assert sa2["Roma Region"]["code"] == "307011177"


def test_nsw_table_keeps_projected_years_only():
    rows = [
        ("2024 NSW Common Planning Assumption Projections",),
        ("Historic (2001-2021) and Projected (2022-2041) Population Totals (persons)",),
        ("Local Government Area", 2020, 2021, 2022, 2041),
        ("Narrabri", 12990, 12808, 12808, 12997),
    ]
    first, areas = parse_nsw(rows)
    assert first == 2022
    assert areas["Narrabri"] == {"2022": 12808, "2041": 12997}


def test_nsw_links_take_the_newest_edition():
    html = ('<a href="/sites/default/files/2022-12/2022-nsw-population-projections-sa2s.xlsx">'
            '<a href="/sites/default/files/2024-11/2024-nsw-population-projections-sa2s.xlsx">'
            '<a href="/sites/default/files/2024-11/2024-nsw-population-projections-local-government-areas.xlsx">')
    links = nsw_links(html)
    assert links["sa2s"] == ("https://www.planning.nsw.gov.au/sites/default/files/2024-11/"
                             "2024-nsw-population-projections-sa2s.xlsx", 2024)
    assert links["local-government-areas"][1] == 2024


# ── Writer ────────────────────────────────────────────────────────────────────

def town(name, state, sa2, lga, code):
    return SimpleNamespace(name=name, state=state, sa2_name=sa2, lga=lga, sa2_code=code, qgso_sa2=code)


TOWNS = [
    town("Moranbah", "QLD", "Moranbah", "Isaac", "312011341"),
    town("Emerald", "QLD", "Emerald", "Central Highlands (Qld)", "308011192"),
    town("Mitchell", "QLD", "Roma Surrounds", "Maranoa", "307011177"),
    town("Narrabri", "NSW", "Narrabri", "Narrabri", "111"),
    town("Singleton", "NSW", "Singleton", "Singleton", "222"),
]

MED = "Projected Population - Medium series: Persons Total (Persons), 2018ed"
SA2_ROW = "Population projection: district 2018e (Medium Series)"
ERP = "Population (ERP)"

# Columns: A label, B sub-label, then 2011, 2017, 2025, 2026, 2031 (columns 3-7).
# As on the sheet, the 2016 projection sits under 2011, 2021 under 2017
# and 2026 under 2025.
GRID = [
    ["LGA", None, 2011, 2017, 2025, 2026, 2031],
    ["Narrabri", None, None, None, None, None, None],
    [ERP, "Residents (LGA)", 13438, 13300, 12797, None, 12913],
    ["Singleton", None, None, None, None, None, None],
    [ERP, "Residents (LGA)", None, None, None, None, None],
    ["LGA", None, None, None, None, None, None],
    ["Isaac", None, None, None, None, None, None],
    [MED, None, 21563, 20762, 21556, None, 22709],
    ["Central Highlands", None, None, None, None, None, None],
    [MED, None, None, None, None, None, None],
    ["SA2", None, None, None, None, None, None],
    ["Moranbah", None, None, 0.03, None, None, None],     # a heading carrying a growth rate
    [SA2_ROW, "Residents projection (SA2)", 9029, 9287.55, 9919.12, None, 10590.22],
    ["Roma Surrounds", None, None, None, None, None, None],
    [SA2_ROW, "Residents projection (SA2)", None, None, None, None, None],
]

DATA = {
    "qld_2018": {
        "lga": {
            "Isaac (R)": {"medium": {"2016": 21563, "2021": 20761.79, "2026": 21556.4, "2031": 22709.3}},
            "Central Highlands (R) (Qld)": {"medium": {"2016": 28000, "2021": 28500.6, "2026": 29000.4,
                                                       "2031": 29500.5}},
        },
        "sa2": {
            "Moranbah": {"code": "312011341", "medium": {"2016": 9029, "2021": 9287.55, "2026": 9919.12,
                                                         "2031": 10590.22}},
            "Roma Region": {"code": "307011177", "medium": {"2016": 5872, "2021": 5726.42, "2026": 5599.4,
                                                            "2031": 5485.35}},
        },
    },
    "nsw": {
        "lga": {"Narrabri": {"2026": 12858, "2031": 12913}, "Singleton": {"2026": 25590, "2031": 26527}},
        "sa2": {},
    },
}


def written(grid=GRID, data=DATA):
    writes, lines = plan_writes(grid, data, TOWNS, last_data_year=2025)
    return {(r, c): v for r, c, v in writes}, lines


def test_qld_lga_projection_goes_where_the_filled_rows_put_each_year():
    writes, _ = written()
    # Central Highlands (row 10): 2016 under 2011, 2021 under 2017, 2026
    # under 2025, 2031 under 2031; whole numbers, as Isaac's row holds.
    assert {k: v for k, v in writes.items() if k[0] == 10} == {
        (10, 3): 28000, (10, 4): 28501, (10, 5): 29000, (10, 7): 29500}


def test_sa2_found_by_code_and_decimals_kept():
    writes, _ = written()
    # Roma Surrounds was "Roma Region" in the 2018 edition (same code).
    assert writes[(15, 4)] == 5726.42 and writes[(15, 7)] == 5485.35


def test_nsw_future_years_only_those_the_filled_rows_show():
    writes, _ = written()
    # Narrabri shows 2031 only (not 2026), so Singleton gets 2031 only;
    # the historic columns are never touched.
    assert {k: v for k, v in writes.items() if k[0] == 5} == {(5, 7): 26527}


def test_filled_cells_are_kept_and_differences_reported():
    grid = [list(r) for r in GRID]
    grid[9][2] = 27000          # Central Highlands 2016 entered by hand, different
    writes, lines = written(grid)
    assert (10, 3) not in writes
    assert any("Central Highlands" in line and "kept" in line and "2016" in line for line in lines)


def test_rows_of_other_areas_and_queensland_erp_rows_are_left_alone():
    writes, _ = written()
    assert not any(r in (3, 8, 13) for r, _ in writes)
