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
NRW_A = "Population projection: NRW 2022e (Series A)"
NRW_B = "Population projection: NRW 2022e (Series B)"
ERP = "Population (ERP)"

YEARS = [2011, 2016, 2017, 2019, 2021, 2022, 2023, 2024, 2025, 2026, 2031]
COL = {y: i + 3 for i, y in enumerate(YEARS)}          # 1-based columns: A label, B sub-label


def row(label, sub=None, **by_year):
    cells = [None] * len(YEARS)
    for key, value in by_year.items():
        cells[YEARS.index(int(key[1:]))] = value
    return [label, sub, *cells]


def grid():
    return [
        ["LGA", None, *YEARS],
        row("Narrabri"),
        row(ERP, "Residents (LGA)", y2011=13438, y2025=12797, y2031=12913),
        row("Singleton"),
        row(ERP, "Residents (LGA)"),
        row("LGA"),
        row("Isaac"),
        # As in the workbooks to 2026: 2016, 2021, 2026 under 2011, 2017, 2025.
        row(MED, y2011=21563, y2017=20762, y2025=21556, y2031=22709),
        # 2022-2025 under 2019, 2021, 2023, 2024; 2026 under its own year.
        row(NRW_A, "NRW Projection A (2022)", y2019=11140, y2021=11510, y2023=11270, y2024=11060, y2026=11060),
        row("Central Highlands"),
        row(MED),
        row(NRW_A, "NRW Projection A (2022)"),
        row(NRW_B, "NRW Projection B (2022)"),
        row("SA2"),
        row("Moranbah", y2017=0.03),                       # a heading carrying a growth rate
        row(SA2_ROW, "Residents projection (SA2)", y2011=9029, y2017=9287.55, y2025=9919.12, y2031=10590.22),
        row("Roma Surrounds"),
        row(SA2_ROW, "Residents projection (SA2)"),
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
    "nrw_2022": {"lga": {
        "Isaac": {"estimate_year": 2021, "table": "Table 4", "not_published": "",
                  "series": {"Series A": {"2021": 12080, "2022": 11140, "2023": 11510, "2024": 11270,
                                          "2025": 11060, "2026": 11060}}},
        "Central Highlands": {"estimate_year": 2021, "table": "Table 3",
                              "not_published": "(a) The Series B projection ... is not published at the LGA level",
                              "series": {"Series A": {"2021": 4100, "2022": 4200, "2023": 4230, "2024": 4230,
                                                      "2025": 4120, "2026": 4120}}},
    }},
    "nsw": {
        "lga": {"Narrabri": {"2026": 12858, "2031": 12913}, "Singleton": {"2026": 25590, "2031": 26527}},
        "sa2": {},
    },
}


def plan(g=None, data=DATA):
    g = g or grid()
    writes, moves, lines = plan_writes(g, data, TOWNS, last_data_year=2025)
    return {(r, c): v for r, c, v in writes}, moves, lines, g


def test_2018_edition_figures_are_moved_to_their_own_years():
    _, moves, lines, g = plan()
    isaac = g[7]
    assert isaac[COL[2016] - 1] == 21563 and isaac[COL[2021] - 1] == 20762 and isaac[COL[2026] - 1] == 21556
    assert isaac[COL[2011] - 1] is None and isaac[COL[2017] - 1] is None and isaac[COL[2025] - 1] is None
    assert isaac[COL[2031] - 1] == 22709                    # already under its own year
    assert (8, COL[2011], COL[2016]) in moves
    assert any("Isaac (row 8): MOVED" in line for line in lines)


def test_non_resident_projections_are_moved_and_the_estimate_is_not_shown():
    _, moves, _, g = plan()
    nrw = g[8]
    assert [nrw[COL[y] - 1] for y in (2022, 2023, 2024, 2025, 2026)] == [11140, 11510, 11270, 11060, 11060]
    assert nrw[COL[2019] - 1] is None and nrw[COL[2021] - 1] is None


def test_a_new_lga_is_written_under_its_own_years():
    writes, _, lines, _ = plan()
    assert {k: v for k, v in writes.items() if k[0] == 11} == {
        (11, COL[2016]): 28000, (11, COL[2021]): 28501, (11, COL[2026]): 29000, (11, COL[2031]): 29500}
    assert {k: v for k, v in writes.items() if k[0] == 12} == {
        (12, COL[y]): v for y, v in ((2022, 4200), (2023, 4230), (2024, 4230), (2025, 4120), (2026, 4120))}
    # Series B is not published for Central Highlands: left empty and reported.
    assert not any(k[0] == 13 for k in writes)
    assert any("Series B not published" in line for line in lines)


def test_sa2_found_by_code_and_decimals_kept():
    writes, _, _, _ = plan()
    # Roma Surrounds was "Roma Region" in the 2018 edition (same code).
    assert writes[(18, COL[2021])] == 5726.42 and writes[(18, COL[2031])] == 5485.35


def test_nsw_future_years_only_those_the_filled_rows_show():
    writes, moves, _, _ = plan()
    # Narrabri shows 2031 only (not 2026), so Singleton gets 2031 only;
    # New South Wales rows are never moved.
    assert {k: v for k, v in writes.items() if k[0] == 5} == {(5, COL[2031]): 26527}
    assert not any(m[0] in (3, 5) for m in moves)


def test_a_row_not_matching_the_source_is_not_moved():
    g = grid()
    g[7][COL[2017] - 1] = 20000                             # entered by hand, not the source's 2021
    _, moves, lines, g = plan(g)
    assert not any(m[0] == 8 for m in moves)
    assert g[7][COL[2011] - 1] == 21563
    assert any("Isaac (row 8): NOT MOVED" in line for line in lines)


def test_a_move_into_a_column_holding_another_figure_is_refused():
    g = grid()
    g[7][COL[2016] - 1] = 99999
    _, moves, lines, _ = plan(g)
    assert not any(m[0] == 8 for m in moves)
    assert any("Isaac (row 8): NOT MOVED" in line for line in lines)


def test_a_workbook_already_in_order_is_left_alone():
    writes, _, _, g = plan()
    for (r, c), value in writes.items():
        g[r - 1][c - 1] = value
    writes, moves, _, _ = plan(g)
    assert moves == [] and writes == {}


# ── Non-resident workers: history from the source register ───────────────────

def test_register_history_reads_verified_town_figures_only(tmp_path):
    import importlib, types
    sys.modules.setdefault("xlwings", types.ModuleType("xlwings"))
    nrw = importlib.import_module("update_population_nrw")
    path = tmp_path / "register.csv"
    path.write_text(
        "geography,geography_level,reference_year,measure,value,unit,data_type,projection_series,"
        "projection_vintage,source_url,source_table_page,verification_status,notes\n"
        "Emerald,UCL,2011,non-resident workers on-shift,850,persons,observed,,,u,t,verified,\n"
        "Emerald,UCL,2019,non-resident workers on-shift,,,observed,,,,,not located,\n"
        "Mitchell,UCL,2014,non-resident workers on-shift,,,observed,,,u,t,not published,\n"
        "Central Highlands,LGA,2023,non-resident workers on-shift,4230,persons,projection,Series A,v,u,t,verified,\n",
        encoding="utf-8")
    assert nrw.register_history(path) == {"Emerald": {2011: 850}}
    assert nrw.register_history(tmp_path / "missing.csv") == {}
