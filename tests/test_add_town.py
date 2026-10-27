"""
tests/test_add_town.py

Tests for adding a town (regional-indicators/transform/add_town.py) and
for filling a new row's history in the writers
(regional-indicators/transform/xlsx_update/base.py, backfill_new_row).
No Excel is needed: the planner reads cell values, and the writers run
against tests/fake_xlwings_sheet.py.
"""

import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform"))
sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform" / "xlsx_update"))

from add_town import (  # noqa: E402
    make_plan, read_units, remap_formula, web_map_lines, area_name,
)


def town(name, sa2, lga, postcode="4000"):
    return SimpleNamespace(name=name, sa2_name=sa2, lga=lga, postcode=postcode)


ROMA = town("Roma", "Roma", "Maranoa", "4455")
MITCHELL = town("Mitchell", "Roma Surrounds", "Maranoa", "4465")
EMERALD = town("Emerald", "Emerald", "Central Highlands (Qld)", "4720")

# A small Population sheet: LGA, SA2 and UCL sections with blocks.
POPULATION = [
    (None, None, 2023, 2024),
    ("LGA", None, None, None),
    ("Maranoa", None, None, None),
    ("Population (ERP)", "Residents (LGA)", 12, 13),
    ("Western Downs", None, None, None),
    ("Population (ERP)", "Residents (LGA)", 30, 31),
    ("SA2", None, None, None),
    ("Roma", None, None, None),
    ("Population (ERP)", "Residents (SA2)", 7, 8),
    ("Roma Surrounds", None, None, None),
    ("Population (ERP)", "Residents (SA2)", 3, 4),
    ("UCL", None, None, None),
    ("Dalby", None, None, None),
    ("Estimated resident population", "Residents (UCL)", 12000, 12100),
    ("Roma", None, None, None),
    ("Estimated resident population", "Residents (UCL)", 6000, 6100),
    ("Tara", None, None, None),
    ("Estimated resident population", "Residents (UCL)", 900, 910),
]

# Employment: rows labelled by area, no blocks.
EMPLOYMENT = [
    (None, None, 2023, 2024),
    ("LGA", "Label", None, None),
    ("Maranoa", "Smoothed Unemployment rate (%)", 2.1, 2.2),
    ("SA2", None, None, None),
    ("Roma", "Smoothed Unemployment rate (%)", 2.5, 2.6),
    ("Roma Surrounds", "Smoothed Unemployment rate (%)", 1.8, 1.9),
    ("Tara", "Smoothed Unemployment rate (%)", 11.0, 11.2),
]


def test_units_are_blocks_or_rows():
    units = read_units("Population", POPULATION)
    assert [(u.section, u.label, u.first, u.last) for u in units] == [
        ("LGA", "Maranoa", 3, 4), ("LGA", "Western Downs", 5, 6),
        ("SA2", "Roma", 8, 9), ("SA2", "Roma Surrounds", 10, 11),
        ("UCL", "Dalby", 13, 14), ("UCL", "Roma", 15, 16), ("UCL", "Tara", 17, 18),
    ]
    rows = read_units("Employment", EMPLOYMENT)
    assert [(u.section, u.label, u.kind) for u in rows] == [
        ("LGA", "Maranoa", "row"), ("SA2", "Roma", "row"), ("SA2", "Roma Surrounds", "row"), ("SA2", "Tara", "row"),
    ]


def test_mitchell_shares_the_lga_reuses_the_sa2_and_gets_a_ucl_block():
    plan = make_plan({"Population": POPULATION, "Employment": EMPLOYMENT}, MITCHELL, ROMA)
    # Mitchell sorts before Roma, so its block goes in at Roma's row.
    assert [(i.sheet, i.template.section, i.label, i.at) for i in plan.inserts] == [("Population", "UCL", "Mitchell", 15)]
    assert {(t.sheet, t.section, e.label) for t, e in plan.reused} == {
        ("Population", "SA2", "Roma Surrounds"), ("Employment", "SA2", "Roma Surrounds")}
    assert {(u.sheet, u.label) for u in plan.shared} == {("Population", "Maranoa"), ("Employment", "Maranoa")}
    # Roma's UCL row (16) moves down two rows; the charts' copy reads Mitchell's row 16.
    assert plan.row_map[("Population", 18)] == 16
    # Roma's SA2 row maps to Roma Surrounds' row; neither moves.
    assert plan.row_map[("Population", 9)] == 11
    assert plan.row_map[("Employment", 5)] == 6


def test_emerald_gets_new_blocks_in_alphabetical_order():
    plan = make_plan({"Population": POPULATION, "Employment": EMPLOYMENT}, EMERALD, ROMA)
    inserted = {(i.sheet, i.template.section): (i.label, i.at) for i in plan.inserts}
    assert inserted[("Population", "LGA")] == ("Central Highlands", 3)     # before Maranoa
    assert inserted[("Population", "SA2")] == ("Emerald", 8)               # before Roma
    assert inserted[("Population", "UCL")] == ("Emerald", 15)              # between Dalby and Roma
    assert inserted[("Employment", "SA2")] == ("Emerald", 5)
    # Final rows: Roma's SA2 ERP row (9) moves down by the LGA and SA2 inserts (2 + 2).
    assert plan.shift("Population", 9) == 13
    assert plan.row_map[("Population", 13)] == 11


def test_remap_formula_changes_only_mapped_rows():
    row_map = {("Population", 16): 18, ("Population", 9): 11}
    formula = "=SERIES(Population!$B$16,Population!$C$1:$D$1,Population!$C$16:$D$16,1)"
    assert remap_formula(formula, row_map) == "=SERIES(Population!$B$18,Population!$C$1:$D$1,Population!$C$18:$D$18,1)"
    quoted = "=SERIES(,'Population'!$C$1:$D$1,'Population'!$C$9:$D$9,2)"
    assert remap_formula(quoted, row_map) == "=SERIES(,'Population'!$C$1:$D$1,'Population'!$C$11:$D$11,2)"
    benchmark = "=SERIES(,Housing!$C$1:$D$1,Housing!$C$124:$D$124,1)"
    assert remap_formula(benchmark, row_map) == benchmark


def test_web_map_lines(tmp_path):
    path = tmp_path / "map.csv"
    path.write_text(
        "folder,file,sheet,section,block,row_label,sub_label,note\n"
        "Roma,Population,Population,UCL,Roma,Estimated resident population,,\n"
        "Roma,Unemployment,Employment,SA2,,Roma,,\n"
        "Roma,House price,Housing,SA2,Roma,Median,,\n"
        "Dalby,Population,Population,UCL,Dalby,Estimated resident population,,\n",
        encoding="utf-8")
    lines, header = web_map_lines(path, MITCHELL, ROMA)
    assert [(l["folder"], l["block"], l["row_label"]) for l in lines] == [
        ("Mitchell", "Mitchell", "Estimated resident population"),
        ("Mitchell", "", "Roma Surrounds"),
        ("Mitchell", "Roma Surrounds", "Median"),
    ]


def test_web_map_lines_for_a_new_south_wales_town(tmp_path):
    # The NSW lines name the town in the section ("Narrabri (LGA)" housing
    # block) and in the State part of the Crime sheet; each is renamed.
    narrabri = town("Narrabri", "Narrabri", "Narrabri")
    singleton = town("Singleton", "Singleton", "Singleton")
    path = tmp_path / "map.csv"
    path.write_text(
        "folder,file,sheet,section,block,row_label,sub_label,note\n"
        "Narrabri,House sales,Housing,Narrabri (LGA),Sales No.,Total,,\n"
        "Narrabri,Crime rate,Crime,State,,Narrabri,,\n"
        "Narrabri,Theft,Crime,State,Narrabri,Theft,,\n"
        "Narrabri,Surrounds,Population,SA2,Narrabri Surrounds,Population (ERP),,\n",
        encoding="utf-8")
    lines, _ = web_map_lines(path, singleton, narrabri)
    assert [(l["section"], l["block"], l["row_label"]) for l in lines] == [
        ("Singleton (LGA)", "Sales No.", "Total"),
        ("State", "", "Singleton"),
        ("State", "Singleton", "Theft"),
        ("SA2", "Narrabri Surrounds", "Population (ERP)"),
    ]


def test_area_name_drops_state_suffix():
    assert area_name(EMERALD, "lga") == "Central Highlands"
    assert area_name(MITCHELL, "sa2") == "Roma Surrounds"


# ── Filling a new row's history ────────────────────────────────────────────────

def _sheet(tmp_path, rows):
    import openpyxl
    from fake_xlwings_sheet import FakeSheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for row in rows:
        ws.append(list(row))
    path = tmp_path / "s.xlsx"
    wb.save(path)
    return FakeSheet(path, "S")


def test_backfill_fills_an_empty_row(tmp_path):
    from base import backfill_new_row
    sheet = _sheet(tmp_path, [
        (None, None, 2021, 2022, 2023, 2024),
        ("Emerald", None, None, None, None, None),
        ("Rate", "label", None, None, None, None),
    ])
    filled = backfill_new_row(sheet, 3, 6, 2024, {2019: 1.0, 2021: 2.0, 2022: 3.0, 2023: 4.0, 2024: 5.0})
    assert filled == [2021, 2022, 2023]
    assert sheet.writes == {(3, 3): 2.0, (3, 4): 3.0, (3, 5): 4.0}


def test_backfill_leaves_a_row_with_figures(tmp_path):
    from base import backfill_new_row
    sheet = _sheet(tmp_path, [
        (None, None, 2021, 2022, 2023, 2024),
        ("Rate", "label", 2.0, None, None, None),
    ])
    assert backfill_new_row(sheet, 2, 6, 2024, {2022: 3.0, 2023: 4.0}) == []
    assert sheet.writes == {}


def test_backfill_follows_a_fiscal_header(tmp_path):
    from base import backfill_new_row
    # The writer puts the source's 2023 under "2022/23": a source year is the
    # fiscal year's end year, as the header reads it.
    sheet = _sheet(tmp_path, [
        (None, None, "2019/20", "2020/21", "2021/22", "2022/23"),
        ("Income", None, None, None, None, None),
    ])
    filled = backfill_new_row(sheet, 2, 6, 2023, {2020: 1.0, 2021: 2.0, 2022: 3.0, 2023: 4.0})
    assert filled == [2020, 2021, 2022]
    assert sheet.writes == {(2, 3): 1.0, (2, 4): 2.0, (2, 5): 3.0}
    # And where the writer puts the source's 2022 under "2022/23" (start year):
    sheet = _sheet(tmp_path, [
        (None, None, "2019/20", "2020/21", "2021/22", "2022/23"),
        ("Income", None, None, None, None, None),
    ])
    assert backfill_new_row(sheet, 2, 6, 2022, {2019: 1.0, 2020: 2.0, 2021: 3.0}) == [2019, 2020, 2021]
    assert sheet.writes == {(2, 3): 1.0, (2, 4): 2.0, (2, 5): 3.0}


# ── Changing the template ──────────────────────────────────────────────────────

def test_a_town_already_added_reuses_its_rows_with_another_template():
    from add_town import read_units
    # Mitchell already has a UCL block (added from Roma); copying from Tara now
    # adds nothing and points the charts at Mitchell's existing rows.
    rows = POPULATION[:16] + [("Mitchell", None, None, None),
                              ("Estimated resident population", "Residents (UCL)", None, None)] + POPULATION[16:]
    tara = town("Tara", "Tara", "Western Downs")
    plan = make_plan({"Population": rows}, MITCHELL, tara)
    assert plan.inserts == []
    assert [(t.label, e.label) for t, e in plan.reused if t.section == "UCL"] == [("Tara", "Mitchell")]
    assert plan.row_map[("Population", 20)] == 18


def test_replace_web_map(tmp_path):
    from add_town import replace_web_map
    path = tmp_path / "map.csv"
    path.write_bytes(
        b"folder,file,sheet,section,block,row_label,sub_label,note\r\n"
        b"Mitchell,Old,Population,UCL,Mitchell,x,,\r\n"
        b"Dalby,Population,Population,UCL,Dalby,y,,\r\n")
    header = ["folder", "file", "sheet", "section", "block", "row_label", "sub_label", "note"]
    new = [dict(zip(header, ["Mitchell", "New", "Population", "UCL", "Mitchell", "z", "", ""]))]
    assert replace_web_map(path, MITCHELL, new, header) == 1
    assert path.read_bytes() == (
        b"folder,file,sheet,section,block,row_label,sub_label,note\r\n"
        b"Dalby,Population,Population,UCL,Dalby,y,,\r\n"
        b"Mitchell,New,Population,UCL,Mitchell,z,,\r\n")


def test_block_heading_matches_lga_name_with_state_suffix():
    from base import same_area
    assert same_area("Central Highlands", "Central Highlands (Qld)")
    assert same_area("Roma", "Roma")
    assert not same_area("Toowoomba", "Toowoomba (West)")


def test_backfill_fills_a_row_holding_only_the_latest_year(tmp_path):
    from base import backfill_new_row
    sheet = _sheet(tmp_path, [
        (None, None, 2022, 2023, 2024),
        ("Rate", "label", None, None, 5.0),
    ])
    assert backfill_new_row(sheet, 2, 5, 2024, {2022: 3.0, 2023: 4.0, 2024: 5.0}) == [2022, 2023]


def test_backfill_uses_the_fuller_of_stacked_year_headers(tmp_path):
    from base import backfill_new_row
    # Income sheet: fiscal labels in row 1, calendar years for recent columns in row 2.
    sheet = _sheet(tmp_path, [
        (None, None, "2019/20", "2020/21", "2021/22", "2022/23", "2023/24"),
        ("Postcode", "Label", None, None, 2022, 2023, 2024),
        ("Earners", None, None, None, None, None, 9.0),
    ])
    assert backfill_new_row(sheet, 3, 7, 2024, {2020: 5.0, 2021: 6.0, 2022: 7.0, 2023: 8.0}) == [2020, 2021, 2022, 2023]


# ── Fixed axis limits on a copied chart ────────────────────────────────────────

def test_template_axis_limits_are_set_to_automatic():
    from types import SimpleNamespace
    from add_town import _free_template_axes

    primary = SimpleNamespace(MaximumScaleIsAuto=False, MaximumScale=500.0,
                              MinimumScaleIsAuto=False, MinimumScale=0.0)

    def axes(kind, group):
        if group == 2:
            raise RuntimeError("no secondary axis")
        return primary

    note = _free_template_axes(SimpleNamespace(Axes=axes))
    assert note == "value axis maximum 500"
    assert primary.MaximumScaleIsAuto is True and primary.MinimumScaleIsAuto is False


def test_automatic_axes_are_left_alone():
    from types import SimpleNamespace
    from add_town import _free_template_axes

    axis = SimpleNamespace(MaximumScaleIsAuto=True, MinimumScaleIsAuto=True)
    assert _free_template_axes(SimpleNamespace(Axes=lambda kind, group: axis)) == ""


# ── New South Wales layouts ────────────────────────────────────────────────────

NARRABRI = town("Narrabri", "Narrabri", "Narrabri", "2390")
MUSWELLBROOK = town("Muswellbrook", "Muswellbrook", "Muswellbrook", "2333")

# The Crime sheet: Queensland blocks under "State", then a part with its own
# year header in which each block carries the total rate on its name row.
CRIME = [
    (None, None, 2023, 2024, 2025),
    ("QPS Division", "Chart label", None, None, None),
    ("Roma", None, None, None, None),
    ("Drug Offences", "Drug Offences", 10.0, 11.0, 12.0),
    ("Theft", "Theft", 20.0, 21.0, 22.0),
    ("State", None, None, None, None),
    ("Queensland", None, None, None, None),
    ("Drug Offences", "Queensland (benchmark)", 9.0, 9.5, 9.7),
    ("Theft", "Queensland (benchmark)", 19.0, 19.5, 19.7),
    (None, None, None, None, None),
    (None, None, 2023, 2024, 2025),
    ("Narrabri", None, 83.9, 62.9, 66.1),
    ("Assault", None, 16.4, 12.6, 18.2),
    ("Theft", "Narrabri", 43.1, 30.7, 28.9),
    ("NSW", None, 105.9, 66.0, 48.6),
    ("Assault", None, 10.3, 8.0, 8.8),
    ("Theft", "NSW", 70.5, 28.0, 21.4),
]

# The Housing sheet: the State section, then Narrabri's block with its own
# year header and sub-headings.
HOUSING = [
    (None, None, 2023, 2024, 2025),
    ("State", None, None, None, None),
    ("Queensland", None, None, None, None),
    ("Detached dwelling: median sale price ($)", "Queensland (benchmark)", 700, 740, 838),
    (None, None, None, None, None),
    ("Narrabri (LGA)", None, 2022, 2023, 2024),
    ("Median Sales Price $", None, None, None, None),
    ("Total", "Median Sale Price", 360000, 382000, 390500),
    ("Sales No.", None, None, None, None),
    ("Total", "Sales No.", 180, 208, 237),
]


def test_nsw_crime_blocks_carry_figures_on_their_name_row():
    units = {u.label: u for u in read_units("Crime", CRIME)}
    assert (units["Narrabri"].first, units["Narrabri"].last, units["Narrabri"].kind) == (12, 14, "block")
    assert (units["NSW"].first, units["NSW"].last) == (15, 17)
    assert units["Narrabri"].section is None                  # not part of "State"
    assert (units["Roma"].first, units["Roma"].last) == (3, 5)


def test_housing_block_with_its_own_year_header_is_one_block():
    units = [u for u in read_units("Housing", HOUSING) if u.label == "Narrabri (LGA)"]
    assert len(units) == 1 and (units[0].first, units[0].last) == (6, 10)


def test_nsw_town_gets_crime_and_housing_blocks_named_like_the_template():
    plan = make_plan({"Crime": CRIME, "Housing": HOUSING}, MUSWELLBROOK, NARRABRI)
    inserts = {(i.sheet, i.label): i for i in plan.inserts}
    crime = inserts[("Crime", "Muswellbrook")]
    assert (crime.at, crime.rows, crime.template.first) == (12, 3, 12)    # before Narrabri, alphabetically
    housing = inserts[("Housing", "Muswellbrook (LGA)")]
    assert housing.template.label == "Narrabri (LGA)" and housing.rows == 5 + 1   # one blank row between blocks
    # Narrabri's rows, wherever they end up, are mapped to Muswellbrook's.
    assert plan.row_map[("Crime", plan.shift("Crime", 12))] == 12


def test_blocks_named_after_the_template_town_are_reported_not_copied():
    employment = [
        (None, None, 2024, 2025),
        ("SA2", None, None, None),
        ("Narrabri", "Smoothed Unemployment rate (%)", 4.1, 4.0),
        ("Narrabri Surrounds", "Smoothed Unemployment rate (%)", 3.2, 3.1),
        ("State", None, None, None),
        ("NSW", "Smoothed Unemployment rate (%)", 4.0, 4.2),
    ]
    plan = make_plan({"Employment": employment}, MUSWELLBROOK, NARRABRI)
    assert [i.label for i in plan.inserts] == ["Muswellbrook"]
    assert plan.related_in("Employment!$C$4:$D$4", final=False) == ["Employment: Narrabri Surrounds"]
    assert plan.related_in("Employment!$C$6:$D$6", final=False) == []          # NSW benchmark
    assert plan.related_in("Employment!$C$5:$D$5", final=True) == ["Employment: Narrabri Surrounds"]


def test_second_nsw_town_finds_the_template_after_a_cleared_block():
    # Muswellbrook was added first: its copied crime block has no figures.
    crime = CRIME[:11] + [
        ("Muswellbrook", None, None, None, None),
        ("Assault", None, None, None, None),
        ("Theft", "Muswellbrook", None, None, None),
    ] + CRIME[11:]
    singleton = town("Singleton", "Singleton", "Singleton", "2330")
    units = {u.label: u for u in read_units("Crime", crime)}
    assert (units["Muswellbrook"].first, units["Muswellbrook"].last) == (12, 14)
    assert (units["Narrabri"].first, units["Narrabri"].last) == (15, 17)
    plan = make_plan({"Crime": crime}, singleton, NARRABRI)
    assert [(i.label, i.template.label) for i in plan.inserts] == [("Singleton", "Narrabri")]


def test_business_heading_with_a_single_total_is_still_a_heading():
    business = [
        (None, "2023/24", "2024/25"),
        ("Non-primary production", None, None),
        ("Muswellbrook", None, None),
        ("0k-50k", None, None),
        ("2m+", None, None),
        ("Narrabri", None, 606),                 # a total in the latest year's column
        ("0k-50k", 110, 107),
        ("2m+", 34, 34),
        ("Roma", "=SUM(B10:B11)", 696),
        ("0k-50k", 168, 201),
        ("2m+", 66, 68),
    ]
    labels = [(u.label, u.first, u.last) for u in read_units("Business", business)]
    assert labels == [("Muswellbrook", 3, 5), ("Narrabri", 6, 8), ("Roma", 9, 11)]


def test_an_empty_labelled_row_added_earlier_is_reused():
    employment = [
        (None, None, 2024, 2025),
        ("SA2", None, None, None),
        ("Muswellbrook", "Smoothed Unemployment rate (%)", None, None),
        ("Narrabri", "Smoothed Unemployment rate (%)", 4.1, 4.0),
    ]
    plan = make_plan({"Employment": employment}, MUSWELLBROOK, NARRABRI)
    assert plan.inserts == [] and [e.label for _, e in plan.reused] == ["Muswellbrook"]
