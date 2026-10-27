"""
tests/test_town_coverage.py

Tests for the check of what is still missing for a town
(regional-indicators/transform/town_coverage.py). Sheets are given as
rows of cell values, as openpyxl returns them; no Excel is needed.
"""

import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform"))

from town_coverage import check_town, report, year_ranges  # noqa: E402


def town(name, state, sa2, lga):
    return SimpleNamespace(name=name, state=state, sa2_name=sa2, lga=lga)


ROMA = town("Roma", "QLD", "Roma", "Maranoa")
DALBY = town("Dalby", "QLD", "Dalby", "Western Downs")
EMERALD = town("Emerald", "QLD", "Emerald", "Central Highlands (Qld)")
NARRABRI = town("Narrabri", "NSW", "Narrabri", "Narrabri")
TOWNS = [ROMA, DALBY, EMERALD, NARRABRI]

POPULATION = [
    (None, None, 2022, 2023, 2024),
    ("LGA", None, None, None, None),
    ("Central Highlands", None, None, None, None),
    ("Population (ERP)", "Residents (LGA)", None, None, 28000),
    ("Maranoa", None, None, None, None),
    ("Population (ERP)", "Residents (LGA)", 12, 13, 14),
    ("Western Downs", None, None, None, None),
    ("Population (ERP)", "Residents (LGA)", 30, 31, 32),
    ("SA2", None, None, None, None),
    ("Dalby", None, None, None, None),
    ("Population (ERP)", "Residents (SA2)", 7, 8, 9),
    ("Roma", None, None, None, None),
    ("Population (ERP)", "Residents (SA2)", 3, 4, 5),
    ("UCL", None, None, None, None),
    # A name row carrying a working formula is not data.
    ("Dalby", "=(D16-C16)/C16", "=(E16-D16)/D16", None, None),
    ("Estimated resident population", "Residents (UCL)", 12000, 12100, 12200),
    ("Emerald", None, None, None, None),
    ("Estimated resident population", "Residents (UCL)", 14000, 14100, 14200),
    ("Roma", None, None, None, None),
    ("Estimated resident population", "Residents (UCL)", 6000, 6100, 6200),
]


def check(rows, subject, sheet="Population"):
    return check_town({sheet: rows}, subject, TOWNS, [subject.name], [])


def test_empty_years_that_the_peers_have_are_listed():
    result = check(POPULATION, EMERALD)
    assert len(result.gaps) == 1
    gap = result.gaps[0]
    assert (gap.section, gap.block, gap.years) == ("LGA", "Central Highlands", [2022, 2023])
    assert gap.peers == ["Maranoa", "Western Downs"]


def test_a_missing_block_is_listed():
    # The SA2 section has Dalby and Roma but no Emerald.
    result = check(POPULATION, EMERALD)
    assert result.missing == ["Population SA2: no block for Emerald (the section has Dalby, Roma)"]


def test_sections_without_peers_of_the_same_state_are_not_checked():
    # Narrabri is the only New South Wales town: no peers, nothing to compare.
    result = check(POPULATION, NARRABRI)
    assert result.missing == [] and result.gaps == []


def test_a_peer_row_empty_throughout_does_not_vote():
    # Western Downs's row empty throughout (a town just added): Maranoa's
    # figures alone decide.
    rows = list(POPULATION)
    rows[7] = ("Population (ERP)", "Residents (LGA)", None, None, None)
    result = check(rows, EMERALD)
    assert result.gaps[0].years == [2022, 2023] and result.gaps[0].peers == ["Maranoa"]


def test_formulas_count_as_filled_in_the_town_s_own_row():
    rows = list(POPULATION)
    rows[3] = ("Population (ERP)", "Residents (LGA)", "=D4*0.98", None, 28000)
    result = check(rows, EMERALD)
    assert result.gaps[0].years == [2023]


def test_report_and_year_ranges():
    assert year_ranges([2001, 2002, 2003, 2010, 2012, 2013]) == "2001-2003, 2010, 2012-2013"
    lines = report(check(POPULATION, EMERALD))
    assert lines[0] == "Emerald"
    assert any("Central Highlands" in line and "2022-2023" in line for line in lines)
