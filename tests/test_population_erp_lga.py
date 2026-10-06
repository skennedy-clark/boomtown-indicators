"""
tests/test_population_erp_lga.py -- LGA-level Population (ERP): config loading,
ABS CSV parsing, and the section-aware row finding the writer depends on.

No network and no real workbook: the sheet tests build a tiny workbook
with the same shape as the real Population sheet (an "LGA" section
header in ROW 2, sharing the row with the year headers, then an "SA2"
section further down, with the same block name and indicator name in
both) and read it through tests/fake_xlwings_sheet.py.
"""

import sys
from pathlib import Path

import pytest

from config import Config

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform" / "xlsx_update"))


# ── config: LGAs are derived from the towns ────────────────────────────────

LGA_TOML = """
[towns.chinchilla]
name = "Chinchilla"
state = "QLD"
postcode = "4413"
postcodes = ["4413"]
sa2_code = "307011172"
lga = "Western Downs"
qgso_lga = "LGA/37310"

[towns.dalby]
name = "Dalby"
state = "QLD"
postcode = "4405"
postcodes = ["4405"]
sa2_code = "307021183"
lga = "Western Downs"
qgso_lga = "LGA/37310"

[towns.brisbane]
name = "Brisbane"
state = "QLD"
lga = "Brisbane"
qgso_lga = "LGA/31000"
benchmark = true

[towns.narrabri]
name = "Narrabri"
state = "NSW"
postcode = "2390"
postcodes = ["2390"]
sa2_code = "110031197"
lga = "Narrabri"
lga_code = "15750"
"""


def _write(tmp_path, text):
    p = tmp_path / "towns.toml"
    p.write_text(text)
    return p


def test_lgas_derived_once_each_including_benchmarks(tmp_path):
    config = Config(_write(tmp_path, LGA_TOML))
    assert config.lgas() == [
        ("Western Downs", "37310", "QLD"),   # shared by two towns, listed once
        ("Brisbane", "31000", "QLD"),        # benchmark town still contributes its LGA
        ("Narrabri", "15750", "NSW"),        # no qgso_lga -- uses lga_code
    ]


def test_benchmark_town_stays_out_of_study_towns(tmp_path):
    config = Config(_write(tmp_path, LGA_TOML))
    assert "Brisbane" not in [t.name for t in config.study_towns()]


def test_towns_without_an_lga_code_are_skipped(sample_toml):
    assert Config(sample_toml).lgas() == []


def test_lga_code_must_be_digits(tmp_path):
    bad = LGA_TOML.replace('lga_code = "15750"', 'lga_code = "LGA/15750"')
    with pytest.raises(ValueError, match="digits only"):
        Config(_write(tmp_path, bad))


def test_same_lga_code_under_two_names_raises(tmp_path):
    bad = LGA_TOML.replace('qgso_lga = "LGA/31000"', 'qgso_lga = "LGA/37310"')
    with pytest.raises(ValueError, match="is named"):
        Config(_write(tmp_path, bad))


# ── ABS CSV parsing ────────────────────────────────────────────────────────

def test_parse_abs_csv():
    from fetchers.fetch_population_erp_lga import _parse_csv

    text = (
        "DATAFLOW,MEASURE,REGION_TYPE,REGION,FREQ,TIME_PERIOD,OBS_VALUE,UNIT_MEASURE,OBS_STATUS,OBS_COMMENT,REPYEAREND\n"
        "ABS:ERP_LGA2025(1.0.0),ERP,LGA2025,31000,A,2024,1353837,PSNS,,,30-6\n"
        "ABS:ERP_LGA2025(1.0.0),ERP,LGA2025,31000,A,2025,1375301,PSNS,,,30-6\n"
        "ABS:ERP_LGA2025(1.0.0),ERP,LGA2025,15750,A,2025,12797,PSNS,,,30-6\n"
        "ABS:ERP_LGA2025(1.0.0),ERP,LGA2025,15750,A,2026,,PSNS,,,30-6\n"
    )
    assert _parse_csv(text) == {
        "31000": {2024: 1353837, 2025: 1375301},
        "15750": {2025: 12797},
    }


# ── section-aware row finding (needs xlwings importable for base.py) ───────

@pytest.fixture
def population_like_sheet(tmp_path):
    pytest.importorskip("xlwings")
    import openpyxl
    from fake_xlwings_sheet import FakeSheet

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Population"
    rows = [
        (None, None, "2022/23", "2023/24", "2024/25"),
        ("LGA", None, 2023, 2024, 2025),                       # section header IN ROW 2
        ("Goondiwindi", None),
        ("Population (ERP)", "Residents (LGA)", 10452, 10495, None),
        ("SA2", None),
        ("Goondiwindi", None),
        ("Population (ERP)", "Residents (SA2)", 6230, 6240, None),
    ]
    for row in rows:
        ws.append(row)
    path = tmp_path / "pop.xlsx"
    wb.save(path)
    return FakeSheet(path, "Population")


def test_lga_section_header_in_row_2_is_seen(population_like_sheet):
    from base import _find_town_indicator_row

    find = lambda section: _find_town_indicator_row(
        population_like_sheet, "Goondiwindi", "Population (ERP)", None, section=section
    )
    assert find("LGA") == 4     # failed with "Could not find" before the 2026-10-06 fix
    assert find("SA2") == 7


def test_no_section_still_refuses_to_guess(population_like_sheet):
    from base import _find_town_indicator_row

    with pytest.raises(ValueError, match="AMBIGUOUS"):
        _find_town_indicator_row(population_like_sheet, "Goondiwindi", "Population (ERP)", None)


def test_writer_lands_in_lga_row_only(population_like_sheet, tmp_path):
    import json
    from update_population_erp_lga import write_lga_erp

    cache = tmp_path / "lga_goondiwindi_population_erp_lga.json"
    cache.write_text(json.dumps({
        "region": "Goondiwindi", "year": 2025, "value": 10438,
        "series_by_year": {"2023": 10417, "2024": 10457, "2025": 10438},
    }))
    results, written, flagged = write_lga_erp(population_like_sheet, [cache])
    assert (written, flagged) == (1, 0)
    assert population_like_sheet.writes == {(4, 5): 10438}   # E4, the LGA row -- not E7
