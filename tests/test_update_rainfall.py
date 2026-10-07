"""
tests/test_update_rainfall.py -- the rainfall writer's block finding on
the Exogenous sheet.

Builds a tiny sheet with the same shape as the real Rainfall section
(town header row, station row, Summer, Winter, Historic Average) and
reads it through tests/fake_xlwings_sheet.py. No network, no real
workbook, no Excel.

The case that matters most: the file the pipeline really runs on (last
year's delivered workbook) labels each station by NAME only
("Harewood"); only the hand-built reference file has the station
number in the label ("Harewood 042078"). Both must work.
"""

import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform" / "xlsx_update"))


def _sheet(tmp_path, rows):
    pytest.importorskip("xlwings")
    import openpyxl
    from fake_xlwings_sheet import FakeSheet

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Exogenous"
    for row in rows:
        ws.append(row)
    path = tmp_path / "exo.xlsx"
    wb.save(path)
    return FakeSheet(path, "Exogenous")


def _rows(chinchilla_label="Harewood", dalby_label="Dalby Airport", third="Winter (Apr-Sept)"):
    return [
        (None, 2023, 2024),
        ("Rainfall",),
        ("Chinchilla ",),                       # trailing space, as in the real sheet
        (chinchilla_label, 441.9, 585.6),
        ("Summer (Jan-Mar, Oct-Dec)", 380.4, 396.6),
        (third, 61.5, 189.0),
        ("Historic Average", 613.9, 613.9),
        ("Dalby",),
        (dalby_label, 439.0, 947.4),
        ("Summer (Jan-Mar, Oct-Dec)", 359.8, 673.4),
        ("Winter (Apr-Sept)", 79.2, 274.0),
        ("Historic Average", 587.6, 587.6),
        (None,),
        ("Education", 2023, 2024, None, "a note far to the right"),
        ("Dalby",),                             # same town name in a LATER section
        ("Full Time Equivalent Enrolments", 3203.6, 3200),
    ]


def _cache(tmp_path, town, station, official_avg=None):
    path = tmp_path / f"{town.lower()}_bom_rainfall.json"
    path.write_text(json.dumps({
        "town": town, "bom_station": station,
        "bom_official_historic_avg_mm": official_avg,
        "bom_official_historic_avg_note": "no page for this station",
        "indicators": {
            "rainfall":        {"2023": 439.0, "2024": 947.4, "2025": 538.4},
            "rainfall_summer": {"2023": 359.8, "2024": 673.4, "2025": 376.8},
            "rainfall_winter": {"2023": 79.2,  "2024": 274.0, "2025": 161.6},
        },
    }))
    return path


def test_finds_block_when_label_has_no_station_number(tmp_path):
    from update_rainfall import _find_rainfall_station_row

    sheet = _sheet(tmp_path, _rows())
    assert _find_rainfall_station_row(sheet, "Chinchilla", "42078") == 4
    assert _find_rainfall_station_row(sheet, "Dalby", "41522") == 9      # not the Education 'Dalby'


def test_finds_block_when_label_has_the_station_number(tmp_path):
    from update_rainfall import _find_rainfall_station_row

    sheet = _sheet(tmp_path, _rows(chinchilla_label="Harewood 042078"))
    assert _find_rainfall_station_row(sheet, "Chinchilla", "42078") == 4


def test_replaced_station_still_writes_and_only_adds_a_note(tmp_path):
    """The label keeps the OLD station's number; towns.toml has the new
    one. The block is found by town, the write goes ahead, and the run
    says the label and towns.toml differ."""
    from update_rainfall import write_rainfall

    sheet = _sheet(tmp_path, _rows(dalby_label="Dalby Airport 041522/ New Site (2026->)"))
    results, written, flagged = write_rainfall(sheet, [_cache(tmp_path, "Dalby", "41599")])
    assert flagged == 0
    assert sheet.writes[(9, 4)] == 538.4
    assert any(line.startswith("Dalby: note:") and "41599" in line for line in results)


def test_matching_station_number_adds_no_note(tmp_path):
    from update_rainfall import write_rainfall

    sheet = _sheet(tmp_path, _rows(dalby_label="Dalby Airport 041522"))
    results, _, _ = write_rainfall(sheet, [_cache(tmp_path, "Dalby", "41522")])
    assert not any("note:" in line for line in results)


def test_town_with_no_block_is_not_an_error(tmp_path):
    from update_rainfall import _find_rainfall_station_row, NoRainfallBlock

    sheet = _sheet(tmp_path, _rows())
    with pytest.raises(NoRainfallBlock):
        _find_rainfall_station_row(sheet, "Shepparton", "81125")


def test_malformed_block_is_refused(tmp_path):
    from update_rainfall import _find_rainfall_station_row

    sheet = _sheet(tmp_path, _rows(third="Something else"))
    with pytest.raises(ValueError, match="isn't laid out"):
        _find_rainfall_station_row(sheet, "Chinchilla", "42078")


def test_write_adds_one_year_column_in_the_right_rows(tmp_path):
    from update_rainfall import write_rainfall

    sheet = _sheet(tmp_path, _rows())
    results, written, flagged = write_rainfall(
        sheet, [_cache(tmp_path, "Dalby", "41522"), _cache(tmp_path, "Shepparton", "81125")]
    )
    assert flagged == 0
    assert sheet.writes == {
        (1, 4): 2025,          # new year header in D1
        (9, 4): 538.4,         # total
        (10, 4): 376.8,        # summer
        (11, 4): 161.6,        # winter
        (12, 4): 587.6,        # historic average carried forward (no official figure)
    }
    assert any("No block" in line and "Shepparton" in line for line in results)


def test_official_average_is_written_under_year_columns_only(tmp_path):
    from update_rainfall import write_rainfall

    sheet = _sheet(tmp_path, _rows())
    write_rainfall(sheet, [_cache(tmp_path, "Dalby", "41522", official_avg=599.1)])
    average_cells = {c: v for (r, c), v in sheet.writes.items() if r == 12}
    # B, C and the new D hold years; E holds a stray note further down the
    # sheet, which stretches used_range but is NOT a year column.
    assert average_cells == {2: 599.1, 3: 599.1, 4: 599.1}
