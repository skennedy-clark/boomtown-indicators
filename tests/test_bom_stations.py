"""
tests/test_bom_stations.py

Station names from BOM's stations list (regional-indicators/fetchers/
bom_stations.py), and their use for the rainfall station label by
transform/add_town.py and transform/xlsx_update/update_rainfall.py.
"""

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform"))
sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform" / "xlsx_update"))

from fetchers.bom_stations import parse, station_label, tidy_name  # noqa: E402

LIST = """Bureau of Meteorology product IDCJMC0014.                                       Produced: 19 Oct 2026

   Site  Dist  Site name                                 Start     End      Lat       Lon Source         STA Height (m)   Bar_ht    WMO
------- ----- ---------------------------------------- ------- ------- -------- --------- -------------- --- ---------- -------- ------
 043020 43    MITCHELL POST OFFICE                        1884      .. -26.4888  147.9777 GPS            QLD      336.5    338.0  94514
 044021 44    CHARLEVILLE AERO                            1942      .. -26.4139  146.2558 GPS            QLD      301.6    303.3  94510
 054038 54    NARRABRI AIRPORT AWS                        1999      .. -30.3154  149.8302 GPS            NSW      218.0    218.6  95504
 043091 43    ROMA AIRPORT                                1985      .. -26.5477  148.7710 Unknown        QLD      307.4    307.8  94515

19511 stations
"""
NAMES = parse(LIST)


def test_stations_list_is_parsed():
    assert NAMES == {43020: "MITCHELL POST OFFICE", 44021: "CHARLEVILLE AERO",
                     54038: "NARRABRI AIRPORT AWS", 43091: "ROMA AIRPORT"}


def test_labels_match_the_workbook_style():
    assert station_label("43020", NAMES) == "Mitchell Post Office 043020"
    assert station_label(54038, NAMES) == "Narrabri Airport AWS 054038"
    assert station_label("99999", NAMES) is None
    assert tidy_name("O'BRIEN CREEK") == "O'Brien Creek"


@pytest.fixture
def exogenous(tmp_path):
    pytest.importorskip("xlwings")
    import openpyxl
    from fake_xlwings_sheet import FakeSheet

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Exogenous"
    for row in [(None, 2024, 2025), ("Rainfall",),
                ("Charleville",), ("Charleville Airport 043091 ", 402.8, 380.1),
                ("Summer (Jan-Mar, Oct-Dec)", 300, 280), ("Winter (Apr-Sept)", 100, 100), ("Historic Average", 490, 490),
                ("Roma",), ("Roma Airport 043091 ", 366.6, 401.2),
                ("Summer (Jan-Mar, Oct-Dec)", 250, 300), ("Winter (Apr-Sept)", 116, 101), ("Historic Average", 580, 580),
                ("Education",), ("Charleville",), ("Full Time Equivalent Enrolments", 300, 310)]:
        ws.append(row)
    path = tmp_path / "exo.xlsx"
    wb.save(path)
    return FakeSheet(path, "Exogenous")


def test_add_town_labels_the_new_towns_station(exogenous):
    from add_town import label_rainfall_station
    town = SimpleNamespace(name="Charleville", bom_station="44021")
    report = label_rainfall_station(exogenous, town, NAMES)
    assert exogenous.writes == {(4, 1): "Charleville Aero 044021"}
    assert "Charleville Aero 044021" in report[0]


def test_a_label_with_the_right_number_is_kept(exogenous):
    from add_town import label_rainfall_station
    assert label_rainfall_station(exogenous, SimpleNamespace(name="Roma", bom_station="43091"), NAMES) == []
    assert exogenous.writes == {}


def test_station_missing_from_the_list_is_reported_not_written(exogenous):
    from add_town import label_rainfall_station
    report = label_rainfall_station(exogenous, SimpleNamespace(name="Charleville", bom_station="12345"), NAMES)
    assert exogenous.writes == {} and "not in BOM's stations list" in report[0]


def test_rainfall_note_suggests_the_label(exogenous, monkeypatch):
    import fetchers.bom_stations as bom
    from update_rainfall import _station_label_note
    monkeypatch.setattr(bom, "station_names", lambda *a, **k: NAMES)
    note = _station_label_note(exogenous, 4, "44021")
    assert "'Charleville Aero 044021'" in note
