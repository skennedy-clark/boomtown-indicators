"""
tests/test_crime_header.py

The Crime sheet's Narrabri/NSW block has its own year header row. Adding
a town inserts rows above it, so the writer finds that row on the sheet;
a fixed row number put the new year into a Queensland town's block.
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform" / "xlsx_update"))


@pytest.fixture
def crime_sheet(tmp_path):
    pytest.importorskip("xlwings")
    import openpyxl
    from fake_xlwings_sheet import FakeSheet

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Crime"
    ws.append([None, None, 2023, 2024, 2025])                              # row 1
    ws.append(["QPS Division", "Chart label"])
    ws.append(["Wallumbilla"])
    ws.append(["Good Order Offences", "Good Order Offences", 0, 0, 1.398161])
    ws.append(["Mitchell"])                                                # rows inserted by add_town
    ws.append(["Good Order Offences", "Good Order Offences", 2.1, 3.3, 2.7])
    ws.append([None, None, 2023, 2024])                                    # NSW block's header, moved down
    ws.append(["Narrabri", None, 51.2, 49.8, 50.3])
    path = tmp_path / "crime.xlsx"
    wb.save(path)
    return FakeSheet(path, "Crime")


def test_header_row_is_found_where_it_now_is(crime_sheet):
    from update_crime import _bocsar_header_row
    assert _bocsar_header_row(crime_sheet) == 7


def test_new_year_goes_into_the_nsw_header_only(crime_sheet):
    from update_crime import _mirror_bocsar_year_header
    assert _mirror_bocsar_year_header(crime_sheet, 2025) is True
    assert crime_sheet.writes == {(7, 5): 2025}


def test_nothing_is_written_without_a_header_row(tmp_path):
    pytest.importorskip("xlwings")
    import openpyxl
    from fake_xlwings_sheet import FakeSheet
    from update_crime import _mirror_bocsar_year_header

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Crime"
    ws.append([None, None, 2023, 2024, 2025])
    ws.append(["Wallumbilla", None, 1.0, 2.0, 3.0])
    path = tmp_path / "crime.xlsx"
    wb.save(path)
    sheet = FakeSheet(path, "Crime")
    assert _mirror_bocsar_year_header(sheet, 2025) is False and sheet.writes == {}
