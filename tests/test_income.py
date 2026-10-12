"""
tests/test_income.py

Tests for the Income sheet writer
(regional-indicators/transform/xlsx_update/update_income.py), run against
a small workbook through tests/fake_xlwings_sheet.py; no Excel is needed.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform" / "xlsx_update"))


def test_income_new_year_column_gets_its_calendar_year(tmp_path):
    """The Income sheet's row 2 carries the calendar year that the charts'
    axis reads; a new year column must receive it."""
    import openpyxl
    from fake_xlwings_sheet import FakeSheet
    from update_income import _find_year_column

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Income"
    ws.append([None, None, "2021/22", "2022/23"])
    ws.append(["Postcode", "Label", 2022, 2023])
    path = tmp_path / "income.xlsx"
    wb.save(path)

    sheet = FakeSheet(path, "Income")
    col = _find_year_column(sheet, 2024)
    assert col == 5
    assert sheet.writes[(1, 5)] == "2023/24" and sheet.writes[(2, 5)] == 2024
