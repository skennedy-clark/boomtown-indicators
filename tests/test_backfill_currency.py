"""
tests/test_backfill_currency.py

Regression tests for the history back-fill on rows in a currency or
accounting format. Excel returns those cells as decimal.Decimal through
xlwings on Windows; a check that recognised only int and float took a
full row for an empty one and refilled it from the source, replacing
the workbook's figures (Income and Housing rows of established towns).

tests/fake_xlwings_sheet.py returns Decimal for such cells, as Excel does.
"""

import sys
from decimal import Decimal
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform" / "xlsx_update"))

ACCOUNTING = '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"??_);_(@_)'


@pytest.fixture
def money_sheet(tmp_path):
    pytest.importorskip("xlwings")
    import openpyxl
    from fake_xlwings_sheet import FakeSheet

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Income"
    ws.append([None, None, "2020/21", "2021/22", "2022/23", "2023/24"])
    ws.append(["Chinchilla"])
    ws.append(["Average income", None, 59044, 59550, 63308, 67738])   # established row
    ws.append(["Average income", None, None, None, None, 73581])      # new row: latest year only
    ws.append(["Earners", None, 3306, 3422, 3881, 3946])              # plain number format
    for row in (3, 4):
        for col in range(3, 7):
            ws.cell(row, col).number_format = ACCOUNTING
    path = tmp_path / "income.xlsx"
    wb.save(path)
    return FakeSheet(path, "Income")


SOURCE = {2021: 1.0, 2022: 2.0, 2023: 3.0, 2024: 4.0}


def test_fake_sheet_returns_currency_cells_as_decimal(money_sheet):
    assert isinstance(money_sheet.cells(3, 3).value, Decimal)
    assert isinstance(money_sheet.cells(5, 3).value, float)


def test_established_currency_row_is_not_refilled(money_sheet):
    from base import backfill_new_row
    assert backfill_new_row(money_sheet, 3, 6, 2024, SOURCE) == []
    assert money_sheet.writes == {}


def test_new_currency_row_is_filled(money_sheet):
    from base import backfill_new_row
    assert backfill_new_row(money_sheet, 4, 6, 2024, SOURCE) == [2021, 2022, 2023]
    assert money_sheet.writes == {(4, 3): 1.0, (4, 4): 2.0, (4, 5): 3.0}


def test_text_in_an_earlier_year_counts_as_content(money_sheet):
    from base import backfill_new_row
    money_sheet.writes[(4, 4)] = "n.p."
    assert backfill_new_row(money_sheet, 4, 6, 2024, SOURCE) == []


def test_as_number():
    from base import as_number
    assert as_number(Decimal("63308.25")) == 63308.25
    assert as_number(5) == 5.0 and as_number(2.5) == 2.5
    assert as_number(True) is None and as_number("12") is None and as_number(None) is None
