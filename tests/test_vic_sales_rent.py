"""
tests/test_vic_sales_rent.py

Tests for the Victorian sales and rent fetcher
(regional-indicators/fetchers/fetch_vic_sales_rent.py): finding the
files in the data catalogue, reading Valuer-General Victoria's houses
by suburb table and DFFH's quarterly median rents by LGA, and the
calendar-year rent means. No network is needed.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators"))

from fetchers.fetch_vic_sales_rent import (  # noqa: E402
    calendar_year_means, latest_rent_resource, parse_lga_rents, parse_suburb_medians, quarter_of, sales_files,
)

VGV_ROWS = [
    (None,) * 8,
    ("Locality", 2023, None, 2024, None, 2025, None, "Prelim", None, "24-25", "15-25", "PA"),
    (None, None, None, None, None, None, None, 2026, None, "(%)", "(%)", "15-25"),
    (None,) * 8,
    ("SHEPPARTON", "445000", " ", "465000", " ", "495000", " ", "520000", " ", "6", "98", "7.1"),
    ("SHEPPARTON NORTH", "610000", " ", "652500", " ", "649000", " ", "682500", "^", "-1", "86", "6.4"),
    ("YARRAM", "NA", " ", "-", " ", "387500", "^", "430000", "^", "-8", "58", "4.7"),
]


def test_suburb_medians_by_year_without_the_preliminary_year():
    found = parse_suburb_medians(VGV_ROWS, {"SHEPPARTON", "YARRAM"})
    assert set(found) == {"SHEPPARTON", "YARRAM"}            # exact locality, not SHEPPARTON NORTH
    assert found["SHEPPARTON"] == {2023: (445000.0, ""), 2024: (465000.0, ""), 2025: (495000.0, "")}
    assert found["YARRAM"] == {2025: (387500.0, "^")}          # "NA" and "-" are left out


def test_houses_by_suburb_files_in_the_catalogue():
    catalogue = {"resources": [
        {"url": "https://www.land.vic.gov.au/__data/assets/excel_doc/0029/709751/Houses-by-suburb-2013-2023.xlsx"},
        {"url": "https://www.land.vic.gov.au/__data/assets/excel_doc/0033/775617/houses-by-suburb-2015-2025.xlsx"},
        {"url": "https://www.land.vic.gov.au/__data/assets/excel_doc/0034/775618/units-by-suburb-2015-2025.xlsx"},
    ]}
    assert sorted(sales_files(catalogue)) == [2023, 2025]


def test_latest_rent_quarter_chosen_by_title():
    catalogue = {"resources": [
        {"name": "Quarterly: Median Rents by LGA Jun 2025", "url": "a"},
        {"name": " Quarterly: Median Rents by LGA Sep 2025", "url": "b"},
        {"name": "Rental Report - Quarterly: Median Rents by LGA Sept 2022", "url": "c"},
    ]}
    assert latest_rent_resource(catalogue) == "b"
    assert quarter_of("Dec 2024") == (2024, 4)
    assert quarter_of("Count") is None


def test_lga_rents_by_quarter():
    rows = [
        ("Quarterly median rents by Local Government Area", None, None, None, None, None),
        ("3 bedroom house", None, "Jun 2025", "Jun 2025", "Sep 2025", "Sep 2025"),
        (None, None, "Count", "Median", "Count", "Median"),
        ("Hume", "Greater Shepparton", 166, 490, 213, 480),
        (None, "Towong", 7, 500, "-", "-"),
        ("Gippsland", "Wellington", 94, 440, 67, 450),
    ]
    rents = parse_lga_rents(rows, {"greater shepparton", "towong"})
    assert rents["greater shepparton"] == {(2025, 2): 490.0, (2025, 3): 480.0}
    assert rents["towong"] == {(2025, 2): 500.0}               # "-": no figure
    assert "wellington" not in rents


def test_calendar_year_means_complete_years_only():
    quarters = {(2024, 1): 400, (2024, 2): 410, (2024, 3): 420, (2024, 4): 430, (2025, 1): 440}
    assert calendar_year_means(quarters) == {2024: 415.0}
