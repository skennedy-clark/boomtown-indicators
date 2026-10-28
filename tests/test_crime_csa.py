"""
tests/test_crime_csa.py

Tests for the Victorian crime fetcher
(regional-indicators/fetchers/fetch_crime_csa.py): grouping the CSA
offence subdivisions into the Crime sheet's New South Wales rows,
choosing the December workbooks that reach furthest back, and reading
Table 02. No network is needed.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators"))

from fetchers.fetch_crime_csa import CATEGORIES, candidates, category_of, december_files, parse_rates  # noqa: E402


def test_every_subdivision_has_one_row():
    assert category_of("A20 Assault and related offences") == "Assault"
    assert category_of("A50 Robbery") == "Robbery"
    assert category_of("A70 Stalking, harassment and threatening behaviour") == "Other offences against the person"
    assert category_of("B30 Burglary/Break and enter") == "Theft"
    assert category_of("B20 Property damage") == "Malicious damage to property"
    assert category_of("C30 Drug use and possession") == "Drug offences"
    assert category_of("F20 Transport regulation offences") == "Transport regulatory offences"
    assert category_of("F10 Regulatory driving offences") == "Other offences"
    assert category_of("E20 Breaches of orders") == "Other offences"
    assert set(CATEGORIES) == {"Assault", "Robbery", "Other offences against the person", "Theft",
                               "Malicious damage to property", "Drug offences",
                               "Transport regulatory offences", "Other offences"}


def test_december_workbooks_found_in_the_catalogue():
    catalogue = {"resources": [
        {"url": "https://files.crimestatistics.vic.gov.au/2026-03/Data_Tables_LGA_Criminal_Incidents_Year_Ending_December_2025.xlsx"},
        {"url": "https://files.crimestatistics.vic.gov.au/2026-06/Data_Tables_LGA_Criminal_Incidents_Year_Ending_March_2026.xlsx"},
        {"url": "https://files.crimestatistics.vic.gov.au/2022-03/Data_Tables_LGA_Criminal_Incidents_Year_Ending_December_2021.xlsx"},
    ]}
    assert sorted(december_files(catalogue)) == [2021, 2025]


def test_older_workbooks_are_tried_reaching_furthest_back_first():
    available = {2025: "a", 2024: "b", 2022: "c", 2021: "d", 2019: "e"}
    # 2025 covers 2016-2025: the 2019 workbook reaches furthest back, then 2021.
    assert candidates(available, 2016, {2025}) == [2019, 2021, 2022, 2024]
    assert candidates(available, 2010, {2025, 2019}) == []


def test_rates_added_by_category_and_converted_to_per_thousand():
    rows = [
        ("Year", "Year ending", "Police Service Area", "Local Government Area", "Offence Division",
         "Offence Subdivision", "Offence Subgroup", "Incidents Recorded",
         "PSA Rate per 100,000 population", "LGA Rate per 100,000 population"),
        (2025, "December", "Shepparton", "Greater Shepparton", "A", "A20 Assault and related offences", "A211", 10, 0, 500.0),
        (2025, "December", "Shepparton", "Greater Shepparton", "A", "A20 Assault and related offences", "A231", 10, 0, 250.0),
        (2025, "December", "Shepparton", "Greater Shepparton", "C", "C30 Drug use and possession", "C32", 5, 0, 100.0),
        (2025, "December", "Ballarat", "Ballarat", "A", "A20 Assault and related offences", "A211", 99, 0, 999.0),
    ]
    rates = parse_rates(rows, {"Greater Shepparton"})
    assert set(rates) == {"Greater Shepparton"}
    assert rates["Greater Shepparton"][2025]["Assault"] == 7.5
    assert rates["Greater Shepparton"][2025]["Drug offences"] == 1.0
