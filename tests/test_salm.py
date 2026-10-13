"""
tests/test_salm.py

Tests for the SALM fetcher (regional-indicators/fetchers/fetch_salm_unemployment.py):
finding the download links on the DEWR pages, and the regions added from
towns.toml.
"""

from types import SimpleNamespace

from fetchers.fetch_salm_unemployment import (
    LGA_REGIONS, SALMUnemploymentFetcher, find_csv_link, find_resource_page_link, find_sa2_csv_link,
)

SLUG = "salm-smoothed-sa2-datafiles-asgs-2021-march-quarter-2026"


def test_relative_link_is_made_absolute():
    html = f'<a href="/download/17200/{SLUG}/43000/{SLUG}/csv">CSV</a>'
    assert find_sa2_csv_link(html) == f"https://www.dewr.gov.au/download/17200/{SLUG}/43000/{SLUG}/csv"


def test_absolute_link_is_kept():
    url = f"https://www.dewr.gov.au/download/17200/{SLUG}/43000/{SLUG}/csv"
    assert find_sa2_csv_link(f"<a class='x' href='{url}'>CSV</a>") == url


def test_xlsx_and_other_links_are_ignored():
    html = (f'<a href="/download/17200/{SLUG}/43001/{SLUG}/xlsx">XLSX</a>'
            '<a href="/download/16000/salm-smoothed-lga-datafiles/1/x/csv">LGA</a>')
    assert find_sa2_csv_link(html) is None


def test_no_page():
    assert find_sa2_csv_link("") is None


LGA_SLUG = "salm-smoothed-lga-datafiles-asgs-2025-june-quarter-2026"


def test_lga_csv_link():
    html = (f'<a href="/download/17069/{LGA_SLUG}/43877/{LGA_SLUG}/xlsx">XLSX</a>'
            f'<a href="/download/17069/{LGA_SLUG}/43878/{LGA_SLUG}/csv">CSV</a>')
    assert find_csv_link(html, "lga") == f"https://www.dewr.gov.au/download/17069/{LGA_SLUG}/43878/{LGA_SLUG}/csv"
    assert find_csv_link(html, "sa2") is None


def test_resource_page_link_on_main_page():
    html = ('<a href="/employment-research/resources/salm-smoothed-sa2-datafiles-asgs-2021">SA2</a>'
            '<a href="/employment-research/resources/salm-smoothed-lga-datafiles-asgs-2025">LGA</a>')
    assert find_resource_page_link(html, "lga") == (
        "https://www.dewr.gov.au/employment-research/resources/salm-smoothed-lga-datafiles-asgs-2025")


def _fetcher(*towns):
    fetcher = SALMUnemploymentFetcher.__new__(SALMUnemploymentFetcher)
    fetcher.config = SimpleNamespace(study_towns=lambda: list(towns))
    return fetcher


def _town(sa2_code, sa2_name, lga, qgso_lga):
    return SimpleNamespace(sa2_code=sa2_code, sa2_name=sa2_name, lga=lga, qgso_lga=qgso_lga)


def test_new_town_regions_are_added_once():
    emerald = _town("308011192", "Emerald", "Central Highlands (Qld)", "LGA/32270")
    mitchell = _town("307011177", "Roma Surrounds", "Maranoa", "LGA/34860")   # both already covered
    fetcher = _fetcher(emerald, mitchell)
    sa2 = fetcher.sa2_regions()
    lga = fetcher.lga_regions()
    assert sa2["Emerald"] == "308011192"
    assert list(sa2.values()).count("307011177") == 1
    assert lga["Central Highlands"] == "32270"
    assert len(lga) == len(LGA_REGIONS) + 1


def test_lga_name_with_state_suffix_matches():
    norm = SALMUnemploymentFetcher._norm_name
    assert norm("Central Highlands (Qld)") == norm("Central Highlands")
    assert norm("Toowoomba (R)") == norm("Toowoomba LGA")
