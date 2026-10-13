"""
tests/test_salm.py

Tests for finding the SA2 CSV download link on the DEWR resource page
(regional-indicators/fetchers/fetch_salm_unemployment.py).
"""

from fetchers.fetch_salm_unemployment import find_sa2_csv_link

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
