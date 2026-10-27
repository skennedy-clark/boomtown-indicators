"""
tests/test_qgso_housing_approvals.py

Tests for the building approvals series of the QRSIS housing fetcher
(regional-indicators/fetchers/fetch_qgso_housing.py): new houses and
non-residential dwelling units, summed by year for the Housing sheet's
regions. No network is needed.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators"))

from fetchers.fetch_qgso_housing import (  # noqa: E402
    REGION_APPROVAL_SERIES, REGION_SERIES_LABELS, SERIES_APPROVALS, SERIES_NONRES,
    QGSOHousingFetcher, add_monthly, approval_totals,
)

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def months(year, values, start=1):
    return {f"{MONTHS[start - 1 + i]} {year}": v for i, v in enumerate(values)}


def test_both_series_are_fetched_for_the_regions_with_the_sheet_labels():
    assert REGION_APPROVAL_SERIES == [SERIES_APPROVALS, SERIES_NONRES]
    assert REGION_SERIES_LABELS[SERIES_NONRES] == \
        "Building Approvals: Non-residential dwelling units (Private) Total (Number)"


def test_annual_totals_need_twelve_months_except_the_latest_year():
    by_year = {
        2001: months(2001, [1] * 6, start=7),     # the collection starts in July 2001
        2002: months(2002, [0] * 11 + [2]),
        2003: months(2003, [1] * 12),
        2004: months(2004, [3, 1]),               # latest year, in progress
    }
    assert approval_totals(by_year) == {"2002": 2, "2003": 12, "2004": 4}
    assert approval_totals({}) == {}


def test_monthly_reports_are_added_by_region_series_and_year():
    raw = {
        "SA2/308011192 Emerald": {
            "Jan 2019": {SERIES_APPROVALS: 5, SERIES_NONRES: 0},
            "Feb 2019": {SERIES_APPROVALS: 3, SERIES_NONRES: 1},
        },
        "SA2/999999999 Elsewhere": {"Jan 2019": {SERIES_APPROVALS: 9}},
    }
    by_region = {}
    add_monthly(by_region, raw, {"SA2/308011192": ("Emerald", "SA2")}, REGION_APPROVAL_SERIES,
                QGSOHousingFetcher._parse_month_year)
    assert set(by_region) == {"SA2/308011192"}
    assert by_region["SA2/308011192"][SERIES_NONRES][2019] == {"Jan 2019": 0, "Feb 2019": 1}
    assert sum(by_region["SA2/308011192"][SERIES_APPROVALS][2019].values()) == 8
