"""
tests/test_nsw_approvals_csv.py

The New South Wales building approvals fetcher
(regional-indicators/fetchers/fetch_nsw_approvals.py) reads each month,
zeros included, from the ABS Data API's CSV response. The SDMX-JSON
response left out zero observations, which dropped whole years for
small LGAs. No network is needed.
"""

import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators"))

import fetchers.fetch_nsw_approvals as nsw  # noqa: E402

CSV = (
    "DATAFLOW,MEASURE,SECTOR,WORK_TYPE,BUILDING_TYPE,REGION_TYPE,REGION,FREQ,TIME_PERIOD,OBS_VALUE,UNIT_MEASURE\n"
    "ABS:BA_LGA2018(1.0.0),1,1,1,110,LGA2018,15650,M,2018-07,0,NUM\n"
    "ABS:BA_LGA2018(1.0.0),1,1,1,100,LGA2018,15650,M,2018-07,0,NUM\n"
    "ABS:BA_LGA2018(1.0.0),1,1,1,110,LGA2018,15650,M,2018-08,1,NUM\n"
    "ABS:BA_LGA2018(1.0.0),1,1,1,100,LGA2018,15650,M,2018-08,3,NUM\n"
)


class Response:
    text = CSV

    def raise_for_status(self):
        pass


def test_zero_months_are_kept(monkeypatch):
    monkeypatch.setattr(nsw.requests, "get", lambda *a, **k: Response())
    fetcher = nsw.NSWApprovalsFetcher.__new__(nsw.NSWApprovalsFetcher)
    fetcher.log = logging.getLogger("test")
    months = fetcher._months("BA_LGA2018", "LGA2018", "15650")
    assert months["2018-07"] == {"new_houses": 0, "new_residential_building": 0}
    assert months["2018-08"] == {"new_houses": 1, "new_residential_building": 3}
