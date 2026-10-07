"""
regional-indicators/fetchers/fetch_nsw_labour.py

Fetches the NSW state-level unemployment rate for the State section of
the Employment sheet.

Source: ABS Labour Force, Australia, Table 002 "Labour force status by
Sex, New South Wales" (62020002.xlsx). NSW has no equivalent of SALM or
QRSIS (fetch_qrsis_labour.py covers Queensland only); the ABS Labour
Force Survey is the source that reproduces the workbook's history.

Methodology: annual value = mean of the 12 monthly "Unemployment rate ;
Persons ;" values of the Original series (not Trend, not Seasonally
Adjusted), complete years only. Values drift slightly between releases
because ABS rebenchmarks the series quarterly: 2024 computes to 3.8988
from the August 2026 release against 3.9039 from an earlier one. 2025
computes to 4.0950757583; the reference workbook's hand-entered 3.9 for
NSW 2025 does not correspond to any release.

URL discovery: the download URL ("<release-slug>/62020002.xlsx") is
scraped from the latest-release page rather than built from a month and
year. The page links the file with a relative href
("/statistics/labour/.../aug-2026/62020002.xlsx", no domain), so the
pattern matches the relative form. If the scrape fails, FALLBACK_URL is
used.

File layout: sheet "Data1", with a "Series Type" metadata row (row 3)
that distinguishes the Trend, Seasonally Adjusted and Original columns.
The layout is unchanged by the ABS Labour Force Modernisation program
(completed with the August 2026 release).

Output: cache/unemployment/regions/state_nsw.json, in the same schema as
the State region files of fetch_qrsis_labour.py (section="State",
indicators.unemployment = {label, values}). update_employment.py is
driven by section rather than by source and reads it without change.
"""

from __future__ import annotations

import json
import re
import sys
import datetime as dt
import statistics
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR, YEAR_START, YEAR_END

try:
    import requests
    import openpyxl
except ImportError:
    raise ImportError("pip install requests openpyxl")


LATEST_RELEASE_PAGE = (
    "https://www.abs.gov.au/statistics/labour/employment-and-unemployment/"
    "labour-force-australia/latest-release"
)
# Update each cycle if the scrape fails. ABS republishes this table under
# a new release-slug folder (for example "aug-2026") every month; this
# URL is only the fallback.
FALLBACK_URL = (
    "https://www.abs.gov.au/statistics/labour/employment-and-unemployment/"
    "labour-force-australia/aug-2026/62020002.xlsx"
)

CACHE_KEY = "abs_nsw_labour_table002"
REGION_LABEL = "NSW"
INDICATOR_LABEL = "Smoothed Unemployment rate (%)"  # the label used for every other Employment row


def _discover_current_url(log) -> str:
    """Return the current download URL of Table 002, scraped from the
    latest-release page, or FALLBACK_URL if the scrape fails. The page
    uses a relative href, which the pattern matches.
    """
    try:
        resp = requests.get(LATEST_RELEASE_PAGE, timeout=30)
        resp.raise_for_status()
        m = re.search(
            r'/statistics/labour/employment-and-unemployment/'
            r'labour-force-australia/([a-z]+-\d{4})/62020002\.xlsx',
            resp.text,
        )
        if m:
            url = f"https://www.abs.gov.au{m.group(0)}"
            log.info(f"  Found current NSW labour force table on the page: {url}")
            return url
    except requests.RequestException as exc:
        log.warning(f"  Could not scrape the latest-release page: {exc}")
    log.warning(f"  Falling back to the hardcoded URL, which may now be stale: {FALLBACK_URL}")
    return FALLBACK_URL


class NSWLabourFetcher(BaseFetcher):

    SOURCE_NAME      = "nsw_labour"
    SUPPORTED_STATES = ["NSW"]

    def fetch_all(self):
        url = _discover_current_url(self.log)
        path = self.download(url, CACHE_KEY, suffix=".xlsx")
        if not path:
            self.result.add_error("ALL", "Could not download ABS Table 002 (NSW labour force)")
            return

        self.log.info(f"  Parsing {path.name} ({path.stat().st_size // 1024} KB)")
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            ws = wb["Data1"]
            rows = list(ws.iter_rows(values_only=True))
        except Exception as exc:
            self.log.error(f"Parse error: {exc}", exc_info=True)
            self.result.add_error("ALL", f"Parse error: {exc}")
            return

        header = rows[0]
        series_type_row = next((r for r in rows[:12] if r and r[0] == "Series Type"), None)
        if series_type_row is None:
            self.result.add_error("ALL", "Could not find the 'Series Type' metadata row -- table layout may have changed")
            return

        original_col = None
        for i, h in enumerate(header):
            if (h and str(h).startswith("Unemployment rate") and "Persons" in str(h)
                    and ">" not in str(h) and str(series_type_row[i]) == "Original"):
                original_col = i
                break
        if original_col is None:
            self.result.add_error("ALL", "Could not find the Original-series unemployment rate column -- table layout may have changed")
            return

        monthly: dict[int, dict[int, float]] = {}
        for r in rows:
            if isinstance(r[0], (dt.datetime, dt.date)) and isinstance(r[original_col], (int, float)):
                monthly.setdefault(r[0].year, {})[r[0].month] = r[original_col]

        values = {}
        for yr, months in monthly.items():
            if len(months) == 12 and YEAR_START <= yr <= YEAR_END:
                values[str(yr)] = round(statistics.mean(months.values()), 4)

        if not values:
            self.result.add_error("ALL", "No complete years found in the parsed data")
            return

        out = {
            "region":     REGION_LABEL,
            "section":    "State",
            "source":     "ABS Labour Force, Australia — Table 002, New South Wales (Original series)",
            "source_url": url,
            "note": (
                "Annual value = plain mean of the 12 monthly Original-series "
                "unemployment rate readings (complete years only), not the "
                "Trend or Seasonally Adjusted series. NSW has no SALM/QRSIS "
                "equivalent."
            ),
            "indicators": {"unemployment": {"label": INDICATOR_LABEL, "values": values}},
        }

        out_dir = CACHE_DIR / "unemployment" / "regions"
        out_dir.mkdir(parents=True, exist_ok=True)
        with open(out_dir / "state_nsw.json", "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        latest = max(values, key=int)
        self.log.info(f"  [State:{REGION_LABEL}] {len(values)} years, latest ({latest}) = {values[latest]:.4f}%")
        self.result.towns_ok.append(f"State:{REGION_LABEL}")


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = NSWLabourFetcher().run()
    sys.exit(0 if result.success else 1)
