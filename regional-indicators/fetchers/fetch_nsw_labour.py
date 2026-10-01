"""
fetchers/fetch_nsw_labour.py
------------------------------
Fetches the NSW state-level unemployment rate for the Employment sheet's
State section, completing the gap left after fetch_qrsis_labour.py (which
only covers Queensland). Confirmed there's no equivalent QLD-style source
for NSW -- ABS's Labour Force Survey (not SALM/QRSIS) is the real source,
confirmed by exact match against the workbook's own history.

METHODOLOGY (re-confirmed live 2026-09-30, not assumed from the earlier
investigation -- ABS's "Labour Force Modernisation program" completed with
the August 2026 release just 6 days before this was built, exactly the
kind of change worth re-checking rather than trusting a weeks-old note):
NSW's annual value = mean of the 12 monthly "Unemployment rate ; Persons ;"
ORIGINAL series values (not Trend, not Seasonally Adjusted) from ABS Table
002 "Labour force status by Sex, New South Wales". Re-verified against the
live file: 2024 = 3.8988 (close to, not identical to, the 3.9039 value
confirmed exactly weeks ago -- the small drift is ABS's own "Quarterly
rebenchmarking of labour force statistics" note on the current release
page, expected, not a bug); 2025 = 4.0950757583, consistent with the
earlier finding that the exemplar's hand-typed "3.9" for NSW 2025 doesn't
match any real source vintage.

URL DISCOVERY: the direct URL ("<release-slug>/62020002.xlsx") is scraped
from the "latest-release" page rather than guessed from a constructed
month/year slug -- confirmed the page's own raw HTML link is a RELATIVE
href ("/statistics/labour/.../aug-2026/62020002.xlsx"), no domain prefix;
regex must match on that, not an absolute URL (same class of bug found
and fixed in fetch_population_nrw.py's and fetch_population_ucl.py's
scrape functions on 2026-09-29 -- this one was written with that lesson
already applied, not discovered again the hard way). Falls back to a
hardcoded last-known-good URL if the scrape fails, same pattern as this
project's other self-updating fetchers.

Table structure confirmed unchanged since the Modernisation program (same
"Data1" sheet, same "Series Type" metadata row at row 3, same three
columns Trend/Seasonally Adjusted/Original) -- checked directly against
the live August 2026 file, not assumed just because the URL still worked.

Output schema matches fetch_qrsis_labour.py's State region files exactly
(section="State", indicators.unemployment = {label, values}), so
update_employment.py picks this up with NO code changes -- it is already
section-driven, not source-driven, confirmed by inspection of that file.
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
# update each cycle if the live scrape below ever fails -- ABS republishes
# this table under a new release-slug folder (e.g. "aug-2026") every month;
# this hardcoded URL is only the fallback.
FALLBACK_URL = (
    "https://www.abs.gov.au/statistics/labour/employment-and-unemployment/"
    "labour-force-australia/aug-2026/62020002.xlsx"
)

CACHE_KEY = "abs_nsw_labour_table002"
REGION_LABEL = "NSW"
INDICATOR_LABEL = "Smoothed Unemployment rate (%)"  # matches the label already used for every other Employment row


def _discover_current_url(log) -> str:
    """Scrape the live 'latest-release' page for Table 002's real current
    download link, rather than guessing a release slug. See module
    docstring for the relative-URL regex lesson this already incorporates.
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
                "unemployment rate readings (complete years only) -- NOT "
                "Trend or Seasonally Adjusted. Confirmed exact match against "
                "the workbook's own history for multiple years; NSW has no "
                "SALM/QRSIS equivalent, ABS Labour Force is the real source."
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
