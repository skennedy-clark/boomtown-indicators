"""
regional-indicators/fetchers/fetch_qrsis_labour.py

Fetches LGA-level and Queensland state-level smoothed unemployment rates
from the QGSO Regional Database (QRSIS), "Labour Force - Small Area"
collection, for the five LGA rows and the "Queensland (benchmark)" row
of the Employment sheet.

SA2 rows are not fetched here; fetch_salm_unemployment.py reproduces
those from DEWR SALM.

Source:
  QRSIS wizard at https://statistics.qgso.qld.gov.au/pls/qis_public/
  collgrp_id=12 ("Labour"), coll_id=1953 ("Labour Force - Small Area
  (Qtr Ended 31 Dec 2010 to Qtr Ended 31 Mar 2026)").

  The wizard is driven in the same way as in fetch_qgso_housing.py and
  fetch_population_erp.py: the same _q() POST encoding, lxml parsing of
  unclosed <OPTION> tags, udqctl_id extraction and six-step flow. The
  code is duplicated rather than shared so that each fetcher stays
  independent of the others (see the fetch_population_erp.py docstring).

Wizard parameters:
  - Series: "Smoothed - Unemployed Persons (Number)", "Smoothed -
    Unemployment Rate (Per cent)", "Smoothed - Labour Force (Number)".
    Only the rate series is fetched.
  - Region types available: GCCSA, LGA, RESREG, S (State), SA2, SA3,
    SA4, SED-2017. Queensland is region "S/3 - Queensland".
  - Time period: period="Quarterly", date_format="Q1" (other collections
    in this pipeline use "Y1"/"Y2"). from_date and to_date are literal
    option strings such as "Qtr Ended 31 Dec 2010" and "Qtr Ended
    31 Mar 2026", the earliest and latest options in the dropdown.
  - The region-type step accepts several types together (its <select>
    is MULTIPLE). LGA and S are selected in one session, as in the
    manually built QGSO extract "QGSO and BoM 2024.xlsx" (sheet Labour),
    which holds LGA, State and SA2 columns from a single query.

Methodology: annual value = mean of the four quarterly smoothed rates in
the calendar year, complete years only. This reproduces the manually
assembled 2024 extract for all five LGAs and Queensland (Queensland's
4/4.1/4.1/4 gives 4.05, the workbook's 2024 value).

Reference values for checking the 2025 column against the reference
workbook: Goondiwindi 3.15, Isaac 1.075, Maranoa 2.225, Toowoomba 3.0,
Western Downs 3.925. There is no reference value for Queensland 2025; an
estimate aggregated from SA2 data (about 4.03) is known not to match the
state series.

Output: cache/unemployment/regions/lga_<slug>.json and
state_queensland_benchmark.json, the directory and schema used by
fetch_salm_unemployment.py (indicators.unemployment = {"label": ...,
"values": {...}}). update_employment.py is driven by section rather
than by source and reads these without change. Note that
fetch_salm_unemployment.py writes lga_<slug>.json files with the same
names when its LGA file is present.
"""

from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR, YEAR_START, YEAR_END

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    raise ImportError("pip install requests beautifulsoup4 lxml")


BASE_URL     = "https://statistics.qgso.qld.gov.au/pls/qis_public/"
PUBLIC_USER  = "edtert"
ACCESS_LEVEL = "85"

COLLGRP_ID = "12"    # Economy > Labour
COLL_ID    = "1953"  # "Labour Force - Small Area (Qtr Ended 31 Dec 2010 to Qtr Ended 31 Mar 2026)"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": "https://statistics.qgso.qld.gov.au/",
    "Origin":  "https://statistics.qgso.qld.gov.au",
}

SERIES_RATE = "Smoothed - Unemployment Rate (Per cent)"
FROM_DATE   = "Qtr Ended 31 Dec 2010"
TO_DATE     = "Qtr Ended 31 Mar 2026"   # update each cycle (fallback only)
                                        # QRSIS is expected to return a
                                        # 500 error for a date beyond the
                                        # latest period, as it does for
                                        # the collections used by
                                        # fetch_qgso_housing.py; this has
                                        # not been tested separately for
                                        # this collection.
PERIOD      = "Quarterly"
DATE_FMT    = "Q1"

# Workbook Employment-sheet row label -> LGA code. These are the codes
# used in the SALM LGA file; QRSIS lists them as "LGA/33610" and so on.
LGA_REGIONS = {
    "Goondiwindi":   "33610",
    "Isaac":         "33980",
    "Maranoa":       "34860",
    "Toowoomba LGA": "36910",
    "Western Downs": "37310",
}
STATE_LABEL = "Queensland (benchmark)"
STATE_CODE  = "3"   # "S/3 - Queensland"

INDICATOR_LABEL = "Smoothed Unemployment rate (%)"


def _q(*pairs) -> list[tuple]:
    result = []
    it = iter(pairs)
    for name in it:
        result.append(("p_names", name))
        result.append(("p_values", next(it)))
    return result


def _parse_options(html: str, select_name: str) -> list[str]:
    """Return the option texts of the named <select>.

    lxml is required: html.parser merges the unclosed <OPTION> tags
    that QRSIS emits.
    """
    soup = BeautifulSoup(html, "lxml")
    for sel in soup.find_all("select", {"name": select_name}):
        return [o.get_text(strip=True) for o in sel.find_all("option") if o.get_text(strip=True)]
    return []


class QRSISLabourFetcher(BaseFetcher):

    SOURCE_NAME      = "qrsis_labour"
    SUPPORTED_STATES = ["QLD"]

    def fetch_all(self):
        session = requests.Session()
        session.headers.update(HEADERS)

        udqctl_id = self._select_collection(session)
        if not udqctl_id:
            self.result.add_error("ALL", f"Failed to get udqctl_id for collection {COLL_ID}")
            return
        self.log.info(f"  udqctl_id={udqctl_id}")

        avail_series = self._get_options(session, "QIS1110W$UDQSER.ProcessSeries",
                                          udqctl_id, "infoser.htm", "p_new_multi")
        self.log.info(f"  Available series: {avail_series}")
        if SERIES_RATE not in avail_series:
            self.result.add_error("ALL", f"Series '{SERIES_RATE}' not found. Available: {avail_series}")
            return
        time_periods_html = self._select_series(session, udqctl_id, [SERIES_RATE])

        self._set_time_period(session, udqctl_id, time_periods_html)

        # The region-type page is a MULTIPLE select, so LGA and State are
        # both chosen in the one session.
        self._select_region_type(session, udqctl_id, "LGA - Local Government Area")
        self._select_region_type(session, udqctl_id, "S - State")

        avail_regions = self._get_options(session, "QIS1110W$UDQREG.ProcessRegions",
                                           udqctl_id, "inforeg.htm", "p_new_multi")
        if avail_regions:
            self.log.info(f"  First available region: {avail_regions[0][:80]}")

        wanted = [f"LGA/{code}" for code in LGA_REGIONS.values()]
        wanted_state = [r for r in avail_regions if r.startswith(f"S/{STATE_CODE} ") or r.startswith(f"S/{STATE_CODE}-")]
        matched = [r for r in avail_regions if any(r.startswith(w) for w in wanted)] + wanted_state
        if not matched:
            self.result.add_error("ALL", "No LGA or State regions matched in the QRSIS region list")
            return
        self._select_regions(session, udqctl_id, matched)
        self.log.info(f"  Regions selected: {len(matched)}")

        html = self._submit_report(session, udqctl_id)
        raw = self._parse_output_html(html)
        self.log.info(f"  Parsed data for {len(raw)} regions")

        annual = self._annualize(raw)

        for label, code in LGA_REGIONS.items():
            self._write_region("LGA", label, annual.get(f"LGA/{code}", {}))
        self._write_region("State", STATE_LABEL, annual.get(f"S/{STATE_CODE}", {}))

    # ── QRSIS wizard steps (same mechanism as fetch_qgso_housing.py) ──────────

    def _select_collection(self, session) -> Optional[str]:
        resp = session.post(
            BASE_URL + "QIS1110W$COLL.ProcessCollection",
            data=_q("usr_id", PUBLIC_USER, "access_lvl", ACCESS_LEVEL,
                    "coll_id", "", "collgrp_id", COLLGRP_ID,
                    "sel_coll_name", COLL_ID, "op_mode", "Next"),
            timeout=30, allow_redirects=True,
        )
        resp.raise_for_status()
        return self._extract_udqctl_id(resp.url, resp.text)

    def _get_options(self, session, proc, udqctl_id, info_page, select_name) -> list[str]:
        resp = session.get(
            BASE_URL + proc,
            params=_q("op_mode", "VIEW", "info_page", info_page,
                      "udqctl_id", udqctl_id, "error_msg", ""),
            timeout=30,
        )
        resp.raise_for_status()
        return _parse_options(resp.text, select_name)

    def _select_series(self, session, udqctl_id, series) -> str:
        for s in series:
            session.post(
                BASE_URL + "QIS1110W$UDQSER.ProcessActions",
                data=_q("udqctl_id", udqctl_id, "info_page", "infoser.htm",
                        "error_msg", "", "op_mode", "->") + [("p_new_multi", s)],
                timeout=30,
            )
            time.sleep(0.2)
        resp = session.post(
            BASE_URL + "QIS1110W$UDQSER.ProcessActions",
            data=_q("udqctl_id", udqctl_id, "info_page", "infoser.htm",
                    "error_msg", "", "op_mode", "Next"),
            timeout=30,
        )
        # The response to this POST is the next page (Time Periods);
        # Oracle PL/SQL WebTK returns each next screen directly. It is
        # returned so that _set_time_period can read the current to_date
        # from the page rather than rely only on TO_DATE.
        return resp.text

    def _discover_max_to_date(self, time_periods_html: str) -> str | None:
        """Return the latest "to date" option on the Time Periods page, or
        None if it cannot be found.

        Duplicated from fetch_qgso_housing.py, which documents the
        rationale, so that the fetchers stay independent of each other.
        """
        soup = BeautifulSoup(time_periods_html, "lxml")
        to_date_marker = soup.find("input", {"name": "p_names", "value": "to_date"})
        if not to_date_marker:
            return None
        select = to_date_marker.find_next("select")
        if not select:
            return None
        options = [o.get_text(strip=True) for o in select.find_all("option") if o.get_text(strip=True)]
        if not options:
            return None
        selected = select.find("option", selected=True)
        if selected and selected.get_text(strip=True):
            return selected.get_text(strip=True)
        return options[0]

    def _set_time_period(self, session, udqctl_id, time_periods_html: str = ""):
        to_date = self._discover_max_to_date(time_periods_html) if time_periods_html else None
        if to_date:
            self.log.info(f"  Latest To Date option on the page: {to_date}")
        else:
            to_date = TO_DATE
            self.log.warning(
                f"  Could not read the To Date option from the page -- "
                f"falling back to the hardcoded {to_date!r}, which may now be stale."
            )
        data = _q("udqctl_id", udqctl_id, "coll_id", COLL_ID, "error_msg", "",
                  "date_format", DATE_FMT, "period", PERIOD,
                  "from_date", FROM_DATE, "to_date", to_date)
        data.append(("p_op_mode", "Next"))
        session.post(BASE_URL + "QIS1110W$UDQCTL.ProcessActions", data=data, timeout=30)

    def _select_region_type(self, session, udqctl_id, geo_label):
        session.post(
            BASE_URL + "QIS1110W$REGTYP.ProcessRegType",
            data=_q("udqctl_id", udqctl_id, "op_mode", "->") + [("p_multi", geo_label)],
            timeout=30,
        )
        time.sleep(0.2)
        session.post(
            BASE_URL + "QIS1110W$REGTYP.ProcessRegType",
            data=_q("udqctl_id", udqctl_id, "op_mode", "Next"),
            timeout=30,
        )

    def _select_regions(self, session, udqctl_id, regions):
        session.post(
            BASE_URL + "QIS1110W$UDQREG.ProcessActions",
            data=_q("udqctl_id", udqctl_id, "info_page", "inforeg.htm",
                    "error_msg", "", "op_mode", "->")
                + [("p_new_multi", r) for r in regions],
            timeout=60,
        )
        time.sleep(0.5)
        session.post(
            BASE_URL + "QIS1110W$UDQREG.ProcessActions",
            data=_q("udqctl_id", udqctl_id, "info_page", "inforeg.htm",
                    "error_msg", "", "op_mode", "Next"),
            timeout=30,
        )

    def _submit_report(self, session, udqctl_id) -> str:
        resp = session.post(
            BASE_URL + "QIS1110W$UDQCTL1.ProcessActions",
            data=_q("udqctl_id", udqctl_id, "coll_id", COLL_ID, "error_msg", "",
                    "ser_sort_col", "Sort Number", "reg_sort_col", "Region Code",
                    "display_style", "For each Region display Time Period by Series",
                    "op_mode", "QRSIS Query"),
            timeout=120,
        )
        resp.raise_for_status()
        return resp.text

    def _extract_udqctl_id(self, url: str, html: str) -> Optional[str]:
        m = re.search(r'p_names=udqctl_id&p_values=(\d+)', url)
        if m:
            return m.group(1)
        m = re.search(r'[?&]udqctl_id=(\d+)', url)
        if m:
            return m.group(1)
        m = re.search(
            r'NAME="p_names"\s+VALUE="udqctl_id"[^>]*>.*?NAME="p_values"\s+VALUE="(\d+)"',
            html, re.IGNORECASE | re.DOTALL,
        )
        if m:
            return m.group(1)
        m = re.search(r'udqctl_id.*?(\d{4,6})', url + html, re.DOTALL)
        return m.group(1) if m else None

    # ── Output parsing ──────────────────────────────────────────────────────

    def _parse_output_html(self, html: str) -> dict:
        """Return { region_code: { period: rate_or_None } }.

        The region-code pattern includes the "S" prefix as well as SA2,
        SA3, SA4 and LGA, so that the state section ("S/3", Queensland)
        is kept. The equivalent pattern in fetch_qgso_housing.py omits
        "S" and would drop state sections.
        """
        result: dict[str, dict] = {}
        sections = re.split(r'Region\s*:\s*', html)
        for section in sections[1:]:
            m = re.match(r'((?:SA2|SA3|SA4|LGA|S)/[\w]+)', section)
            if not m:
                continue
            region_code = m.group(1)

            soup = BeautifulSoup(section, "lxml")
            table = soup.find("table", border=True) or soup.find("table")
            if not table:
                result[region_code] = {}
                continue
            rows = table.find_all("tr")
            if not rows:
                result[region_code] = {}
                continue

            header_cells = rows[0].find_all(["th", "td"])
            series_names = [c.get_text(strip=True) for c in header_cells[1:]]
            rate_col = next((i for i, n in enumerate(series_names) if n == SERIES_RATE), None)

            region_data = {}
            for row in rows[1:]:
                cells = row.find_all(["td", "th"])
                if not cells:
                    continue
                period = cells[0].get_text(strip=True)
                if not period or rate_col is None or rate_col + 1 >= len(cells):
                    continue
                raw = cells[rate_col + 1].get_text(strip=True).replace(",", "")
                try:
                    region_data[period] = float(raw)
                except ValueError:
                    region_data[period] = None
            result[region_code] = region_data
        return result

    def _annualize(self, raw: dict) -> dict:
        """Return { region_code: { year: mean_of_4_quarters } }, complete
        years only.
        """
        import statistics as stats
        out = {}
        for code, periods in raw.items():
            by_year: dict[int, list] = {}
            for period, val in periods.items():
                if val is None:
                    continue
                m = re.search(r'\b(19|20)\d{2}\b', period)
                if not m:
                    continue
                by_year.setdefault(int(m.group(0)), []).append(val)
            out[code] = {
                yr: round(stats.mean(vals), 4)
                for yr, vals in by_year.items()
                if len(vals) == 4 and YEAR_START <= yr <= YEAR_END
            }
        return out

    def _write_region(self, section: str, label: str, annual: dict):
        tag = f"{section}:{label}"
        if not annual:
            self.log.warning(f"  [{tag}] no complete-year data returned")
            self.result.towns_failed.append(tag)
            return

        values = {str(y): v for y, v in sorted(annual.items())}
        out = {
            "region":     label,
            "section":    section,
            "code":       LGA_REGIONS.get(label, STATE_CODE),
            "source":     "QGSO Regional Database (QRSIS) — Labour Force - Small Area",
            "source_url": (
                "https://www.qgso.qld.gov.au/statistics/queensland-regions/"
                "regional-tools-statistics/queensland-regional-database"
            ),
            "note": (
                "Annual value = plain mean of the 4 quarterly smoothed rates "
                "in the calendar year (complete years only). Verified against "
                "the workbook's 2024 column."
            ),
            "indicators": {"unemployment": {"label": INDICATOR_LABEL, "values": values}},
        }

        slug = re.sub(r"[^a-z0-9]+", "_", label.lower()).strip("_")
        out_dir = CACHE_DIR / "unemployment" / "regions"
        out_dir.mkdir(parents=True, exist_ok=True)
        with open(out_dir / f"{section.lower()}_{slug}.json", "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        latest = max(annual)
        self.log.info(f"  [{tag}] {len(values)} years, latest ({latest}) = {annual[latest]:.3f}%")
        self.result.towns_ok.append(tag)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = QRSISLabourFetcher().run()
    sys.exit(0 if result.success else 1)