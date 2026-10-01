"""
fetchers/fetch_qrsis_labour.py
--------------------------------
Fetches LGA-level and Queensland state-level smoothed unemployment rates from
the QGSO Regional Database (QRSIS) -- the "Labour Force - Small Area"
collection. Fills the two gaps fetch_salm_unemployment.py cannot: the 5 LGA
rows and the "Queensland (benchmark)" row on the Employment sheet.
(SA2 rows are NOT duplicated here -- fetch_salm_unemployment.py's DEWR SALM
source already reproduces those exactly; see its docstring.)

COLLECTION IDENTIFIED 2026-09-29 from a real walkthrough of the QRSIS wizard
that Steve stepped through and captured in full (collgrp_id=12 "Labour",
coll_id=1953 "Labour Force - Small Area (Qtr Ended 31 Dec 2010 to Qtr Ended
31 Mar 2026)"). This is the SAME live HTML technique that found ERP's
collgrp_id=1/coll_id=1961 (see fetch_population_erp.py) and that is proven
working end-to-end in fetch_qgso_housing.py (sales/rent, live-confirmed).
Reuses that exact mechanism (same _q() POST encoding, same lxml unclosed-
<OPTION> parsing, same udqctl_id extraction, same 6-step wizard flow) --
deliberately duplicated rather than shared, per this project's established
practice of not refactoring a proven fetcher to share code with an unproven
one (see fetch_population_erp.py's docstring for the reasoning).

CONFIRMED REAL FROM THE WALKTHROUGH (not guessed):
  - Series names: "Smoothed - Unemployed Persons (Number)",
    "Smoothed - Unemployment Rate (Per cent)", "Smoothed - Labour Force
    (Number)". Only the rate series is fetched here.
  - Region types available: GCCSA, LGA, RESREG, S (State), SA2, SA3, SA4,
    SED-2017. Queensland is region "S/3 - Queensland".
  - Time period: period="Quarterly", date_format="Q1" (the hidden field's
    real value -- confirmed by the walkthrough, not "Y1"/"Y2" as other
    collections in this project use). from_date/to_date are literal
    "Qtr Ended 31 Mar 2026" / "Qtr Ended 31 Dec 2010" strings -- the
    dropdown's own earliest and latest options.
  - The region-type step accepts multiple types selected together (its
    <select> has MULTIPLE) -- this fetcher selects LGA and S in the one
    session, matching how last year's manually-built QGSO extract
    ("QGSO and BoM 2024.xlsx", sheet Labour) had LGA + State + SA2 columns
    side by side in a single query.

METHODOLOGY, already verified against the workbook (see TODO.md "Exemplar
archaeology"): annual value = plain mean of the 4 quarterly smoothed rates
in the calendar year (complete years only). VERIFIED EXACTLY against last
year's manually-assembled 2024 extract for all 5 LGAs and Queensland
(Queensland's 4/4.1/4.1/4 -> 4.05, matching the workbook's 2024 exactly).

*** NOT YET TESTED against the live QRSIS endpoint. *** collgrp_id/coll_id
and the series/region-type/date-format strings are all confirmed real
(read directly from Steve's captured wizard HTML), but this exact
Python flow has not itself been run live yet -- test on the real workbook's
2025 column against the known values before trusting it (Goondiwindi 3.15,
Isaac 1.075, Maranoa 2.225, Toowoomba 3.0, Western Downs 3.925, Queensland
TBD -- see TODO.md for why the SA2-aggregate estimate of ~4.03 is known
to be wrong and this collection is expected to fix it).

OUTPUT: writes into the SAME cache/unemployment/regions/ directory
fetch_salm_unemployment.py uses (lga_*.json, state_queensland.json), in the
SAME schema (indicators.unemployment = {"label":..., "values": {...}}) so
update_employment.py needs no changes to pick these up -- it is already
section-driven, not source-driven.
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
TO_DATE     = "Qtr Ended 31 Mar 2026"   # update each cycle -- likely the same
                                        # 500-error-on-future-date behaviour
                                        # confirmed for fetch_qgso_housing.py's
                                        # collections (same QRSIS wizard
                                        # mechanism), not independently
                                        # re-tested for this collection
PERIOD      = "Quarterly"
DATE_FMT    = "Q1"

# Workbook Employment-sheet row label -> QRSIS LGA code (confirmed against
# the March-2026 SALM LGA file, same codes; QRSIS lists these as "LGA/33610"
# etc.)
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
    """MUST use lxml -- html.parser merges QRSIS's unclosed <OPTION> tags."""
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

        # Multiple region types in one session (LGA + State), per the real
        # wizard's MULTIPLE-select region-type page.
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
        # ADDED 2026-09-29, same port as fetch_qgso_housing.py: this response
        # IS the next page (Time Periods) -- Oracle PL/SQL WebTK returns each
        # next screen directly from the POST. Returned so _set_time_period
        # can discover the real current to_date instead of relying only on
        # the hardcoded TO_DATE.
        return resp.text

    def _discover_max_to_date(self, time_periods_html: str) -> str | None:
        """See fetch_qgso_housing.py's identical method for the full
        rationale -- ported here unchanged (duplicated deliberately, per
        this project's practice of not coupling fetchers to each other)."""
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
            self.log.info(f"  Discovered real max To Date from the page: {to_date}")
        else:
            to_date = TO_DATE
            self.log.warning(
                f"  Could not discover the real To Date option from the page -- "
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
        """
        Returns { region_code: { period: rate_or_None } }.
        FIXED vs fetch_qgso_housing.py's version: that one's region-code
        regex is (?:SA2|SA3|SA4|LGA)/[\\w]+ and silently drops "S/3" (state)
        sections entirely -- confirmed by inspection, not by a live failure,
        since this fetcher was never run against LGA/State data before.
        Added "S" to the prefix alternation so the Queensland row is kept.
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
        """{ region_code: { year: mean_of_4_quarters } }, complete years only."""
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