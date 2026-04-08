"""
fetchers/fetch_qgso_housing.py
-------------------------------
Fetches housing indicators from the QGSO Regional Database (QRSIS).
QLD only — all three collections are fetched in one session.

COLLECTIONS FETCHED:
  1925  Residential land and dwelling sales   (Sep 2000 – Sep 2025, quarterly)
  1929  Median rent                            (Dec 1989 – Mar 2026, quarterly)
  2075  Building Approvals (Historical)        (Jul 2001 – Dec 2018, monthly)
  2031  Building Approvals (Current)           (Jan 2019 – present, monthly)

INDICATORS PRODUCED (per town, annual):
  housing_sales_count    Detached dwelling: number of sales
                         Annual = Dec-quarter rolling total (already 12-month sum)
  housing_median_price   Detached dwelling: median sale price ($)
                         Annual = mean of 4 quarterly medians
  rent_3bed_median       House - 3 bedroom - median rent of lodgements ($/week)
                         Annual = mean of 4 quarterly medians
  building_approvals     Residential dwelling units (Private); New Houses
                         Annual = sum of 12 monthly values

HOW THE SCRAPER WORKS:
  QRSIS is an Oracle PL/SQL Web Toolkit multi-step wizard.
  The udqctl_id is a server-side query handle assigned per session — it
  changes every run and cannot be hardcoded.

  Step flow per collection:
    1. POST ProcessCollection   → assigns udqctl_id
    2. GET  ProcessSeries (VIEW) → list available series
    3. POST ProcessActions (->)  → move desired series to "Selected"
    4. POST ProcessActions (Next)
    5. POST ProcessActions (time) → set from/to date range (Next)
    6. POST ProcessRegType (->)  → select SA2 region type
    7. POST ProcessRegType (Next)
    8. GET  ProcessRegions (VIEW) → list available regions
    9. POST ProcessActions (->)  → move desired SA2 regions to "Selected"
   10. POST ProcessActions (Next)
   11. POST ProcessActions (QRSIS Query) → generate output HTML
   12. Parse HTML table directly (no XLS download needed)

  Region values are the full option text, e.g.:
    "SA2/307021183 - Wambo (01/07/2011 - )"

  Series values are the full option text, e.g.:
    "Detached dwelling: number of sales (Number)"

SOURCE:
  https://www.qgso.qld.gov.au/statistics/queensland-regions/regional-tools-statistics/queensland-regional-database

Website CSVs produced:
  House price.csv
  House sales.csv
  Rent.csv
  Building approvals.csv
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
from config import CACHE_DIR

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    raise ImportError("pip install requests beautifulsoup4")


# ── QRSIS constants ─────────────────────────────────────────────────────────────

BASE_URL     = "https://statistics.qgso.qld.gov.au/pls/qis_public/"
PUBLIC_USER  = "edtert"
ACCESS_LEVEL = "85"
COLLGRP_ID   = "22"   # Housing group

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": "https://statistics.qgso.qld.gov.au/",
    "Origin":  "https://statistics.qgso.qld.gov.au",
}

# Rent series — try each in order until one matches
RENT_SERIES_CANDIDATES = [
    "House - 3 bedroom - median rent of lodgements ($/week)",
    "3 bedroom house - median rent of lodgements ($/week)",
    "3 Bedroom House: Median rent of lodgements ($/week)",
]

# ── Collection definitions ───────────────────────────────────────────────────────
# Update to_date each year after QGSO releases new data.

COLLECTIONS = {
    "sales": {
        "id":        "1925",
        "series":    [
            "Detached dwelling: number of sales (Number)",
            "Detached dwelling: median sale price ($)",
        ],
        "from_date": "Year Ended 30 Sep 2000",
        "to_date":   "Year Ended 30 Sep 2025",
        "period":    "Quarterly",
    },
    "rent": {
        "id":        "1929",
        "series":    RENT_SERIES_CANDIDATES[:1],   # first candidate; fallback in code
        "from_date": "Year Ended 31 Dec 2000",
        "to_date":   "Year Ended 31 Mar 2026",
        "period":    "Quarterly",
    },
    "approvals_hist": {
        "id":        "2075",
        "series":    ["Residential dwelling units (Private); New Houses"],
        "from_date": "Jul 2001",
        "to_date":   "Dec 2018",
        "period":    "Monthly",
    },
    "approvals_curr": {
        "id":        "2031",
        "series":    ["Residential dwelling units (Private); New Houses"],
        "from_date": "Jan 2019",
        "to_date":   "Jan 2026",
        "period":    "Monthly",
    },
}


class QGSOHousingFetcher(BaseFetcher):

    SOURCE_NAME      = "qgso_housing"
    SUPPORTED_STATES = ["QLD"]

    def fetch_all(self):
        towns = self.applicable_towns()
        if not towns:
            self.log.info("No QLD towns configured — nothing to fetch")
            return

        # Build SA2 code → town map using qgso_sa2 (preferred) or sa2_code
        town_sa2_map: dict[str, object] = {}
        for town in towns:
            sa2 = getattr(town, "qgso_sa2", None) or town.sa2_code
            if sa2:
                town_sa2_map[str(sa2)] = town
            else:
                self.log.warning(f"  [{town.name}] no qgso_sa2 or sa2_code — skipping")

        if not town_sa2_map:
            self.result.add_error("ALL", "No SA2 codes available for QGSO lookup")
            return

        # Fetch each collection and accumulate by town slug
        all_data: dict[str, dict] = {}

        for coll_key, cfg in COLLECTIONS.items():
            self.log.info(f"  Collection: {coll_key} (id={cfg['id']})")
            try:
                coll_data = self._fetch_collection(coll_key, cfg, town_sa2_map)
                for slug, indicators in coll_data.items():
                    if slug not in all_data:
                        all_data[slug] = {}
                    # For approvals, merge historical + current
                    for ind_name, values in indicators.items():
                        if ind_name in all_data[slug] and isinstance(values, dict):
                            all_data[slug][ind_name].update(values)
                        else:
                            all_data[slug][ind_name] = values
            except Exception as exc:
                self.log.error(f"  Collection {coll_key} failed: {exc}", exc_info=True)
                self.result.add_error("ALL", f"Collection {coll_key} failed: {exc}")

        # Write cache JSON per town
        out_dir = CACHE_DIR / "housing"
        out_dir.mkdir(exist_ok=True)

        for town in towns:
            indicators = all_data.get(town.slug, {})
            if not indicators:
                self.log.warning(f"  [{town.name}] no housing data retrieved")
                self.result.towns_failed.append(town.name)
                continue

            out = {
                "town":       town.name,
                "state":      town.state,
                "source":     "QGSO Regional Database (QRSIS)",
                "source_url": (
                    "https://www.qgso.qld.gov.au/statistics/queensland-regions/"
                    "regional-tools-statistics/queensland-regional-database"
                ),
                "note": (
                    "Sales count = Dec-quarter rolling 12-month total. "
                    "Median price and rent = mean of 4 quarterly medians. "
                    "Building approvals = sum of 12 monthly values."
                ),
                "indicators": indicators,
            }

            out_path = out_dir / f"{town.slug}_qgso.json"
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2)

            summary = ", ".join(
                f"{k}({len(v)}yr)" for k, v in indicators.items()
            )
            self.log.info(f"  {town.name}: {summary}")
            self.result.towns_ok.append(town.name)

    # ── Per-collection fetch ─────────────────────────────────────────────────────

    def _fetch_collection(
        self,
        coll_key: str,
        cfg: dict,
        town_sa2_map: dict,
    ) -> dict[str, dict]:
        """Run one QRSIS wizard session. Returns { slug: { indicator: {yr: val} } }."""
        session = requests.Session()
        session.headers.update(HEADERS)

        # Step 1: Select collection → get udqctl_id
        udqctl_id = self._select_collection(session, cfg["id"])
        if not udqctl_id:
            raise RuntimeError(f"Failed to get udqctl_id for collection {cfg['id']}")
        self.log.info(f"    udqctl_id={udqctl_id}")

        # Step 2: Get available series, match, select
        available_series = self._get_available_series(session, udqctl_id)
        matched_series   = self._match_series(cfg["series"], available_series)
        if not matched_series:
            raise RuntimeError(f"No series matched for collection {cfg['id']}")
        self._select_series(session, udqctl_id, matched_series)
        self.log.info(f"    Series selected: {matched_series}")

        # Step 3: Set time period
        self._set_time_period(session, udqctl_id, cfg["id"], cfg["from_date"], cfg["to_date"])

        # Step 4: Select SA2 region type
        self._select_region_type(session, udqctl_id)

        # Step 5: Get available regions, match our SA2s, select
        available_regions = self._get_available_regions(session, udqctl_id)
        matched_regions   = self._match_regions(town_sa2_map, available_regions)
        if not matched_regions:
            raise RuntimeError("No SA2 regions matched in QRSIS region list")
        self._select_regions(session, udqctl_id, matched_regions)
        self.log.info(f"    Regions selected: {len(matched_regions)}")

        # Step 6: Submit and parse
        html = self._submit_report(session, udqctl_id, cfg["id"])
        if not html:
            raise RuntimeError("Empty output from QRSIS report")

        raw = self._parse_output_html(html)
        self.log.info(f"    Parsed data for {len(raw)} SA2 regions")

        return self._aggregate(coll_key, matched_series, raw, town_sa2_map)

    # ── QRSIS wizard steps ───────────────────────────────────────────────────────

    def _select_collection(self, session: requests.Session, coll_id: str) -> Optional[str]:
        """Step 1: POST to ProcessCollection, extract udqctl_id."""
        resp = session.post(
            BASE_URL + "QIS1110W$COLL.ProcessCollection",
            data={
                "p_names":  ["usr_id",    "access_lvl",  "coll_id", "collgrp_id", "sel_coll_name", "op_mode"],
                "p_values": [PUBLIC_USER, ACCESS_LEVEL,  "",        COLLGRP_ID,   coll_id,          "Next"],
            },
            timeout=30,
            allow_redirects=True,
        )
        resp.raise_for_status()
        return self._extract_udqctl_id(resp.url, resp.text)

    def _get_available_series(self, session: requests.Session, udqctl_id: str) -> list[str]:
        """Step 2a: GET series page, return available option texts."""
        resp = session.get(
            BASE_URL + "QIS1110W$UDQSER.ProcessSeries",
            params=[
                ("p_names",  "op_mode"),    ("p_values", "VIEW"),
                ("p_names",  "info_page"),  ("p_values", "infoser.htm"),
                ("p_names",  "udqctl_id"),  ("p_values", udqctl_id),
                ("p_names",  "error_msg"),  ("p_values", ""),
            ],
            timeout=30,
        )
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        for sel in soup.find_all("select", {"name": "p_new_multi"}):
            return [opt.get_text(strip=True) for opt in sel.find_all("option") if opt.get_text(strip=True)]
        return []

    def _match_series(self, wanted: list[str], available: list[str]) -> list[str]:
        """Match wanted series against available, with rent fallbacks."""
        matched = []
        for want in wanted:
            if want in available:
                matched.append(want)
                continue
            # Rent fallbacks
            candidates = RENT_SERIES_CANDIDATES if "rent" in want.lower() or "bedroom" in want.lower() else []
            found = False
            for fb in candidates:
                if fb in available:
                    self.log.warning(f"    Series fallback: '{want}' → '{fb}'")
                    matched.append(fb)
                    found = True
                    break
            if not found:
                # Loose partial match on first 25 chars
                prefix = want[:25].lower()
                for avail in available:
                    if prefix in avail.lower():
                        self.log.warning(f"    Series partial match: '{want}' → '{avail}'")
                        matched.append(avail)
                        found = True
                        break
            if not found:
                self.log.error(f"    Series not found: '{want}'  Available (first 5): {available[:5]}")
        return matched

    def _select_series(self, session: requests.Session, udqctl_id: str, series: list[str]):
        """Step 2b: Move each series to Selected, then click Next."""
        for s in series:
            session.post(
                BASE_URL + "QIS1110W$UDQSER.ProcessActions",
                data={
                    "p_names":     ["udqctl_id", "info_page",   "error_msg", "op_mode"],
                    "p_values":    [udqctl_id,   "infoser.htm", "",          "->"],
                    "p_new_multi": s,
                },
                timeout=30,
            )
            time.sleep(0.2)
        # Next
        session.post(
            BASE_URL + "QIS1110W$UDQSER.ProcessActions",
            data={
                "p_names":  ["udqctl_id", "info_page",   "error_msg", "op_mode"],
                "p_values": [udqctl_id,   "infoser.htm", "",          "Next"],
            },
            timeout=30,
        )

    def _set_time_period(
        self, session: requests.Session, udqctl_id: str, coll_id: str, from_date: str, to_date: str
    ):
        """Step 3: POST time period selection."""
        session.post(
            BASE_URL + "QIS1110W$UDQCTL.ProcessActions",
            data={
                "p_names":  ["udqctl_id", "coll_id", "error_msg", "date_format", "period",    "p_concorded_data", "from_date", "to_date"],
                "p_values": [udqctl_id,   coll_id,   "",          "Y1",          "Quarterly", "N",                from_date,   to_date],
                "p_op_mode": "Next",
            },
            timeout=30,
        )

    def _select_region_type(self, session: requests.Session, udqctl_id: str):
        """Step 4: Move SA2 to selected region type, then Next."""
        session.post(
            BASE_URL + "QIS1110W$REGTYP.ProcessRegType",
            data={
                "p_names":  ["udqctl_id", "op_mode"],
                "p_values": [udqctl_id,   "->"],
                "p_multi":  "SA2 - Statistical Area Level 2",
            },
            timeout=30,
        )
        time.sleep(0.2)
        session.post(
            BASE_URL + "QIS1110W$REGTYP.ProcessRegType",
            data={
                "p_names":  ["udqctl_id", "op_mode"],
                "p_values": [udqctl_id,   "Next"],
            },
            timeout=30,
        )

    def _get_available_regions(self, session: requests.Session, udqctl_id: str) -> list[str]:
        """Step 5a: GET region page, return available option texts."""
        resp = session.get(
            BASE_URL + "QIS1110W$UDQREG.ProcessRegions",
            params=[
                ("p_names",  "op_mode"),    ("p_values", "VIEW"),
                ("p_names",  "info_page"),  ("p_values", "inforeg.htm"),
                ("p_names",  "udqctl_id"),  ("p_values", udqctl_id),
                ("p_names",  "error_msg"),  ("p_values", ""),
            ],
            timeout=30,
        )
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        for sel in soup.find_all("select", {"name": "p_new_multi"}):
            return [opt.get_text(strip=True) for opt in sel.find_all("option") if opt.get_text(strip=True)]
        return []

    def _match_regions(self, town_sa2_map: dict, available: list[str]) -> list[str]:
        """Match SA2 codes against available region strings."""
        matched = []
        for sa2_code, town in town_sa2_map.items():
            prefix = f"SA2/{sa2_code}"
            hits = [r for r in available if r.startswith(prefix)]
            if hits:
                matched.extend(hits)
                self.log.info(f"    Matched {town.name}: {hits[0]}")
            else:
                self.log.warning(f"    [{town.name}] {prefix} not in QRSIS list")
        return matched

    def _select_regions(self, session: requests.Session, udqctl_id: str, regions: list[str]):
        """Step 5b: Move all regions to Selected at once, then Next."""
        session.post(
            BASE_URL + "QIS1110W$UDQREG.ProcessActions",
            data={
                "p_names":     ["udqctl_id", "info_page",   "error_msg", "op_mode"],
                "p_values":    [udqctl_id,   "inforeg.htm", "",          "->"],
                "p_new_multi": regions,
            },
            timeout=60,
        )
        time.sleep(0.5)
        session.post(
            BASE_URL + "QIS1110W$UDQREG.ProcessActions",
            data={
                "p_names":  ["udqctl_id", "info_page",   "error_msg", "op_mode"],
                "p_values": [udqctl_id,   "inforeg.htm", "",          "Next"],
            },
            timeout=30,
        )

    def _submit_report(self, session: requests.Session, udqctl_id: str, coll_id: str) -> str:
        """Step 6: POST report parameters, return output HTML."""
        resp = session.post(
            BASE_URL + "QIS1110W$UDQCTL1.ProcessActions",
            data={
                "p_names":  ["udqctl_id", "coll_id", "error_msg", "ser_sort_col", "reg_sort_col",
                             "display_style",                               "op_mode"],
                "p_values": [udqctl_id,   coll_id,   "",          "Sort Number",  "Region Code",
                             "For each Region display Time Period by Series", "QRSIS Query"],
            },
            timeout=120,
        )
        resp.raise_for_status()
        return resp.text

    # ── Helpers ──────────────────────────────────────────────────────────────────

    def _extract_udqctl_id(self, url: str, html: str) -> Optional[str]:
        """Extract udqctl_id from redirect URL or HTML hidden fields."""
        for pattern in [
            r'udqctl_id[=&](\d+)',
            r'p_udqctl_id[=&](\d+)',
            r'ProcessSeries[^"\']*udqctl_id[=&](\d+)',
        ]:
            m = re.search(pattern, url)
            if m:
                return m.group(1)
        # Fall back to HTML
        m = re.search(r'VALUE="(\d{3,6})"', html)
        if m:
            return m.group(1)
        return None

    # ── HTML parsing ─────────────────────────────────────────────────────────────

    def _parse_output_html(self, html: str) -> dict:
        """
        Parse QRSIS output HTML into:
          { sa2_code_str: { period_str: { series_name: float|None } } }

        The output page contains one table per selected SA2, preceded by a
        "Region : SA2/XXXXXXXXX - Name" heading.
        """
        result: dict[str, dict] = {}

        # Split on "Region : SA2/" occurrences — each section is one town
        sections = re.split(r'Region\s*:\s*SA2/', html)
        for section in sections[1:]:   # skip preamble before first region
            sa2_match = re.match(r'(\d+)', section)
            if not sa2_match:
                continue
            sa2_code = sa2_match.group(1)
            region_data: dict[str, dict] = {}

            soup = BeautifulSoup(section, "html.parser")
            table = soup.find("table", attrs={"border": True}) or soup.find("table")
            if not table:
                result[sa2_code] = region_data
                continue

            rows = table.find_all("tr")
            if not rows:
                result[sa2_code] = region_data
                continue

            # Header: Period | Series1 | Series2 ...
            header_cells = rows[0].find_all(["th", "td"])
            series_names = [c.get_text(strip=True) for c in header_cells[1:]]

            for row in rows[1:]:
                cells = row.find_all(["td", "th"])
                if not cells:
                    continue
                period = cells[0].get_text(strip=True)
                if not period:
                    continue
                period_data: dict[str, Optional[float]] = {}
                for i, sname in enumerate(series_names, start=1):
                    if i < len(cells):
                        raw = cells[i].get_text(strip=True).replace(",", "").replace("$", "").strip()
                        try:
                            period_data[sname] = float(raw)
                        except ValueError:
                            period_data[sname] = None   # suppressed (< 10 sales etc.)
                region_data[period] = period_data

            result[sa2_code] = region_data

        return result

    # ── Annual aggregation ───────────────────────────────────────────────────────

    def _aggregate(
        self,
        coll_key: str,
        matched_series: list[str],
        raw: dict,
        town_sa2_map: dict,
    ) -> dict[str, dict]:
        """Aggregate raw period data to annual indicators by town slug."""
        from config import YEAR_START, YEAR_END
        import statistics as stats

        result: dict[str, dict] = {}

        for sa2_code, town in town_sa2_map.items():
            if sa2_code not in raw:
                continue
            periods = raw[sa2_code]
            slug    = town.slug

            if coll_key == "sales":
                count_series = "Detached dwelling: number of sales (Number)"
                price_series = "Detached dwelling: median sale price ($)"

                # Sales count: Dec-quarter rolling total = calendar year total
                count_by_year: dict[str, int] = {}
                for period, vals in periods.items():
                    if "31 Dec" not in period:
                        continue
                    year = self._parse_year(period)
                    if year and YEAR_START <= year <= YEAR_END:
                        v = vals.get(count_series)
                        if v is not None:
                            count_by_year[str(year)] = int(v)

                # Median price: mean of 4 quarterly medians per year
                price_by_year: dict[str, list] = {}
                for period, vals in periods.items():
                    year = self._parse_year(period)
                    if not year or not (YEAR_START <= year <= YEAR_END):
                        continue
                    v = vals.get(price_series)
                    if v is not None:
                        price_by_year.setdefault(str(year), []).append(v)
                price_annual = {yr: round(stats.mean(vs)) for yr, vs in price_by_year.items() if vs}

                result[slug] = {
                    "housing_sales_count":  count_by_year,
                    "housing_median_price": price_annual,
                }

            elif coll_key == "rent":
                # Find whichever series is present
                series = None
                for period_vals in periods.values():
                    for s in RENT_SERIES_CANDIDATES:
                        if s in period_vals:
                            series = s
                            break
                    if series:
                        break
                if not series and matched_series:
                    series = matched_series[0]

                rent_by_year: dict[str, list] = {}
                for period, vals in periods.items():
                    year = self._parse_year(period)
                    if not year or not (YEAR_START <= year <= YEAR_END):
                        continue
                    v = vals.get(series) if series else None
                    if v is not None:
                        rent_by_year.setdefault(str(year), []).append(v)
                rent_annual = {yr: round(stats.mean(vs)) for yr, vs in rent_by_year.items() if vs}

                result[slug] = {"rent_3bed_median": rent_annual}

            elif coll_key in ("approvals_hist", "approvals_curr"):
                series = matched_series[0] if matched_series else None
                approvals_by_year: dict[str, int] = {}
                for period, vals in periods.items():
                    parsed = self._parse_month_year(period)
                    if not parsed:
                        continue
                    year, _month = parsed
                    if not (YEAR_START <= year <= YEAR_END):
                        continue
                    v = vals.get(series) if series else None
                    if v is not None:
                        approvals_by_year[str(year)] = approvals_by_year.get(str(year), 0) + int(v)
                result[slug] = {"building_approvals": approvals_by_year}

        return result

    @staticmethod
    def _parse_year(period: str) -> Optional[int]:
        """Extract year from 'Year Ended 31 Dec 2024' etc."""
        m = re.search(r'\b(20\d{2}|19\d{2})\b', period)
        return int(m.group(1)) if m else None

    @staticmethod
    def _parse_month_year(period: str) -> Optional[tuple[int, int]]:
        """Parse 'Jul 2001' → (2001, 7)."""
        months = {
            "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
            "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
        }
        m = re.match(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})', period.strip())
        return (int(m.group(2)), months[m.group(1)]) if m else None


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = QGSOHousingFetcher().run()
    sys.exit(0 if result.success else 1)