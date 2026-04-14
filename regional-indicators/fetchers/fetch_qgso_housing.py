"""
fetchers/fetch_qgso_housing.py
-------------------------------
Fetches housing indicators from the QGSO Regional Database (QRSIS).
QLD only.

COLLECTIONS FETCHED:
  1925  Residential land and dwelling sales   (Sep 2000 – Sep 2025, quarterly, SA2)
  1929  Median rent                            (Dec 1989 – Mar 2026, quarterly, SA2)
  2075  Building Approvals (Historical)        (Jul 2001 – Dec 2018, monthly, LGA)
  2031  Building Approvals (Current)           (Jan 2019 – present, monthly, LGA)

INDICATORS PRODUCED (per town, annual):
  housing_sales_count    Detached dwelling: number of sales
  housing_median_price   Detached dwelling: median sale price ($)
  rent_3bed_median       House - 3 bedrooms - median rent of lodgements ($/week)
  building_approvals     Residential dwelling units (Private); New Houses (Number)

CRITICAL IMPLEMENTATION NOTES:
  1. POST encoding: Oracle PL/SQL WebTK reads p_names/p_values as interleaved
     parallel arrays. Must use list of (key,val) tuples — dicts group all
     p_names first then all p_values which Oracle misreads. Use _q() helper.

  2. HTML parsing: QRSIS serves HTML with unclosed <OPTION> tags. BeautifulSoup's
     html.parser merges option texts into one string. Must use lxml parser.

  3. udqctl_id extraction: Redirect URL uses interleaved format
     ?p_names=udqctl_id&p_values=3908, not ?udqctl_id=3908.

  4. Building approvals geography: SA2-level data not available for most small
     QLD towns. Use LGA-level (qgso_lga field from towns.toml) instead.

  5. Exact series names from QRSIS (verified from live session):
     - 'Detached dwelling: number of sales (Number)'
     - 'Detached dwelling: median sale price ($)'
     - 'House - 3 bedrooms - median rent of lodgements ($/week)'
     - 'Residential dwelling units (Private); New Houses (Number)'

SOURCE:
  https://www.qgso.qld.gov.au/statistics/queensland-regions/regional-tools-statistics/queensland-regional-database
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
    raise ImportError("pip install requests beautifulsoup4 lxml")


# ── QRSIS constants ─────────────────────────────────────────────────────────────

BASE_URL     = "https://statistics.qgso.qld.gov.au/pls/qis_public/"
PUBLIC_USER  = "edtert"
ACCESS_LEVEL = "85"
COLLGRP_ID   = "22"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": "https://statistics.qgso.qld.gov.au/",
    "Origin":  "https://statistics.qgso.qld.gov.au",
}

# Exact series names as returned by QRSIS (verified from live session)
SERIES_SALES_COUNT = "Detached dwelling: number of sales (Number)"
SERIES_SALES_PRICE = "Detached dwelling: median sale price ($)"
SERIES_RENT        = "House - 3 bedrooms - median rent of lodgements ($/week)"
SERIES_APPROVALS   = "Residential dwelling units (Private); New Houses (Number)"

COLLECTIONS = {
    "sales": {
        "id":         "1925",
        "series":     [SERIES_SALES_COUNT, SERIES_SALES_PRICE],
        "from_date":  "Year Ended 30 Sep 2000",
        "to_date":    "Year Ended 30 Sep 2025",
        "geo":        "SA2",       # region type to select
    },
    "rent": {
        "id":         "1929",
        "series":     [SERIES_RENT],
        "from_date":  "Year Ended 31 Dec 2000",
        "to_date":    "Year Ended 31 Mar 2026",
        "geo":        "SA2",
    },
    "approvals_hist": {
        "id":         "2075",
        "series":     [SERIES_APPROVALS],
        "from_date":  "Jul 2001",
        "to_date":    "Dec 2018",
        "period":     "Monthly",
        "date_fmt":   "M1",
        "concorded":  "Y",
        "geo":        "SA2",
        "approvals":  True,
    },
    "approvals_curr": {
        "id":         "2031",
        "series":     [SERIES_APPROVALS],
        "from_date":  "Jan 2019",
        "to_date":    "Jan 2026",
        "period":     "Monthly",
        "date_fmt":   "M1",
        "concorded":  "Y",
        "geo":        "SA2",
        "approvals":  True,
    },
}



# Toowoomba approvals use LGA boundary per previous booklets (not SA2)
# All other towns use SA2. LGA code for Toowoomba = LGA/36910
TOOWOOMBA_LGA_SLUGS = {"toowoomba", "toowoomba_central", "toowoomba_harlaxton", "toowoomba_west"}
TOOWOOMBA_LGA_CODE  = "LGA/36910"

def _q(*pairs) -> list[tuple]:
    """Build interleaved p_names/p_values list from alternating (name, value) args."""
    result = []
    it = iter(pairs)
    for name in it:
        value = next(it)
        result.append(("p_names", name))
        result.append(("p_values", value))
    return result


def _parse_options(html: str, select_name: str) -> list[str]:
    """
    Parse <option> text values from a named <select>.
    MUST use lxml — html.parser merges unclosed <OPTION> tags into one string.
    """
    soup = BeautifulSoup(html, "lxml")
    for sel in soup.find_all("select", {"name": select_name}):
        opts = [o.get_text(strip=True) for o in sel.find_all("option")]
        return [o for o in opts if o]
    return []


class QGSOHousingFetcher(BaseFetcher):

    SOURCE_NAME      = "qgso_housing"
    SUPPORTED_STATES = ["QLD"]

    def fetch_all(self):
        towns = self.applicable_towns()
        if not towns:
            self.log.info("No QLD towns configured — nothing to fetch")
            return

        # SA2 map: sa2_code → town (for sales/rent)
        sa2_map: dict[str, list] = {}  # sa2_code → [town, ...] (multiple towns may share an SA2)
        for town in towns:
            sa2 = getattr(town, "qgso_sa2", None) or town.sa2_code
            if sa2:
                sa2_map.setdefault(str(sa2), []).append(town)

        if not sa2_map:
            self.result.add_error("ALL", "No SA2 codes available for QGSO lookup")
            return

        # For building approvals: most towns use SA2, but Toowoomba uses LGA
        # (per previous booklets: Roma/Chinchilla/etc = SA2, Toowoomba = LGA)
        approvals_sa2_map: dict[str, list] = {}
        approvals_lga_map: dict[str, list] = {}
        for town in towns:
            if town.slug in TOOWOOMBA_LGA_SLUGS:
                approvals_lga_map.setdefault(TOOWOOMBA_LGA_CODE, []).append(town)
            else:
                sa2 = getattr(town, "qgso_sa2", None) or town.sa2_code
                if sa2:
                    approvals_sa2_map.setdefault(str(sa2), []).append(town)
                    # Note: towns with shared SA2 (Roma+Wallumbilla, Miles+Wandoan)
                    # both get added here so both receive the data

        all_data: dict[str, dict] = {}

        for coll_key, cfg in COLLECTIONS.items():
            self.log.info(f"  Collection: {coll_key} (id={cfg['id']})")
            if cfg.get("approvals"):
                # Merge SA2 and LGA region maps for approvals
                region_map = {**{k: v for k, v in approvals_sa2_map.items()},
                              **{k: v for k, v in approvals_lga_map.items()}}
            else:
                region_map = sa2_map  # values are already lists of towns
            if not region_map:
                self.log.warning(f"  No regions for {coll_key} — skipping")
                continue
            try:
                coll_data = self._fetch_collection(coll_key, cfg, region_map)
                for slug, indicators in coll_data.items():
                    if slug not in all_data:
                        all_data[slug] = {}
                    for ind_name, values in indicators.items():
                        if ind_name in all_data[slug] and isinstance(values, dict):
                            all_data[slug][ind_name].update(values)
                        else:
                            all_data[slug][ind_name] = values
            except Exception as exc:
                self.log.error(f"  Collection {coll_key} failed: {exc}", exc_info=True)
                self.result.add_error("ALL", f"Collection {coll_key} failed: {exc}")

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
                    "Building approvals = sum of 12 monthly values (LGA level)."
                ),
                "indicators": indicators,
            }

            out_path = out_dir / f"{town.slug}_qgso.json"
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2)

            summary = ", ".join(f"{k}({len(v)}yr)" for k, v in indicators.items())
            self.log.info(f"  {town.name}: {summary}")
            self.result.towns_ok.append(town.name)

    # ── Per-collection fetch ─────────────────────────────────────────────────────

    def _fetch_collection(self, coll_key, cfg, region_map) -> dict[str, dict]:
        """
        region_map: { region_code: [town, ...] }
          For SA2: { "307011176": [roma_town] }
          For LGA: { "LGA/34860": [roma_town, wallumbilla_town] }
        """
        session = requests.Session()
        session.headers.update(HEADERS)

        # Step 1: select collection
        udqctl_id = self._select_collection(session, cfg["id"])
        if not udqctl_id:
            raise RuntimeError(f"Failed to get udqctl_id for collection {cfg['id']}")
        self.log.info(f"    udqctl_id={udqctl_id}")

        # Step 2: series — get available, match exact names, select
        avail_series  = self._get_options(session, "QIS1110W$UDQSER.ProcessSeries",
                                          udqctl_id, "infoser.htm", "p_new_multi")
        self.log.info(f"    Available series: {avail_series}")
        matched = self._match_series(cfg["series"], avail_series)
        if not matched:
            raise RuntimeError(f"No series matched for collection {cfg['id']}")
        self._select_series(session, udqctl_id, matched)
        self.log.info(f"    Series selected: {matched}")

        # Step 3: time period
        self._set_time_period(session, udqctl_id, cfg["id"], cfg)

        # Step 4: region type
        # Determine which region type to select based on region codes
        # For mixed SA2+LGA approvals, select both types
        # Check if region_map contains LGA codes (for Toowoomba approvals)
        has_lga = any(str(k).startswith("LGA/") for k in region_map)
        has_sa2 = any(not str(k).startswith("LGA/") for k in region_map)
        geo_label = "SA2 - Statistical Area Level 2"  # default
        if has_sa2:
            self._select_region_type(session, udqctl_id, "SA2 - Statistical Area Level 2")
        if has_lga:
            self._select_region_type(session, udqctl_id, "LGA - Local Government Area")

        # Step 5: regions
        avail_regions   = self._get_options(session, "QIS1110W$UDQREG.ProcessRegions",
                                            udqctl_id, "inforeg.htm", "p_new_multi")
        matched_regions = self._match_regions(region_map, avail_regions, cfg["geo"])
        if not matched_regions:
            raise RuntimeError(f"No {cfg['geo']} regions matched in QRSIS region list")
        self._select_regions(session, udqctl_id, matched_regions)
        self.log.info(f"    Regions selected: {len(matched_regions)}")

        # Step 6: submit
        html = self._submit_report(session, udqctl_id, cfg["id"])
        raw  = self._parse_output_html(html)
        self.log.info(f"    Parsed data for {len(raw)} regions")

        return self._aggregate(coll_key, matched, raw, region_map, cfg["geo"])

    # ── QRSIS wizard steps ───────────────────────────────────────────────────────

    def _select_collection(self, session, coll_id) -> Optional[str]:
        resp = session.post(
            BASE_URL + "QIS1110W$COLL.ProcessCollection",
            data=_q("usr_id", PUBLIC_USER, "access_lvl", ACCESS_LEVEL,
                    "coll_id", "", "collgrp_id", COLLGRP_ID,
                    "sel_coll_name", coll_id, "op_mode", "Next"),
            timeout=30, allow_redirects=True,
        )
        resp.raise_for_status()
        return self._extract_udqctl_id(resp.url, resp.text)

    def _get_options(self, session, proc, udqctl_id, info_page, select_name) -> list[str]:
        """GET a QRSIS page and return the options from the named select."""
        resp = session.get(
            BASE_URL + proc,
            params=_q("op_mode", "VIEW", "info_page", info_page,
                      "udqctl_id", udqctl_id, "error_msg", ""),
            timeout=30,
        )
        resp.raise_for_status()
        return _parse_options(resp.text, select_name)

    def _match_series(self, wanted: list[str], available: list[str]) -> list[str]:
        matched = []
        for want in wanted:
            if want in available:
                matched.append(want)
            else:
                self.log.error(f"    Series not found: '{want}'")
                self.log.error(f"    Available: {available}")
        return matched

    def _select_series(self, session, udqctl_id, series):
        for s in series:
            session.post(
                BASE_URL + "QIS1110W$UDQSER.ProcessActions",
                data=_q("udqctl_id", udqctl_id, "info_page", "infoser.htm",
                        "error_msg", "", "op_mode", "->") + [("p_new_multi", s)],
                timeout=30,
            )
            time.sleep(0.2)
        session.post(
            BASE_URL + "QIS1110W$UDQSER.ProcessActions",
            data=_q("udqctl_id", udqctl_id, "info_page", "infoser.htm",
                    "error_msg", "", "op_mode", "Next"),
            timeout=30,
        )

    def _set_time_period(self, session, udqctl_id, coll_id, cfg):
        from_date  = cfg["from_date"]
        to_date    = cfg["to_date"]
        period     = cfg.get("period", "Quarterly")
        date_fmt   = cfg.get("date_fmt", "Y1")
        concorded  = cfg.get("concorded", "N")
        data = _q("udqctl_id", udqctl_id, "coll_id", coll_id, "error_msg", "",
                  "date_format", date_fmt, "period", period,
                  "from_date", from_date, "to_date", to_date)
        # concorded data: sales/rent use hidden field "N", approvals use checkbox "Y"
        if concorded == "Y":
            data.append(("p_concorded_data", "Y"))
        data.append(("p_op_mode", "Next"))
        session.post(BASE_URL + "QIS1110W$UDQCTL.ProcessActions", data=data, timeout=30)

    def _select_region_type(self, session, udqctl_id, geo_label):
        session.post(
            BASE_URL + "QIS1110W$REGTYP.ProcessRegType",
            data=_q("udqctl_id", udqctl_id, "op_mode", "->")
                + [("p_multi", geo_label)],
            timeout=30,
        )
        time.sleep(0.2)
        session.post(
            BASE_URL + "QIS1110W$REGTYP.ProcessRegType",
            data=_q("udqctl_id", udqctl_id, "op_mode", "Next"),
            timeout=30,
        )

    def _match_regions(self, region_map, available, geo) -> list[str]:
        if available:
            self.log.info(f"    First available region: {available[0][:80]}")
        matched = []
        for code, towns in region_map.items():
            town_names = [t.name for t in towns] if isinstance(towns, list) else [towns.name]
            if str(code).startswith("LGA/"):
                # Already has prefix
                prefix = str(code)
            elif geo == "SA3":
                prefix = f"SA3/{code}"
            elif geo == "LGA":
                prefix = code  # already "LGA/NNNNN"
            else:
                prefix = f"SA2/{code}"
            hits = [r for r in available if r.startswith(prefix)]
            if hits:
                matched.extend(hits)
                self.log.info(f"    Matched {'/'.join(town_names)}: {hits[0][:60]}")
            else:
                self.log.warning(f"    [{'/'.join(town_names)}] {prefix} not in QRSIS list")
        return matched

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

    def _submit_report(self, session, udqctl_id, coll_id) -> str:
        resp = session.post(
            BASE_URL + "QIS1110W$UDQCTL1.ProcessActions",
            data=_q("udqctl_id", udqctl_id, "coll_id", coll_id, "error_msg", "",
                    "ser_sort_col", "Sort Number", "reg_sort_col", "Region Code",
                    "display_style", "For each Region display Time Period by Series",
                    "op_mode", "QRSIS Query"),
            timeout=120,
        )
        resp.raise_for_status()
        return resp.text

    # ── Helpers ──────────────────────────────────────────────────────────────────

    def _extract_udqctl_id(self, url: str, html: str) -> Optional[str]:
        m = re.search(r'p_names=udqctl_id&p_values=(\d+)', url)
        if m:
            return m.group(1)
        m = re.search(r'[?&]udqctl_id=(\d+)', url)
        if m:
            return m.group(1)
        m = re.search(
            r'NAME="p_names"\s+VALUE="udqctl_id"[^>]*>.*?NAME="p_values"\s+VALUE="(\d+)"',
            html, re.IGNORECASE | re.DOTALL
        )
        if m:
            return m.group(1)
        m = re.search(r'udqctl_id.*?(\d{4,6})', url + html, re.DOTALL)
        if m:
            return m.group(1)
        return None

    # ── HTML parsing ─────────────────────────────────────────────────────────────

    def _parse_output_html(self, html: str) -> dict:
        """
        Parse QRSIS output HTML. Returns:
          { region_code: { period: { series_name: float|None } } }
        Region code is the SA2 number or LGA code (e.g. "307011176" or "LGA/34860")
        """
        result: dict[str, dict] = {}

        # Output has sections like "Region : SA2/307011176 - Roma" or
        # "Region : LGA/34860 - Maranoa (R)"
        # Split on "Region :" to get one section per region
        sections = re.split(r'Region\s*:\s*', html)
        for section in sections[1:]:
            # Extract region code — SA2/NNNNN, SA3/NNNNN or LGA/NNNNN
            m = re.match(r'((?:SA2|SA3|SA4|LGA)/[\w]+)', section)
            if not m:
                continue
            region_code = m.group(1)
            region_data: dict[str, dict] = {}

            soup  = BeautifulSoup(section, "lxml")
            table = soup.find("table", border=True) or soup.find("table")
            if not table:
                result[region_code] = region_data
                continue

            rows = table.find_all("tr")
            if not rows:
                result[region_code] = region_data
                continue

            header_cells = rows[0].find_all(["th", "td"])
            series_names = [c.get_text(strip=True) for c in header_cells[1:]]

            for row in rows[1:]:
                cells = row.find_all(["td", "th"])
                if not cells:
                    continue
                period = cells[0].get_text(strip=True)
                if not period:
                    continue
                period_data: dict = {}
                for i, sname in enumerate(series_names, start=1):
                    if i < len(cells):
                        raw = cells[i].get_text(strip=True).replace(",", "").replace("$", "").strip()
                        try:
                            period_data[sname] = float(raw)
                        except ValueError:
                            period_data[sname] = None
                region_data[period] = period_data

            result[region_code] = region_data

        return result

    # ── Annual aggregation ───────────────────────────────────────────────────────

    def _aggregate(self, coll_key, matched_series, raw, region_map, geo) -> dict:
        from config import YEAR_START, YEAR_END
        import statistics as stats

        result: dict[str, dict] = {}

        # Build region_code → list of towns
        for code, towns in region_map.items():
            if str(code).startswith("LGA/"):
                region_key = str(code)
            elif geo == "SA3":
                region_key = f"SA3/{code}"
            elif geo == "LGA":
                region_key = code
            else:
                region_key = f"SA2/{code}"
            periods = raw.get(region_key) or raw.get(code) or {}
            if not periods:
                continue

            town_list = towns if isinstance(towns, list) else [towns]

            for town in town_list:
                slug = town.slug

                if coll_key == "sales":
                    count_yr: dict[str, int] = {}
                    for period, vals in periods.items():
                        if "31 Dec" not in period:
                            continue
                        yr = self._parse_year(period)
                        if yr and YEAR_START <= yr <= YEAR_END:
                            v = vals.get(SERIES_SALES_COUNT)
                            if v is not None:
                                count_yr[str(yr)] = int(v)

                    price_yr: dict[str, list] = {}
                    for period, vals in periods.items():
                        yr = self._parse_year(period)
                        if not yr or not (YEAR_START <= yr <= YEAR_END):
                            continue
                        v = vals.get(SERIES_SALES_PRICE)
                        if v is not None:
                            price_yr.setdefault(str(yr), []).append(v)
                    price_annual = {yr: round(stats.mean(vs)) for yr, vs in price_yr.items() if vs}

                    result[slug] = {
                        "housing_sales_count":  count_yr,
                        "housing_median_price": price_annual,
                    }

                elif coll_key == "rent":
                    rent_yr: dict[str, list] = {}
                    for period, vals in periods.items():
                        yr = self._parse_year(period)
                        if not yr or not (YEAR_START <= yr <= YEAR_END):
                            continue
                        v = vals.get(SERIES_RENT)
                        if v is not None:
                            rent_yr.setdefault(str(yr), []).append(v)
                    result[slug] = {"rent_3bed_median": {yr: round(stats.mean(vs)) for yr, vs in rent_yr.items() if vs}}

                elif coll_key in ("approvals_hist", "approvals_curr"):
                    app_yr: dict[str, int] = {}
                    for period, vals in periods.items():
                        parsed = self._parse_month_year(period)
                        if not parsed:
                            continue
                        yr, _ = parsed
                        if not (YEAR_START <= yr <= YEAR_END):
                            continue
                        v = vals.get(SERIES_APPROVALS)
                        if v is not None:
                            app_yr[str(yr)] = app_yr.get(str(yr), 0) + int(v)
                    # Merge into existing if present (hist + curr)
                    if slug in result and "building_approvals" in result[slug]:
                        result[slug]["building_approvals"].update(app_yr)
                    else:
                        result.setdefault(slug, {})["building_approvals"] = app_yr

        return result

    @staticmethod
    def _parse_year(period: str) -> Optional[int]:
        m = re.search(r'\b(20\d{2}|19\d{2})\b', period)
        return int(m.group(1)) if m else None

    @staticmethod
    def _parse_month_year(period: str) -> Optional[tuple[int, int]]:
        months = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
                  "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
        m = re.match(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})', period.strip())
        return (int(m.group(2)), months[m.group(1)]) if m else None


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = QGSOHousingFetcher().run()
    sys.exit(0 if result.success else 1)