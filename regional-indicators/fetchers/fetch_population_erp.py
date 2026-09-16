"""
regional-indicators/fetchers/fetch_population_erp.py
-------------------------------------------------------
Fetches the main SA2-level "Population (ERP)" figure -- the primary
indicator row in each town's block on the Population sheet, distinct
from fetch_population_ucl.py's UCL-level figure and
fetch_population_nrw.py's non-resident-worker figure.

LIVE-API VERSION (2026-09-16) -- replaces the earlier manual-file-only
fallback. Confirmed real values, read directly from the QRSIS wizard's
live HTML by Steve:
  collgrp_id = "1"     (Population group)
  coll_id    = "1961"  ("Population (ERP)(a) persons only", 1991-2025,
                         ASGS 2021 boundaries -- the current version;
                         there's also an older ASGS 2016 version at
                         coll_id 1298 and an ASGS 2011 one, deliberately
                         not used)

Reuses the exact QRSIS wizard mechanism proven working in
fetch_qgso_housing.py (same _q() POST-encoding, same lxml HTML parsing
for unclosed <OPTION> tags, same udqctl_id extraction, same 6-step
wizard flow: select collection -> series -> time period -> region type
-> regions -> submit). Deliberately duplicated here rather than shared
via a common module for now -- refactoring the proven housing fetcher
to share code carries real risk of breaking something that already
works, with no way to test that live from this environment. Worth
extracting into a shared QRSISFetcherMixin once this fetcher is ALSO
confirmed working live -- see TODO.md.

*** NOT YET TESTED against the live QRSIS endpoint. *** collgrp_id and
coll_id are confirmed real (read directly from the live wizard HTML),
but: the exact series name(s) this collection returns, and the correct
from_date/to_date/period/date_fmt values for an ANNUAL (not quarterly)
series are genuinely unknown -- housing's time-period format strings
are collection-specific, and ERP is likely annual/financial-year based
on the source catalog ("Financial Year" frequency), unlike housing's
quarterly data. Best-effort values below; expect the first live run to
need adjustment, and read its logged "Available series" output to fix
the exact series name if it doesn't match on the first try.

If the live query fails for any reason, falls back to the previously-
working manual-file path: reads a manually-assembled
cache/qgso_and_bom_{YEAR}.xlsx if present, documented in
docs/manual_processes.md.

SUPPORTED_STATES = ["QLD"] -- this is QRSIS-sourced, QLD only. NSW/VIC
towns (Narrabri, Shepparton, Yarram) need a separate ABS-based fetcher
-- see the naming-collision TODO note (this file's name was originally
planned for that ABS fetcher).
"""

from __future__ import annotations

import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR, YEAR_START, YEAR_END

try:
    import requests
    import openpyxl
    from bs4 import BeautifulSoup
except ImportError:
    raise ImportError("pip install requests openpyxl beautifulsoup4 lxml")


# ── QRSIS constants (same system as fetch_qgso_housing.py) ─────────────────

BASE_URL     = "https://statistics.qgso.qld.gov.au/pls/qis_public/"
PUBLIC_USER  = "edtert"
ACCESS_LEVEL = "85"

# Confirmed real values, read from the live wizard HTML (2026-09-16):
COLLGRP_ID = "1"    # Population group
COLL_ID    = "1961" # "Population (ERP)(a) persons only", 1991-2025, ASGS 2021

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": "https://statistics.qgso.qld.gov.au/",
    "Origin":  "https://statistics.qgso.qld.gov.au",
}

# Corrected 2026-09-16 after the first live run confirmed the format;
# "Financial Year" / "Year Ended 30 Jun YYYY" is now proven correct.
# TO_DATE is computed, not hardcoded -- a literal year here would
# silently cap this fetcher at 2025 forever once next year arrives,
# never erroring, just quietly going stale. Requesting one year past
# "now" is safe: this is a RANGE query, confirmed to just return
# whatever's actually published up to that point (QGSO updates this
# collection annually each March per its catalog listing, so the
# newest available year usually lags "now" by several months anyway --
# asking for more than exists doesn't fail, it just returns what's
# there).
FROM_DATE = "Year Ended 30 Jun 2001"
TO_DATE   = f"Year Ended 30 Jun {datetime.now().year + 1}"
PERIOD    = "Financial Year"
DATE_FMT  = "Y1"

ASSEMBLED_FILE_PATTERN = "qgso_and_bom_{year}.xlsx"
REGIONAL_DATABASE_URL = (
    "http://www.qgso.qld.gov.au/products/tables/qld-regional-database/index.php"
)
POP_SHEET_NAME_CANDIDATES = ["Pop", "Population", "Pop sheet"]


def _q(*pairs) -> list[tuple]:
    """Build interleaved p_names/p_values list from alternating (name, value)
    args -- Oracle PL/SQL WebTK reads these as parallel arrays, not a dict.
    Identical helper to fetch_qgso_housing.py's."""
    result = []
    it = iter(pairs)
    for name in it:
        value = next(it)
        result.append(("p_names", name))
        result.append(("p_values", value))
    return result


def _parse_options(html: str, select_name: str) -> list[str]:
    """Parse <option> text values from a named <select>. MUST use lxml --
    html.parser merges unclosed <OPTION> tags into one string (confirmed
    real quirk in fetch_qgso_housing.py)."""
    soup = BeautifulSoup(html, "lxml")
    for sel in soup.find_all("select", {"name": select_name}):
        opts = [o.get_text(strip=True) for o in sel.find_all("option")]
        return [o for o in opts if o]
    return []


class QGSOPopulationERPFetcher(BaseFetcher):

    SOURCE_NAME      = "population_erp"
    SUPPORTED_STATES = ["QLD"]

    def fetch_all(self):
        towns = self.applicable_towns()
        if not towns:
            self.log.info("No QLD towns configured — nothing to fetch")
            return

        sa2_map: dict[str, list] = {}
        for town in towns:
            sa2 = getattr(town, "qgso_sa2", None) or town.sa2_code
            if sa2:
                sa2_map.setdefault(str(sa2), []).append(town)

        if not sa2_map:
            self.result.add_error("ALL", "No SA2 codes available for QRSIS lookup")
            return

        try:
            data = self._fetch_live(sa2_map)
            self._write_results(towns, data, source_label="QRSIS live API")
            return
        except Exception as exc:
            self.log.warning(f"Live QRSIS query failed, falling back to manual file: {exc}", exc_info=True)

        self._fetch_manual_fallback(towns)

    # ── Live QRSIS path ──────────────────────────────────────────────────

    def _fetch_live(self, sa2_map: dict) -> dict:
        session = requests.Session()
        session.headers.update(HEADERS)

        udqctl_id = self._select_collection(session)
        if not udqctl_id:
            raise RuntimeError(f"Failed to get udqctl_id for collection {COLL_ID}")
        self.log.info(f"  udqctl_id={udqctl_id}")

        avail_series = self._get_options(
            session, "QIS1110W$UDQSER.ProcessSeries", udqctl_id, "infoser.htm", "p_new_multi"
        )
        self.log.info(f"  Available series: {avail_series}")
        if not avail_series:
            raise RuntimeError("No series available for this collection")
        series = avail_series[:1]
        self._select_series(session, udqctl_id, series)
        self.log.info(f"  Series selected: {series}")

        self._set_time_period(session, udqctl_id)

        self._select_region_type(session, udqctl_id, "SA2 - Statistical Area Level 2")

        avail_regions = self._get_options(
            session, "QIS1110W$UDQREG.ProcessRegions", udqctl_id, "inforeg.htm", "p_new_multi"
        )
        matched_regions = self._match_regions(sa2_map, avail_regions)
        if not matched_regions:
            raise RuntimeError("No SA2 regions matched in QRSIS region list")
        self._select_regions(session, udqctl_id, matched_regions)
        self.log.info(f"  Regions selected: {len(matched_regions)}")

        html = self._submit_report(session, udqctl_id)
        raw = self._parse_output_html(html)
        self.log.info(f"  Parsed data for {len(raw)} regions")

        return self._aggregate(series[0], raw, sa2_map)

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

    def _set_time_period(self, session, udqctl_id):
        data = _q("udqctl_id", udqctl_id, "coll_id", COLL_ID, "error_msg", "",
                  "date_format", DATE_FMT, "period", PERIOD,
                  "from_date", FROM_DATE, "to_date", TO_DATE)
        data.append(("p_op_mode", "Next"))
        resp = session.post(BASE_URL + "QIS1110W$UDQCTL.ProcessActions", data=data, timeout=30)
        self.log.info(f"  Time period step: HTTP {resp.status_code}, {len(resp.text)} bytes back")
        self._warn_if_qrsis_error(resp.text, "Time period step")

    def _select_region_type(self, session, udqctl_id, geo_label):
        resp1 = session.post(
            BASE_URL + "QIS1110W$REGTYP.ProcessRegType",
            data=_q("udqctl_id", udqctl_id, "op_mode", "->") + [("p_multi", geo_label)],
            timeout=30,
        )
        self.log.info(f"  Region type select ('{geo_label}'): HTTP {resp1.status_code}")
        time.sleep(0.2)
        resp2 = session.post(
            BASE_URL + "QIS1110W$REGTYP.ProcessRegType",
            data=_q("udqctl_id", udqctl_id, "op_mode", "Next"),
            timeout=30,
        )
        self.log.info(f"  Region type confirm: HTTP {resp2.status_code}, {len(resp2.text)} bytes back")
        self._warn_if_qrsis_error(resp2.text, "Region type step")

    def _warn_if_qrsis_error(self, text: str, step_name: str):
        # Confirmed real false positive (2026-09-16): standard HTML frameset
        # fallback text ("Your browser does not support this functionality...
        # upgrade your browser") contains generic words like "invalid" in a
        # completely innocuous context. Look for QRSIS's own actual error
        # markers instead of generic English words.
        if re.search(r'error_msg\s*=\s*[\'"][^\'"]+[\'"]|ORA-\d{5}|no data found|no rows returned', text, re.IGNORECASE):
            snippet = re.sub(r'\s+', ' ', text)[:500]
            self.log.warning(f"  {step_name} response looks like a genuine QRSIS/Oracle error: {snippet}")

    def _match_regions(self, sa2_map, available) -> list[str]:
        if available:
            self.log.info(f"  First available region: {available[0][:80]}")
        matched = []
        for code, towns in sa2_map.items():
            town_names = [t.name for t in towns]
            prefix = f"SA2/{code}"
            hits = [r for r in available if r.startswith(prefix)]
            if hits:
                matched.extend(hits)
                self.log.info(f"  Matched {'/'.join(town_names)}: {hits[0][:60]}")
            else:
                self.log.warning(f"  [{'/'.join(town_names)}] {prefix} not in QRSIS list")
        return matched

    def _select_regions(self, session, udqctl_id, regions):
        session.post(
            BASE_URL + "QIS1110W$UDQREG.ProcessActions",
            data=_q("udqctl_id", udqctl_id, "info_page", "inforeg.htm",
                    "error_msg", "", "op_mode", "->") + [("p_new_multi", r) for r in regions],
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
            html, re.IGNORECASE | re.DOTALL
        )
        if m:
            return m.group(1)
        m = re.search(r'udqctl_id.*?(\d{4,6})', url + html, re.DOTALL)
        if m:
            return m.group(1)
        return None

    def _parse_output_html(self, html: str) -> dict:
        result: dict[str, dict] = {}
        sections = re.split(r'Region\s*:\s*', html)
        for section in sections[1:]:
            m = re.match(r'((?:SA2|SA3|SA4|LGA)/[\w]+)', section)
            if not m:
                continue
            region_code = m.group(1)
            region_data: dict[str, dict] = {}

            soup = BeautifulSoup(section, "lxml")
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

    def _aggregate(self, series_name: str, raw: dict, sa2_map: dict) -> dict:
        result: dict[str, dict] = {}
        for code, towns in sa2_map.items():
            region_key = f"SA2/{code}"
            periods = raw.get(region_key) or raw.get(code) or {}
            if not periods:
                continue
            for town in towns:
                year_vals: dict[str, int] = {}
                for period, vals in periods.items():
                    m = re.search(r'\b(19|20)\d{2}\b', period)
                    if not m:
                        continue
                    yr = int(m.group(0))
                    if not (YEAR_START <= yr <= YEAR_END):
                        continue
                    v = vals.get(series_name)
                    if v is not None:
                        year_vals[str(yr)] = int(v)
                if year_vals:
                    result[town.slug] = year_vals
        return result

    def _write_results(self, towns, data: dict, source_label: str):
        out_dir = Path(__file__).parent.parent / "cache" / "population"
        out_dir.mkdir(parents=True, exist_ok=True)

        for town in towns:
            year_vals = data.get(town.slug)
            if not year_vals:
                self.log.warning(f"  [{town.name}] no ERP data retrieved")
                self.result.towns_failed.append(town.name)
                continue

            latest_year = max(year_vals, key=int)
            out = {
                "town":   town.name,
                "state":  town.state,
                "source": source_label,
                "year":   int(latest_year),
                "value":  year_vals[latest_year],
                "series_by_year": year_vals,
            }
            out_path = out_dir / f"{town.slug}_population_erp.json"
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2)

            self.log.info(f"  {town.name}: {latest_year} = {year_vals[latest_year]:,}")
            self.result.towns_ok.append(town.name)

    # ── Manual-file fallback (previously-working path, kept as-is) ─────────

    def _fetch_manual_fallback(self, towns):
        path = self._find_assembled_file()
        if not path:
            current_year = datetime.now().year
            self.result.add_error(
                "ALL",
                f"No manually-downloaded QGSO Regional Database export found, "
                f"and the live QRSIS query also failed (see warning above). "
                f"Manual fix: visit {REGIONAL_DATABASE_URL}, download the "
                f"Population (ERP) data for this year's towns, assemble it "
                f"into a workbook with a 'Pop' sheet, and save it as one of:\n"
                f"    cache/{ASSEMBLED_FILE_PATTERN.format(year=current_year)}\n"
                f"    cache/{ASSEMBLED_FILE_PATTERN.format(year=current_year - 1)}\n"
                f"  Then re-run this fetcher."
            )
            return

        self.log.info(f"  Found manually-assembled file: {path.name}")
        pop_data = self._parse_pop_sheet(path)
        if not pop_data:
            self.result.add_error(
                "ALL",
                f"Found {path.name} but could not locate/parse its 'Pop' sheet."
            )
            return

        for town in towns:
            self._extract_town_manual(town, pop_data)

    def _find_assembled_file(self) -> Optional[Path]:
        current_year = datetime.now().year
        for year in (current_year, current_year - 1):
            candidate = CACHE_DIR / ASSEMBLED_FILE_PATTERN.format(year=year)
            if candidate.exists():
                return candidate
        return None

    def _parse_pop_sheet(self, path: Path) -> dict:
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            sheet_name = next((n for n in wb.sheetnames if n in POP_SHEET_NAME_CANDIDATES), None)
            if sheet_name is None:
                self.log.error(f"No sheet named any of {POP_SHEET_NAME_CANDIDATES} found in {path.name}")
                return {}

            ws = wb[sheet_name]
            header_row_idx = None
            region_col = None
            for r in range(1, min(15, ws.max_row) + 1):
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and v.strip().lower() == "region":
                        header_row_idx = r
                        region_col = c
                        break
                if header_row_idx:
                    break
            if header_row_idx is None:
                self.log.error(f"Could not find a 'Region' header cell in {sheet_name}")
                return {}

            year_col = None
            year_val = None
            for c in range(1, ws.max_column + 1):
                v = ws.cell(header_row_idx, c).value
                if isinstance(v, (int, float)) and 1990 <= v <= 2100:
                    year_col, year_val = c, int(v)
                    break
                if isinstance(v, str) and v.strip().isdigit() and 1990 <= int(v.strip()) <= 2100:
                    year_col, year_val = c, int(v.strip())
                    break
            if year_col is None:
                self.log.error(f"Could not find a year column in {sheet_name}'s header row")
                return {}

            result = {}
            for r in range(header_row_idx + 1, ws.max_row + 1):
                region = ws.cell(r, region_col).value
                value = ws.cell(r, year_col).value
                if region and isinstance(value, (int, float)):
                    result[str(region).strip()] = {"value": int(value), "year": year_val}
            self.log.info(f"  Parsed {len(result)} region entries for year {year_val}")
            return result
        except Exception as exc:
            self.log.error(f"Pop sheet parse error: {exc}", exc_info=True)
            return {}

    def _extract_town_manual(self, town, pop_data: dict):
        candidates = [town.sa2_name, town.lga, town.name]
        match = None
        matched_as = None
        for candidate in candidates:
            if not candidate:
                continue
            if candidate in pop_data:
                match, matched_as = pop_data[candidate], candidate
                break
            for key in pop_data:
                if key.lower() == candidate.lower():
                    match, matched_as = pop_data[key], key
                    break
            if match:
                break

        if not match:
            self.log.warning(f"  [{town.name}] no Population (ERP) match (tried: {[c for c in candidates if c]}) -- skipping")
            self.result.towns_skipped.append(town.name)
            return

        out = {
            "town": town.name, "state": town.state,
            "source": "QGSO Regional Database (manually assembled export)",
            "matched_as": matched_as, "year": match["year"], "value": match["value"],
        }
        out_dir = Path(__file__).parent.parent / "cache" / "population"
        out_dir.mkdir(parents=True, exist_ok=True)
        with open(out_dir / f"{town.slug}_population_erp.json", "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)
        self.log.info(f"  {town.name} ('{matched_as}'): {match['year']} = {match['value']:,}")
        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = QGSOPopulationERPFetcher().run()
    sys.exit(0 if result.success else 1)