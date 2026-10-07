"""
regional-indicators/fetchers/fetch_population_erp.py

Fetches SA2-level estimated resident population (ERP): the "Population
(ERP)" row in each town's block of the Population sheet. Town (UCL)
figures come from fetch_population_ucl.py, non-resident workers from
fetch_population_nrw.py and LGA figures from fetch_population_erp_lga.py.

Source: QGSO Queensland Regional Database (QRSIS), queried through its
web wizard.
    collgrp_id = "1"     Population group
    coll_id    = "1961"  "Population (ERP)(a) persons only", 1991 to
                         the latest year, ASGS 2021 boundaries.
                         Collections on ASGS 2016 (coll_id 1298) and
                         ASGS 2011 boundaries also exist and are not
                         used.

The wizard is driven in the same way as in fetch_qgso_housing.py:
interleaved p_names/p_values POST encoding (_q), lxml parsing for
unclosed <OPTION> tags, udqctl_id taken from the redirect, and the same
six steps (collection, series, time period, region type, regions,
submit). The shared logic is duplicated in the two fetchers and has not
been factored into a common module.

The series is annual (financial year). The latest available year is
read from the wizard's time-period page.

Fallback: if the live query fails, a manually assembled export at
cache/qgso_and_bom_<year>.xlsx is read if present. See
docs/manual_processes.md.

Queensland only (SUPPORTED_STATES = ["QLD"]). SA2 figures for towns in
other states require a separate ABS-based fetcher.

Output: cache/population/<slug>_population_erp.json
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


# ── QRSIS constants (as in fetch_qgso_housing.py) ──────────────────────────

BASE_URL     = "https://statistics.qgso.qld.gov.au/pls/qis_public/"
PUBLIC_USER  = "edtert"
ACCESS_LEVEL = "85"

# Collection identifiers, as given in the wizard's HTML
COLLGRP_ID = "1"    # Population group
COLL_ID    = "1961" # "Population (ERP)(a) persons only", from 1991, ASGS 2021

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": "https://statistics.qgso.qld.gov.au/",
    "Origin":  "https://statistics.qgso.qld.gov.au",
}

# Time-period fields, as the wizard submits them:
#   - date_format: the hidden field value for this collection is "Y2".
#   - from_date/to_date: the <select> offers plain year numbers
#     ("2025", "2024", ... "1991"). Any other form (for example
#     "Year Ended 30 Jun YYYY") invalidates the session without an
#     error, and the region list then comes back empty.
# TO_DATE is a default; the newest year offered is read from the page
# (see _discover_max_to_date).
FROM_DATE = "2001"
TO_DATE   = str(datetime.now().year + 1)
PERIOD    = "Financial Year"
DATE_FMT  = "Y2"

ASSEMBLED_FILE_PATTERN = "qgso_and_bom_{year}.xlsx"
REGIONAL_DATABASE_URL = (
    "http://www.qgso.qld.gov.au/products/tables/qld-regional-database/index.php"
)
POP_SHEET_NAME_CANDIDATES = ["Pop", "Population", "Pop sheet"]


def _q(*pairs) -> list[tuple]:
    """Build the interleaved p_names/p_values list from alternating
    (name, value) arguments. The Oracle PL/SQL Web Toolkit reads these as
    parallel arrays in submission order, so a dict cannot be used.
    """
    result = []
    it = iter(pairs)
    for name in it:
        value = next(it)
        result.append(("p_names", name))
        result.append(("p_values", value))
    return result


def _parse_options(html: str, select_name: str) -> list[str]:
    """Return the <option> text values of a named <select>.

    Uses the lxml parser: QRSIS serves unclosed <OPTION> tags, which
    html.parser merges into a single string.
    """
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
        time_periods_html = self._select_series(session, udqctl_id, series)
        self.log.info(f"  Series selected: {series}")

        self._set_time_period(session, udqctl_id, time_periods_html)

        self._select_region_type(session, udqctl_id, "SA2 - Statistical Area Level 2")

        avail_regions = self._get_options(
            session, "QIS1110W$UDQREG.ProcessRegions", udqctl_id, "inforeg.htm", "p_new_multi"
        )
        matched_regions, sa2_name_by_code = self._match_regions(sa2_map, avail_regions)
        if not matched_regions:
            raise RuntimeError("No SA2 regions matched in QRSIS region list")
        # Log the SA2 names read from the region list. They differ from
        # town.name for several towns (Dalby, Dysart, the Toowoomba
        # sub-areas) and determine which workbook block is written.
        self.log.info(f"  sa2_name_by_code captured: {sa2_name_by_code}")
        self._select_regions(session, udqctl_id, matched_regions)
        self.log.info(f"  Regions selected: {len(matched_regions)}")

        html = self._submit_report(session, udqctl_id)
        raw = self._parse_output_html(html)
        self.log.info(f"  Parsed data for {len(raw)} regions")

        return self._aggregate(series[0], raw, sa2_map, sa2_name_by_code)

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
        # The response to this POST is the next wizard page (Time Periods);
        # it is returned so that the available date range can be read
        # from it.
        return resp.text

    def _discover_max_to_date(self, time_periods_html: str) -> str | None:
        """Return the newest option of the "To Date" dropdown on the Time
        Periods page.

        The dropdown is a <select> of plain year numbers, most recent first,
        with the newest marked selected. Requesting a year that QRSIS does
        not hold fails the whole session, so the available year is read
        from the page and not computed.

        From Date and To Date both use <select name="p_values">; they are
        distinguished only by the preceding hidden
        <input name="p_names" value="to_date">. From Date comes first and
        defaults to the oldest year, so the select is identified by that
        marker and not by position.
        """
        soup = BeautifulSoup(time_periods_html, "lxml")

        to_date_marker = soup.find(
            "input", {"name": "p_names", "value": "to_date"}
        )
        if not to_date_marker:
            return None

        select = to_date_marker.find_next("select")
        if not select:
            return None

        options = select.find_all("option")
        years = [o.get_text(strip=True) for o in options if o.get_text(strip=True).isdigit()]
        if not years:
            return None

        selected = select.find("option", selected=True)
        if selected and selected.get_text(strip=True).isdigit():
            return selected.get_text(strip=True)
        return max(years, key=int)

    def _set_time_period(self, session, udqctl_id, time_periods_html: str = ""):
        to_date = self._discover_max_to_date(time_periods_html) if time_periods_html else None
        if to_date:
            self.log.info(f"  Latest To Date option on the page: {to_date}")
        else:
            to_date = TO_DATE
            self.log.warning(
                f"  Could not read the To Date option from the page -- "
                f"falling back to computed guess {to_date!r}, which may not "
                f"be a valid option and could break this step silently."
            )

        data = _q("udqctl_id", udqctl_id, "coll_id", COLL_ID, "error_msg", "",
                  "date_format", DATE_FMT, "period", PERIOD,
                  "from_date", FROM_DATE, "to_date", to_date)
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
        # The frameset fallback text served to every client ("Your browser
        # does not support this functionality...") contains words such as
        # "invalid", so errors are detected by QRSIS's own error markers
        # and not by generic words.
        if re.search(r'error_msg\s*=\s*[\'"][^\'"]+[\'"]|ORA-\d{5}|no data found|no rows returned', text, re.IGNORECASE):
            snippet = re.sub(r'\s+', ' ', text)[:500]
            self.log.warning(f"  {step_name} response looks like a genuine QRSIS/Oracle error: {snippet}")

    def _match_regions(self, sa2_map, available) -> tuple[list[str], dict[str, str]]:
        """Return (matched region strings, {sa2_code: display_name}).

        The display name is parsed from the full region string, e.g.
        "SA2/317011456 - Toowoomba - Central (01/07/2011 - 30/06/2026)"
        -> "Toowoomba - Central". Writes target the SA2 label, which
        differs from town.name for several towns (each of the three
        Toowoomba sub-areas has its own SA2 name).
        """
        if available:
            self.log.info(f"  First available region: {available[0][:80]}")
        matched = []
        sa2_name_by_code: dict[str, str] = {}
        for code, towns in sa2_map.items():
            town_names = [t.name for t in towns]
            prefix = f"SA2/{code}"
            hits = [r for r in available if r.startswith(prefix)]
            if hits:
                matched.extend(hits)
                self.log.info(f"  Matched {'/'.join(town_names)}: {hits[0][:60]}")
                m = re.match(r'^SA2/\d+ - (.+?) \(', hits[0])
                if m:
                    sa2_name_by_code[code] = m.group(1).strip()
                else:
                    self.log.warning(f"  Could not parse SA2 name out of: {hits[0][:80]}")
            else:
                self.log.warning(f"  [{'/'.join(town_names)}] {prefix} not in QRSIS list")
        return matched, sa2_name_by_code

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
                    # The wizard's default display style, which the
                    # output parser expects.
                    "display_style", "For each Series display Time Period by Region",
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
        """Return {region_code: {year: value}}.

        Output layout: one table headed "Period" in its first column, with
        region labels (e.g. "SA2/307011176 - Roma",
        "LGA/33610 - Goondiwindi (R)") as the remaining column headers and
        one row per year.
        """
        result: dict[str, dict] = {}
        soup = BeautifulSoup(html, "lxml")

        for table in soup.find_all("table"):
            rows = table.find_all("tr")
            if not rows:
                continue
            header_cells = rows[0].find_all(["th", "td"])
            header_texts = [c.get_text(strip=True) for c in header_cells]
            if not header_texts or header_texts[0].lower() != "period":
                continue  # not the data table

            region_codes = []
            for h in header_texts[1:]:
                m = re.match(r'(SA2|LGA|SA3|SA4)/(\S+?)\s*-', h)
                region_codes.append(m.group(2) if m else None)

            for row in rows[1:]:
                cells = row.find_all(["td", "th"])
                if not cells:
                    continue
                period_text = cells[0].get_text(strip=True)
                m = re.search(r'\b(19|20)\d{2}\b', period_text)
                if not m:
                    continue
                year = int(m.group(0))

                for i, code in enumerate(region_codes, start=1):
                    if code is None or i >= len(cells):
                        continue
                    raw_val = cells[i].get_text(strip=True).replace(",", "").replace("$", "").strip()
                    try:
                        value = float(raw_val)
                    except ValueError:
                        continue
                    result.setdefault(code, {})[year] = value

        return result

    def _aggregate(self, series_name: str, raw: dict, sa2_map: dict, sa2_name_by_code: dict) -> dict:
        """Return {town.slug: {"year_vals": {...}, "sa2_name": "..."}}.

        sa2_name is carried with the values so that the workbook writer can
        target the SA2 label (see _match_regions).

        raw is {region_code: {year_int: value}}; the collection has a
        single series ("Persons (Persons)"). series_name is retained in the
        signature for logging and is not used for lookup.
        """
        result: dict[str, dict] = {}
        for code, towns in sa2_map.items():
            year_vals_raw = raw.get(code) or {}
            if not year_vals_raw:
                continue
            sa2_name = sa2_name_by_code.get(code)
            for town in towns:
                year_vals: dict[str, int] = {
                    str(yr): int(v)
                    for yr, v in year_vals_raw.items()
                    if YEAR_START <= yr <= YEAR_END
                }
                if year_vals:
                    result[town.slug] = {"year_vals": year_vals, "sa2_name": sa2_name}
        return result

    def _write_results(self, towns, data: dict, source_label: str):
        out_dir = Path(__file__).parent.parent / "cache" / "population"
        out_dir.mkdir(parents=True, exist_ok=True)

        for town in towns:
            entry = data.get(town.slug)
            if not entry or not entry.get("year_vals"):
                self.log.warning(f"  [{town.name}] no ERP data retrieved")
                self.result.towns_failed.append(town.name)
                continue

            year_vals = entry["year_vals"]
            # town.sa2_name from towns.toml takes precedence when present.
            # The name parsed from the region list (entry["sa2_name"]) is
            # the fallback for a town with no sa2_name configured.
            sa2_name = town.sa2_name or entry.get("sa2_name") or town.name
            if not town.sa2_name and entry.get("sa2_name") and entry["sa2_name"] != town.name:
                self.log.info(
                    f"  [{town.name}] using live-discovered SA2 name "
                    f"'{entry['sa2_name']}' -- consider adding this to "
                    f"towns.toml as sa2_name for reliability."
                )
            latest_year = max(year_vals, key=int)
            out = {
                "town":     town.name,
                "state":    town.state,
                "source":   source_label,
                "sa2_name": sa2_name,
                "year":     int(latest_year),
                "value":    year_vals[latest_year],
                "series_by_year": year_vals,
            }
            out_path = out_dir / f"{town.slug}_population_erp.json"
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2)

            self.log.info(f"  {town.name}: {latest_year} = {year_vals[latest_year]:,}")
            self.result.towns_ok.append(town.name)

    # ── Manual-file fallback ───────────────────────────────────────────────

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