"""
regional-indicators/fetchers/fetch_population_ucl.py

Fetches estimated resident population (ERP) by Urban Centre and Locality
(UCL) for Queensland towns, 2001 to the latest release.

Source: Queensland Government Statistician's Office (QGSO)
    https://www.qgso.qld.gov.au/issues/<n>/
        estimated-resident-population-urban-centre-locality-qld-2001-<year>p.csv

The issue number <n> and the year in the file name change with each
release. The configured URL is tried first; if it fails, the QGSO
regional population page is searched for the current link.

File layout:
    Row 1   title
    Row 4   year headers: "2001", "2002", ... "2025p"
            (suffix p = provisional, r = revised)
    Row 5   "- persons -" label
    Row 6+  UCL name | values, comma-formatted (e.g. "6,310")

Small localities are published with an "(L)" suffix ("Tara (L)",
"Wallumbilla (L)", "Wandoan (L)"); the suffix is ignored when matching
town names.

Encoding: latin-1 (the file contains the Windows-1252 em dash, 0x97).

Output: cache/population/<slug>_population_ucl.json
Website CSV: "Population - town.csv" (Queensland towns only)
"""

from __future__ import annotations

import csv
import io
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher

try:
    import requests
except ImportError:
    raise ImportError("pip install requests")


# ── Configuration ──────────────────────────────────────────────────────────────

# Update each cycle: the issue number and end year change with each release.
# Pattern: https://www.qgso.qld.gov.au/issues/{N}/estimated-resident-population-urban-centre-locality-qld-2001-{year}p.csv
UCL_URL = (
    "https://www.qgso.qld.gov.au/issues/5496/"
    "estimated-resident-population-urban-centre-locality-qld-2001-2025p.csv"
)

# QGSO page listing the current release; searched when UCL_URL fails.
QGSO_STATS_URL = (
    "https://www.qgso.qld.gov.au/statistics/theme/population/"
    "population-estimates/regions"
)

CACHE_KEY = "qgso_ucl_erp"
FILE_ENC  = "latin-1"   # the file contains the Windows-1252 em dash (0x97)


class QGSOPopulationUCLFetcher(BaseFetcher):

    SOURCE_NAME      = "population_ucl"
    SUPPORTED_STATES = ["QLD"]

    def fetch_all(self):
        # ── Configured URL ────────────────────────────────────────────────────
        path = self.download(UCL_URL, CACHE_KEY, suffix=".csv")

        # ── Fallback: find the current URL on the QGSO page ───────────────────
        if not path:
            self.log.info("  Direct URL failed — scraping QGSO page for current link")
            url = self._find_current_url()
            if url:
                path = self.download(url, CACHE_KEY, suffix=".csv")

        if not path:
            self.result.add_error(
                "ALL",
                "Could not download UCL population file.\n"
                "  Manual fix: download the file from\n"
                "  https://www.qgso.qld.gov.au/statistics/theme/population/"
                "population-estimates/regions\n"
                f"  and save it as:  cache/{CACHE_KEY}.csv"
            )
            return

        self.log.info(f"  Parsing {path.name} ({path.stat().st_size // 1024} KB)")

        ucl_data = self._parse_ucl_csv(path)
        if not ucl_data:
            self.result.add_error("ALL", "UCL CSV parse returned no data")
            return

        for town in self.applicable_towns():
            self._extract_town(town, ucl_data)

    def _find_current_url(self) -> str | None:
        """Return the current CSV download URL from the QGSO statistics page.

        The page uses relative hrefs ("/issues/<n>/...csv"), so both patterns
        match the relative form and the host is prepended to the result.
        """
        try:
            resp = requests.get(QGSO_STATS_URL, timeout=30)
            resp.raise_for_status()
            # CSV link matching the UCL population file name
            pattern = r'/[^\s"\']*urban-centre-locality-qld[^\s"\']*\.csv'
            paths = re.findall(pattern, resp.text)
            if paths:
                url = f"https://www.qgso.qld.gov.au{paths[0]}"
                self.log.info(f"  Found URL on page: {url}")
                return url
            # Any CSV under /issues/ with "ucl" in its name
            issue_pat = r'/issues/\d+/[^\s"\']*ucl[^\s"\']*\.csv'
            paths = re.findall(issue_pat, resp.text, re.IGNORECASE)
            if paths:
                url = f"https://www.qgso.qld.gov.au{paths[0]}"
                self.log.info(f"  Found URL on page: {url}")
                return url
        except Exception as exc:
            self.log.warning(f"  Page scrape failed: {exc}")
        return None

    # ── Parser ─────────────────────────────────────────────────────────────────

    def _parse_ucl_csv(self, path: Path) -> dict:
        """Parse the UCL CSV and return, for example:
          { "Roma": {"2001": 6310, ..., "2025": 6757} }

        Year keys have the trailing "p" (provisional) or "r" (revised)
        removed. Values have thousands separators removed.
        """
        try:
            with open(path, encoding=FILE_ENC) as f:
                rows = list(csv.reader(f))

            # The header row is the first row with years in columns 1 onwards
            header_idx = None
            for i, row in enumerate(rows):
                if len(row) > 1 and str(row[1]).strip().rstrip('rp').isdigit():
                    header_idx = i
                    break

            if header_idx is None:
                self.log.error("Could not find year header row in UCL CSV")
                return {}

            header = rows[header_idx]
            year_map = {}  # column index -> year string
            for i, v in enumerate(header[1:], start=1):
                yr = str(v).strip().rstrip('rp')
                if yr.isdigit() and 1990 <= int(yr) <= 2030:
                    year_map[i] = yr

            years = sorted(year_map.values(), key=int)
            self.log.info(
                f"  Year range: {years[0]} to {years[-1]} ({len(years)} years)"
            )

            result = {}
            for row in rows[header_idx + 1:]:
                if not row or not row[0].strip():
                    continue
                # Skip the "— persons —" label row
                if row[0].strip().startswith('\x97') or '— persons' in row[0]:
                    continue

                ucl_raw   = row[0].strip()
                ucl_clean = ucl_raw.replace('(L)', '').replace('(l)', '').strip()

                year_vals = {}
                for col_i, yr in year_map.items():
                    if col_i >= len(row):
                        continue
                    raw_val = row[col_i].replace(',', '').strip()
                    if raw_val and raw_val.lstrip('-').isdigit():
                        year_vals[yr] = int(raw_val)

                if year_vals:
                    result[ucl_clean] = year_vals
                    if ucl_raw != ucl_clean:
                        result[ucl_raw] = year_vals   # also keyed with the "(L)" suffix

            self.log.info(f"  Parsed {len(result)} UCL entries")
            return result

        except Exception as exc:
            self.log.error(f"UCL CSV parse error: {exc}", exc_info=True)
            return {}

    # ── Per-town extraction ────────────────────────────────────────────────────

    def _extract_town(self, town, ucl_data: dict):
        """Match the town name to a UCL entry and write its cache JSON."""
        candidates = [town.name, town.sa2_name, town.name.title()]

        year_vals  = None
        matched_as = None
        for candidate in candidates:
            if candidate in ucl_data:
                year_vals  = ucl_data[candidate]
                matched_as = candidate
                break
            for key in ucl_data:
                if key.lower() == candidate.lower():
                    year_vals  = ucl_data[key]
                    matched_as = key
                    break
            if year_vals:
                break

        if not year_vals:
            self.log.warning(
                f"  [{town.name}] no UCL match (tried: {candidates}) — skipping"
            )
            self.result.towns_skipped.append(town.name)
            return

        latest_yr  = max(year_vals, key=int)
        latest_val = year_vals[latest_yr]

        out = {
            "town":               town.name,
            "state":              town.state,
            "sa2_code":           town.sa2_code,
            "source":             "QGSO Estimated Resident Population by UCL",
            "source_url":         UCL_URL,
            "ucl_name":           matched_as,
            "population_by_year": year_vals,
        }

        out_dir  = Path(__file__).parent.parent / "cache" / "population"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{town.slug}_population_ucl.json"

        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        self.log.info(
            f"  {town.name} ('{matched_as}'): "
            f"{len(year_vals)} years, "
            f"latest ({latest_yr}) = {latest_val:,}"
        )
        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = QGSOPopulationUCLFetcher().run()
    sys.exit(0 if result.success else 1)