"""
fetchers/fetch_bom_rainfall.py
--------------------------------
Fetches annual rainfall totals from the Bureau of Meteorology (BOM)
Climate Data Online for each town's weather station.

Source: BOM Climate Data Online — Annual climate statistics
URL pattern: http://www.bom.gov.au/jsp/ncc/cdio/weatherData/av?
             p_nccObsCode=139&p_display_type=dataFile&p_startYear=&
             p_c=-17626285&p_stn_num={station}

Station numbers are stored in towns.toml as bom_station.

The data file is a CSV with one row per year containing monthly totals
and an annual total column.

Coverage: All towns with a bom_station configured. Currently:
  QLD: Roma, Chinchilla, Dalby, Miles, Tara, Wandoan, Wallumbilla,
       Goondiwindi, Moranbah, Dysart, Toowoomba
  NSW: Narrabri

Website CSV produced:
  Environment - total rainfall.csv
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
from config import CACHE_DIR

try:
    import requests
except ImportError:
    raise ImportError("pip install requests")


# ── Configuration ──────────────────────────────────────────────────────────────

# BOM Climate Data Online — Annual rainfall data file
# p_nccObsCode=139 = annual rainfall
BOM_URL_TEMPLATE = (
    "http://www.bom.gov.au/jsp/ncc/cdio/weatherData/av"
    "?p_nccObsCode=139&p_display_type=dataFile&p_startYear="
    "&p_c=-17626285&p_stn_num={station}"
)

# Alternatively — the all-years annual summary CSV
# This URL gives a single CSV with all years for one station
BOM_ALL_YEARS_TEMPLATE = (
    "http://www.bom.gov.au/jsp/ncc/cdio/weatherData/av"
    "?p_nccObsCode=139&p_display_type=dataFile&p_startYear="
    "&p_stn_num={station}"
)


class BOMRainfallFetcher(BaseFetcher):

    SOURCE_NAME      = "bom_rainfall"
    SUPPORTED_STATES = []   # covers any town with bom_station set

    def fetch_all(self):
        for town in self.applicable_towns():
            if not town.bom_station:
                self.log.warning(f"  [{town.name}] no bom_station in towns.toml — skipping")
                self.result.towns_skipped.append(town.name)
                continue
            self._fetch_town(town)

    def _fetch_town(self, town):
        station = town.bom_station
        url     = BOM_URL_TEMPLATE.format(station=station)
        key     = f"bom_rainfall_{station}"

        path = self._download_bom(url, key)
        if not path:
            url2 = BOM_ALL_YEARS_TEMPLATE.format(station=station)
            path = self._download_bom(url2, key)

        if not path:
            self.log.warning(f"  [{town.name}] BOM download failed for station {station}")
            self.result.towns_failed.append(town.name)
            return

        annual = self._parse_bom_csv(path, town.name, station)
        if not annual:
            self.log.warning(f"  [{town.name}] No rainfall data parsed from {path.name}")
            self.result.towns_failed.append(town.name)
            return

        self._write_cache(town, station, annual)

    def _download_bom(self, url: str, cache_key: str) -> Path | None:
        """
        Download BOM Climate Data Online file with browser User-Agent.
        BOM returns 403 to default requests/urllib user-agents.
        """
        from config import CACHE_DIR
        out_path = CACHE_DIR / f"{cache_key}.csv"
        if out_path.exists() and not getattr(self, '_force', False):
            return out_path
        try:
            headers = {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                ),
                "Referer": "http://www.bom.gov.au/climate/data/",
                "Accept":  "text/html,application/xhtml+xml,*/*",
            }
            self.log.info(f"  Downloading {url.split('?')[0]}?... → {out_path.name}")
            resp = requests.get(url, headers=headers, timeout=30)
            resp.raise_for_status()

            # Check we got CSV not HTML
            if b"<html" in resp.content[:200].lower():
                self.log.warning(f"  Got HTML (not CSV) from BOM — URL may be wrong")
                return None

            with open(out_path, "wb") as f:
                f.write(resp.content)
            self.log.info(f"  Saved {out_path.stat().st_size // 1024} KB")
            return out_path
        except Exception as exc:
            self.log.warning(f"  BOM download failed: {exc}")
            return None

    def _parse_bom_csv(self, path: Path, town_name: str, station: str) -> dict:
        """
        Parse BOM annual rainfall CSV.

        BOM file format (typical):
          Row 1: station header info
          Row 2: column headers: Year, Jan, Feb, ..., Dec, Annual, ...
          Row 3+: one row per year

        Returns { year_int: annual_mm }
        """
        try:
            # Try UTF-8 first, then latin-1
            for enc in ["utf-8", "latin-1"]:
                try:
                    with open(path, encoding=enc) as f:
                        content = f.read()
                    break
                except UnicodeDecodeError:
                    continue

            lines = content.splitlines()

            # Find header row containing 'Year' and 'Annual'
            header_idx = None
            for i, line in enumerate(lines):
                if "Year" in line and "Annual" in line.upper():
                    header_idx = i
                    break

            if header_idx is None:
                # Try finding it differently — some BOM files have varied formats
                for i, line in enumerate(lines):
                    if re.search(r'\bYear\b', line, re.IGNORECASE):
                        header_idx = i
                        break

            if header_idx is None:
                self.log.error(f"  Could not find header row in BOM CSV for {town_name}")
                return {}

            reader = csv.reader(lines[header_idx:])
            headers = [h.strip().lower() for h in next(reader)]

            # Find year column and annual column
            year_col   = next((i for i, h in enumerate(headers) if h == "year"), None)
            annual_col = next((i for i, h in enumerate(headers) if "annual" in h), None)

            if year_col is None or annual_col is None:
                self.log.error(
                    f"  Could not find Year/Annual columns in {path.name}. "
                    f"Headers: {headers[:8]}"
                )
                return {}

            result = {}
            for row in reader:
                if not row or len(row) <= annual_col:
                    continue
                yr_raw = row[year_col].strip()
                if not yr_raw.isdigit():
                    continue
                year = int(yr_raw)
                val  = row[annual_col].strip()
                if val and val not in ("", "-", "N/A"):
                    try:
                        result[year] = round(float(val), 1)
                    except ValueError:
                        pass

            if result:
                yrs = sorted(result)
                self.log.info(
                    f"  {town_name} (station {station}): {len(result)} years, "
                    f"{yrs[0]}-{yrs[-1]}, latest = {result[yrs[-1]]} mm"
                )

            return result

        except Exception as exc:
            self.log.error(f"BOM CSV parse error for {town_name}: {exc}", exc_info=True)
            return {}

    def _write_cache(self, town, station: str, annual: dict):
        from config import YEAR_START, YEAR_END
        values = {
            str(yr): val for yr, val in annual.items()
            if YEAR_START <= yr <= YEAR_END
        }

        if not values:
            self.log.warning(f"  [{town.name}] no rainfall values in range {YEAR_START}-{YEAR_END}")
            self.result.towns_failed.append(town.name)
            return

        out = {
            "town":        town.name,
            "state":       town.state,
            "bom_station": station,
            "source":      "Bureau of Meteorology — Climate Data Online",
            "source_url":  BOM_URL_TEMPLATE.format(station=station),
            "note":        "Annual total rainfall in mm",
            "indicators":  {"rainfall": values},
        }

        out_dir  = CACHE_DIR / "rainfall"
        out_dir.mkdir(exist_ok=True)
        out_path = out_dir / f"{town.slug}_bom_rainfall.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = BOMRainfallFetcher().run()
    sys.exit(0 if result.success else 1)