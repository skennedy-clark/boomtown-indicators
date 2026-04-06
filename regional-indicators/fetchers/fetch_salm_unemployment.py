"""
fetchers/fetch_salm_unemployment.py
-------------------------------------
Fetches Small Area Labour Markets (SALM) unemployment rates from the
Department of Employment and Workplace Relations (DEWR).

Source: DEWR Small Area Labour Markets — Smoothed SA2 datafiles (ASGS 2021)
Source page: https://www.dewr.gov.au/employment-research/small-area-labour-markets
Published: Quarterly (Mar/Jun/Sep/Dec), ~3 months after reference quarter

The CSV contains smoothed unemployment rates for all Australian SA2s and LGAs
from December quarter 2010 onwards. "Smoothed" = 4-quarter rolling average,
which is what we want for annual indicators.

Annual value = mean of the 4 quarterly Dec/Mar/Jun/Sep values in a calendar year.

SA2 matching: uses towns.toml sa2_code (ASGS 2021 Edition 3).
Coverage: NATIONAL — covers QLD, NSW (Narrabri), and VIC (Shepparton, Yarram).

URL strategy:
  The download URL contains numeric IDs that change each quarterly release.
  The fetcher scrapes the stable resource page to find the current CSV link.
  Fallback: a known-good URL from the previous release is tried first.

Data source page:
  https://www.dewr.gov.au/employment-research/resources/salm-smoothed-sa2-datafiles-asgs-2021

Website CSV produced:
  Unemployment rate.csv
"""

from __future__ import annotations

import csv
import io
import json
import re
import statistics
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

# Stable resource page — scrape this to find current CSV URL
# Main SALM page (accessible) and resource page (may be blocked)
SALM_MAIN_PAGE = "https://www.dewr.gov.au/employment-research/small-area-labour-markets"
RESOURCE_PAGE  = (
    "https://www.dewr.gov.au/employment-research/resources/"
    "salm-smoothed-sa2-datafiles-asgs-2021"
)

# Known-good URL from December quarter 2025 release (try first, fallback if 404)
FALLBACK_CSV_URL = (
    "https://www.dewr.gov.au/download/17068/"
    "salm-smoothed-sa2-datafiles-asgs-2021-december-quarter-2025/"
    "42403/salm-smoothed-sa2-datafiles-asgs-2021-december-quarter-2025/csv"
)

CACHE_KEY = "salm_smoothed_sa2"

# CSV column format: "Data Item", "SA2 name", "SA2 Code (2021 ASGS)", "Mar-10", "Jun-10", ...
# Rows alternate between unemployment level and unemployment rate
RATE_ROW_LABEL = "Smoothed unemployment rate"


class SALMUnemploymentFetcher(BaseFetcher):

    SOURCE_NAME      = "salm_unemployment"
    SUPPORTED_STATES = []   # national

    def fetch_all(self):
        # ── Find and download the CSV ─────────────────────────────────────────
        url  = self._find_csv_url()
        path = self._download_with_browser_ua(url, CACHE_KEY)

        if not path:
            self.result.add_error("ALL", "Could not download SALM SA2 CSV")
            return

        self.log.info(f"  Parsing {path.name} ({path.stat().st_size // 1024} KB)")

        data = self._parse_csv(path)
        if not data:
            self.result.add_error("ALL", "SALM CSV parse returned no data")
            return

        for town in self.applicable_towns():
            self._extract_town(town, data)

    def _download_with_browser_ua(self, url: str, cache_key: str) -> Path | None:
        """Download with browser User-Agent to bypass bot detection."""
        from config import CACHE_DIR
        out_path = CACHE_DIR / f"{cache_key}.csv"
        if out_path.exists() and not getattr(self, '_force', False):
            self.log.info(f"  Using cached: {out_path.name}")
            return out_path
        try:
            headers = {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                )
            }
            self.log.info(f"  Downloading {url} → {out_path.name}")
            resp = requests.get(url, headers=headers, timeout=60, stream=True)
            resp.raise_for_status()
            content_type = resp.headers.get("Content-Type", "")
            if "html" in content_type.lower():
                self.log.error(f"  Got HTML response (not CSV) — URL may be wrong")
                return None
            with open(out_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=65536):
                    f.write(chunk)
            size_kb = out_path.stat().st_size // 1024
            self.log.info(f"  Saved {size_kb} KB")
            if size_kb < 500:
                self.log.warning(f"  File is small ({size_kb} KB) — expected ~2000+ KB. May be an error page.")
            return out_path
        except Exception as exc:
            self.log.error(f"  Download failed: {exc}")
            return None

    def _find_csv_url(self) -> str:
        """
        Scrape resource page to find current CSV download URL.
        Falls back to the known-good URL if the page can't be reached.
        """
        try:
            headers = {"User-Agent": "Mozilla/5.0 (research pipeline; contact uq.edu.au)"}
            resp = requests.get(RESOURCE_PAGE, headers=headers, timeout=20)
            resp.raise_for_status()

            # Find CSV download link — pattern: /download/{id}/salm-smoothed-sa2...
            pattern = r'(https://www\.dewr\.gov\.au/download/\d+/salm-smoothed-sa2[^"\'>\s]+/csv)'
            matches = re.findall(pattern, resp.text)
            if matches:
                url = matches[0]
                self.log.info(f"  Found CSV URL on resource page: {url}")
                return url

            self.log.warning("  Could not find CSV link on resource page — using fallback URL")
        except Exception as exc:
            self.log.warning(f"  Resource page scrape failed ({exc}) — using fallback URL")

        return FALLBACK_CSV_URL

    def _parse_csv(self, path: Path) -> dict:
        """
        Parse SALM SA2 CSV into:
          { sa2_code: { year: annual_rate } }

        Annual rate = mean of 4 quarterly smoothed rates in that calendar year.
        Quarter labels: "Mar-10", "Jun-10", "Sep-10", "Dec-10", ...
        """
        try:
            with open(path, encoding="utf-8-sig", errors="replace") as f:
                lines = f.readlines()

            # Find the header row (contains "SA2 Code")
            header_idx = None
            for i, line in enumerate(lines):
                if "SA2 Code" in line or "SA2 name" in line.lower() or "Data Item" in line:
                    header_idx = i
                    break

            if header_idx is None:
                self.log.error("Could not find header row in SALM CSV")
                return {}

            reader = csv.reader(lines[header_idx:])
            headers = next(reader)

            # Map column index → (year, quarter_num)
            # Format: "Mar-10" = March 2010, "Dec-24" = December 2024
            month_to_q = {"Mar": 1, "Jun": 2, "Sep": 3, "Dec": 4}
            col_to_yrq = {}
            for i, h in enumerate(headers[3:], start=3):
                m = re.match(r'^(Mar|Jun|Sep|Dec)-(\d{2})$', h.strip())
                if m:
                    mon, yr2 = m.group(1), int(m.group(2))
                    year = 2000 + yr2 if yr2 <= 50 else 1900 + yr2
                    col_to_yrq[i] = (year, month_to_q[mon])

            result: dict[str, dict[int, list]] = {}

            for row in reader:
                if len(row) < 3:
                    continue
                data_item = row[0].strip()
                if RATE_ROW_LABEL not in data_item:
                    continue

                sa2_code = str(row[2]).strip().replace(".0", "")
                if not sa2_code.isdigit():
                    continue

                if sa2_code not in result:
                    result[sa2_code] = {}

                for col_i, (year, q) in col_to_yrq.items():
                    if col_i >= len(row):
                        continue
                    val = row[col_i].strip().replace(",", "")
                    if val and val != "-":
                        try:
                            if year not in result[sa2_code]:
                                result[sa2_code][year] = []
                            result[sa2_code][year].append(float(val))
                        except ValueError:
                            pass

            # Compute annual mean from quarterly values
            annual: dict[str, dict[int, float]] = {}
            for sa2, yr_vals in result.items():
                annual[sa2] = {}
                for year, vals in yr_vals.items():
                    if vals:
                        annual[sa2][year] = round(statistics.mean(vals), 4)

            self.log.info(
                f"  Parsed {len(annual)} SA2s, "
                f"years {min(y for d in annual.values() for y in d)} "
                f"to {max(y for d in annual.values() for y in d)}"
            )
            return annual

        except Exception as exc:
            self.log.error(f"SALM CSV parse error: {exc}", exc_info=True)
            return {}

    def _extract_town(self, town, data: dict):
        """Match town SA2 code and write cache JSON."""
        sa2 = town.sa2_code
        if not sa2:
            self.log.warning(f"  [{town.name}] no sa2_code in towns.toml")
            self.result.towns_skipped.append(town.name)
            return

        annual = data.get(sa2)
        if not annual:
            self.log.warning(f"  [{town.name}] SA2 {sa2} not found in SALM data")
            self.result.towns_failed.append(town.name)
            return

        from config import YEAR_START, YEAR_END
        values = {
            str(yr): val for yr, val in annual.items()
            if YEAR_START <= yr <= YEAR_END
        }

        latest_yr = max(annual)
        out = {
            "town":       town.name,
            "state":      town.state,
            "sa2_code":   sa2,
            "source":     "DEWR Small Area Labour Markets (SALM) — smoothed SA2",
            "source_url": RESOURCE_PAGE,
            "note":       "Annual value = mean of 4 quarterly smoothed unemployment rates",
            "indicators": {"unemployment": values},
        }

        out_dir = CACHE_DIR / "unemployment"
        out_dir.mkdir(exist_ok=True)
        out_path = out_dir / f"{town.slug}_salm.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        self.log.info(
            f"  {town.name} (SA2 {sa2}): {len(values)} years, "
            f"latest ({latest_yr}) = {annual[latest_yr]:.2f}%"
        )
        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = SALMUnemploymentFetcher().run()
    sys.exit(0 if result.success else 1)