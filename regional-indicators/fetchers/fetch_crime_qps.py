"""
fetchers/fetch_crime_qps.py
----------------------------
Fetches QPS Reported Offence Rates by police division from the Queensland
Government open data portal.

Source: data.qld.gov.au — Offence rates, police divisions, monthly from July 2001
Direct download (no auth, updated monthly):
  https://open-crime-data.s3-ap-southeast-2.amazonaws.com/Crime%20Statistics/division_Reported_Offences_Rates.csv

See README.md → Data sources → QPS Crime Statistics for full column mapping.

Raw data structure:
  Columns: Division | Month Year | Homicide | ... | (90 offence columns)
  Values : rates per 100,000 persons, monthly

Processing:
  1. Filter rows by QPS division (from town.qps_division in towns.toml)
  2. For each calendar year: take all months in that year, compute mean
  3. Divide by 100 to convert per-100,000 → per-1,000 (matching website format)
  4. Aggregate across categories into the 5 output indicators

Output CSVs (values are rates per 1,000 persons):
  Crime rate - all offences.csv        ← total of all categories
  Drug offences.csv                    ← Drug Offences column
  Good order offences.csv              ← Good Order Offences column
  Theft.csv                            ← Other Theft (excl. Unlawful Entry) column
  Traffic offences.csv                 ← Traffic and Related Offences column

QPS Division → Town mapping is via towns.toml qps_division field.
Chinchilla uses the Dalby division. Toowoomba sub-areas share Toowoomba division.
"""

from __future__ import annotations

import csv
import io
import json
import statistics
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher

try:
    import requests
except ImportError:
    raise ImportError("pip install requests")


# ── Configuration ──────────────────────────────────────────────────────────────

DIVISION_RATES_URL = (
    "https://open-crime-data.s3-ap-southeast-2.amazonaws.com"
    "/Crime%20Statistics/division_Reported_Offences_Rates.csv"
)

CACHE_KEY = "qps_division_offence_rates"

# Column names in the raw CSV that we aggregate into each output indicator.
# "Crime rate - all offences" = sum of all offence category columns.
# All other categories are single columns.
INDICATOR_COLS = {
    "drug":       "Drug Offences",
    "good_order": "Good Order Offences",
    "theft":      "Other Theft (excl. Unlawful Entry)",
    "traffic":    "Traffic and Related Offences",
}

# All offence category columns — used for "all offences" total
# These are the pre-aggregated category columns (not the sub-type breakdowns)
ALL_OFFENCE_COLS = [
    "Offences Against the Person",
    "Offences Against Property",
    "Drug Offences",
    "Prostitution Offences",
    "Weapons Act Offences",
    "Good Order Offences",
    "Traffic and Related Offences",
    "Other Offences",
]


class QPSCrimeFetcher(BaseFetcher):

    SOURCE_NAME      = "crime_qps"
    SUPPORTED_STATES = ["QLD"]

    def fetch_all(self):
        # ── Download ──────────────────────────────────────────────────────────
        path = self.download(DIVISION_RATES_URL, CACHE_KEY, suffix=".csv")
        if not path:
            self.result.add_error("ALL", "Could not download QPS division offence rates")
            return

        self.log.info(f"  Parsing {path.name} ({path.stat().st_size // 1024} KB)")

        # ── Parse ─────────────────────────────────────────────────────────────
        division_data = self._parse_csv(path)
        if not division_data:
            self.result.add_error("ALL", "QPS CSV parse returned no data")
            return

        divisions_found = sorted(division_data.keys())
        self.log.info(f"  Divisions: {divisions_found}")

        # ── Extract per town ──────────────────────────────────────────────────
        for town in self.applicable_towns():
            self._extract_town(town, division_data)

    # ── Parser ─────────────────────────────────────────────────────────────────

    def _parse_csv(self, path: Path) -> dict:
        """
        Parse the QPS division rates CSV.

        Returns:
          {
            "Roma": {
              2022: {"drug": 27.6, "good_order": 21.2, "theft": 21.4,
                     "traffic": 23.1, "all": 168.4},
              2023: { ... },
              ...
            },
            ...
          }

        Values are annual means of monthly per-1,000 rates
        (raw CSV is per-100,000; we divide by 100).
        """
        try:
            with open(path, encoding="utf-8", errors="replace") as f:
                reader = csv.DictReader(f)
                rows = list(reader)

            if not rows:
                self.log.error("QPS CSV is empty")
                return {}

            self.log.info(f"  {len(rows):,} rows, columns: {list(rows[0].keys())[:6]}...")

            # Group monthly rows by (division, year)
            # Month Year column format: "2024-01-01 00:00:00" or "2001-07-01 00:00:00"
            monthly: dict[tuple, dict[str, list]] = defaultdict(lambda: defaultdict(list))

            for row in rows:
                division = row.get("Division", "").strip()
                if not division:
                    continue

                # Parse year from Month Year column
                month_yr_raw = row.get("Month Year", "")
                try:
                    yr = int(str(month_yr_raw)[:4])
                except (ValueError, TypeError):
                    continue

                key = (division, yr)

                # Accumulate per-100,000 values for each indicator
                def safe(col: str) -> float:
                    try:
                        return float(row.get(col, 0) or 0)
                    except (ValueError, TypeError):
                        return 0.0

                for ind_key, col_name in INDICATOR_COLS.items():
                    monthly[key][ind_key].append(safe(col_name))

                # Total all offences
                total = sum(safe(c) for c in ALL_OFFENCE_COLS if c in row)
                monthly[key]["all"].append(total)

            # Annual mean per division per year, convert /100k → /1k
            result: dict[str, dict[int, dict]] = defaultdict(dict)
            for (division, yr), indicators in monthly.items():
                annual = {}
                for ind_key, values in indicators.items():
                    if values:
                        # Mean of monthly rates, then /100 to get per-1,000
                        annual[ind_key] = round(statistics.mean(values) / 100, 6)
                result[division][yr] = annual

            self.log.info(
                f"  Parsed {len(result)} divisions, "
                f"years {min(yr for d in result.values() for yr in d)} "
                f"to {max(yr for d in result.values() for yr in d)}"
            )
            return dict(result)

        except Exception as exc:
            self.log.error(f"QPS CSV parse error: {exc}", exc_info=True)
            return {}

    # ── Per-town extraction ────────────────────────────────────────────────────

    def _extract_town(self, town, division_data: dict):
        """
        Look up the town's QPS division, extract annual series, write JSON.
        """
        division = town.qps_division
        if not division:
            self.log.warning(f"  [{town.name}] no qps_division in towns.toml — skipping")
            self.result.towns_skipped.append(town.name)
            return

        data = division_data.get(division)
        if not data:
            self.log.warning(
                f"  [{town.name}] division '{division}' not found in QPS data"
            )
            self.result.towns_failed.append(town.name)
            return

        # Sort by year
        years = sorted(data.keys())
        latest_yr = years[-1]
        latest = data[latest_yr]

        out = {
            "town":         town.name,
            "state":        town.state,
            "qps_division": division,
            "source":       "QPS Reported Offence Rates by Division",
            "source_url":   DIVISION_RATES_URL,
            "note":         "Rates per 1,000 persons. Annual value = mean of 12 monthly rates.",
            "indicators": {
                "all":        {str(yr): data[yr]["all"]        for yr in years},
                "drug":       {str(yr): data[yr]["drug"]       for yr in years},
                "good_order": {str(yr): data[yr]["good_order"] for yr in years},
                "theft":      {str(yr): data[yr]["theft"]      for yr in years},
                "traffic":    {str(yr): data[yr]["traffic"]    for yr in years},
            }
        }

        out_dir  = Path(__file__).parent.parent / "cache" / "crime"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{town.slug}_crime_qps.json"

        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        self.log.info(
            f"  {town.name} (div='{division}'): {len(years)} years, "
            f"latest ({latest_yr}) all={latest['all']:.3f} "
            f"drug={latest['drug']:.3f}"
        )
        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = QPSCrimeFetcher().run()
    sys.exit(0 if result.success else 1)
