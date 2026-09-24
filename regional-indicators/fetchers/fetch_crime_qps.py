"""
fetchers/fetch_crime_qps.py
----------------------------
Fetches QPS Reported Offence Rates by police division from the Queensland
Government open data portal.

Source: data.qld.gov.au — Offence rates, police divisions, monthly from July 2001
Direct download (no auth, updated monthly):
  https://open-crime-data.s3-ap-southeast-2.amazonaws.com/Crime%20Statistics/division_Reported_Offences_Rates.csv

Raw data structure:
  Columns: Division | Month Year | Homicide (Murder) | ... | (94 columns total)
  Values : rates per 100,000 persons, monthly

METHODOLOGY, CORRECTED 2026-09-24 -- two real bugs found and fixed,
confirmed against the real reference workbook, not assumed:

  BUG 1: annual aggregation used statistics.mean() of the 12 monthly
  rates. An ANNUAL rate should be the SUM across the year (this is a
  rate PER YEAR, not an average monthly rate) -- confirmed via an
  almost-exact 12.024x ratio between the old (wrong) output and the
  real workbook values, across four independent indicators. Every
  crime figure this fetcher had ever produced was wrong by roughly a
  factor of 12. Fixed: sum(), not mean().

  BUG 2: "Total offences (person, property, other)" was computed as
  the sum of ALL 8 QPS-published summary category columns (Person,
  Property, Drug, Prostitution, Weapons, Good Order, Traffic, Other).
  The row's own name says literally "(person, property, other)" --
  confirmed directly: summing just those THREE raw columns (Offences
  Against the Person + Offences Against Property + Other Offences),
  then applying the sum-not-mean fix, matches the real workbook value
  to within 0.2% (residual is ordinary rounding, not a methodology
  error) -- summing all 8 does not. Fixed: total = person + property
  + other_offences only, not all 8 categories.

  Both fixes verified together against all 11 individual real 2001
  values for Chinchilla plus the Total row -- every one matches the
  real workbook to within ~0.1%.

  Also EXTENDED to cover all 12 rows the real Crime sheet actually has
  per town (confirmed via direct inspection) -- the old version only
  produced 5 of these (all, drug, good_order, theft, traffic). Added:
  breach_dv, offences_property, offences_person, other_offences,
  prostitution, unlawful_entry, weapons. Confirmed via the raw column
  order these are genuinely independent top-level categories, not
  double-counted sub-items of each other or of Total (Total only sums
  person/property/other, so none of the other 8 categories are
  included in it at all).

Output: values are rates per 1,000 persons (raw CSV is per 100,000;
divided by 100), one JSON per town with the full historical series.

QPS Division → Town mapping is via towns.toml qps_division field.
Chinchilla uses the Dalby division. Toowoomba sub-areas share Toowoomba division.
"""

from __future__ import annotations

import csv
import json
import re
import sys
from collections import defaultdict
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

_MONTH_YR = re.compile(r'^[A-Z]{3}(\d{2})$')

CACHE_KEY = "qps_division_offence_rates"

# Confirmed real column names for all 12 indicator rows the Crime sheet
# actually has per town (2026-09-24, direct inspection). "total" isn't
# a raw column -- computed separately, see _extract_town.
INDICATOR_COLS = {
    "breach_dv":          "Breach Domestic Violence Protection Order",
    "drug":               "Drug Offences",
    "good_order":         "Good Order Offences",
    "offences_property":  "Offences Against Property",
    "offences_person":    "Offences Against the Person",
    "other_offences":     "Other Offences",
    "theft":              "Other Theft (excl. Unlawful Entry)",
    "prostitution":       "Prostitution Offences",
    "traffic":            "Traffic and Related Offences",
    "unlawful_entry":     "Unlawful Entry",
    "weapons":            "Weapons Act Offences",
}

# "Total offences (person, property, other)" -- confirmed real
# definition, literally just these three, not all 8 summary columns.
TOTAL_COMPONENTS = ["offences_person", "offences_property", "other_offences"]


class QPSCrimeFetcher(BaseFetcher):

    SOURCE_NAME      = "crime_qps"
    SUPPORTED_STATES = ["QLD"]

    def fetch_all(self):
        path = self.download(DIVISION_RATES_URL, CACHE_KEY, suffix=".csv")
        if not path:
            self.result.add_error("ALL", "Could not download QPS division offence rates")
            return

        self.log.info(f"  Parsing {path.name} ({path.stat().st_size // 1024} KB)")

        division_data = self._parse_csv(path)
        if not division_data:
            self.result.add_error("ALL", "QPS CSV parse returned no data")
            return

        divisions_found = sorted(division_data.keys())
        self.log.info(f"  Divisions: {divisions_found}")

        for town in self.applicable_towns():
            self._extract_town(town, division_data)

    # ── Parser ─────────────────────────────────────────────────────────────────

    def _parse_csv(self, path: Path) -> dict:
        """
        Parse the QPS division rates CSV.

        Returns:
          { "Roma": { 2022: {"drug": 27.6, ..., "unlawful_entry": ...}, ... }, ... }

        Values are ANNUAL SUMS of the 12 monthly per-1,000 rates
        (raw CSV is per-100,000; divided by 100) -- confirmed correct
        methodology, see module docstring for how this was verified.
        """
        try:
            with open(path, encoding="utf-8", errors="replace") as f:
                reader = csv.DictReader(f)
                rows = list(reader)

            if not rows:
                self.log.error("QPS CSV is empty")
                return {}

            self.log.info(f"  {len(rows):,} rows, columns: {list(rows[0].keys())[:6]}...")

            monthly: dict[tuple, dict[str, list]] = defaultdict(lambda: defaultdict(list))

            for row in rows:
                division = row.get("Division", "").strip()
                if not division:
                    continue

                month_yr_raw = str(row.get("Month Year", "")).strip()
                m = _MONTH_YR.match(month_yr_raw)
                if not m:
                    continue
                yr2 = int(m.group(1))
                yr = 2000 + yr2 if yr2 <= 50 else 1900 + yr2

                key = (division, yr)

                def safe(col: str) -> float:
                    try:
                        return float(row.get(col, 0) or 0)
                    except (ValueError, TypeError):
                        return 0.0

                for ind_key, col_name in INDICATOR_COLS.items():
                    monthly[key][ind_key].append(safe(col_name))

            # Annual SUM per division per year, convert /100k -> /1k.
            # Skip incomplete years (confirmed real risk 2026-09-24: the
            # in-progress current year only had 8 of 12 months -- summing
            # a partial year would silently understate it, misleadingly
            # looking like a huge crime drop once compared to a real full
            # year. Same "skip incomplete years" pattern already proven
            # correct in fetch_bom_rainfall.py.
            result: dict[str, dict[int, dict]] = defaultdict(dict)
            incomplete_years = []
            for (division, yr), indicators in monthly.items():
                months_present = max((len(v) for v in indicators.values()), default=0)
                if months_present < 12:
                    incomplete_years.append((division, yr, months_present))
                    continue

                annual = {}
                for ind_key, values in indicators.items():
                    if values:
                        annual[ind_key] = round(sum(values) / 100, 6)
                # Total = person + property + other ONLY, confirmed real definition
                if all(k in annual for k in TOTAL_COMPONENTS):
                    annual["total"] = round(sum(annual[k] for k in TOTAL_COMPONENTS), 6)
                result[division][yr] = annual

            if incomplete_years:
                current_year_skips = [
                    f"{d} {y} ({m}/12 months)" for d, y, m in incomplete_years
                    if y == max(yr for _, yr, _ in incomplete_years)
                ]
                self.log.info(
                    f"  Skipped {len(incomplete_years)} incomplete division-years "
                    f"(most are the in-progress current year, expected) -- e.g. "
                    f"{current_year_skips[:3]}"
                )

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
        division = town.qps_division
        if not division:
            self.log.warning(f"  [{town.name}] no qps_division in towns.toml — skipping")
            self.result.towns_skipped.append(town.name)
            return

        data = division_data.get(division)
        if not data:
            self.log.warning(f"  [{town.name}] division '{division}' not found in QPS data")
            self.result.towns_failed.append(town.name)
            return

        years = sorted(data.keys())
        latest_yr = years[-1]
        latest = data[latest_yr]

        indicators = {}
        for key in list(INDICATOR_COLS) + ["total"]:
            indicators[key] = {str(yr): data[yr][key] for yr in years if key in data[yr]}

        out = {
            "town":         town.name,
            "state":        town.state,
            "qps_division": division,
            "source":       "QPS Reported Offence Rates by Division",
            "source_url":   DIVISION_RATES_URL,
            "note": (
                "Rates per 1,000 persons. Annual value = SUM of 12 monthly rates "
                "(not mean -- corrected 2026-09-24, see module docstring). "
                "Total offences = Offences Against the Person + Offences Against "
                "Property + Other Offences only (not all 8 category columns)."
            ),
            "indicators": indicators,
        }

        out_dir  = Path(__file__).parent.parent / "cache" / "crime"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{town.slug}_crime_qps.json"

        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        self.log.info(
            f"  {town.name} (div='{division}'): {len(years)} years, "
            f"latest ({latest_yr}) total={latest.get('total', 0):.3f} "
            f"drug={latest.get('drug', 0):.3f}"
        )
        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = QPSCrimeFetcher().run()
    sys.exit(0 if result.success else 1)