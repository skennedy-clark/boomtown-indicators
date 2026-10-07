"""
regional-indicators/fetchers/fetch_crime_qps.py

Fetches Queensland Police Service (QPS) reported offence rates by police
division, plus the Queensland statewide benchmark, from the Queensland
Government open data portal.

Source: data.qld.gov.au - Offence rates, police divisions, monthly from
July 2001. Direct download (no authentication, updated monthly):
  https://open-crime-data.s3-ap-southeast-2.amazonaws.com/Crime%20Statistics/division_Reported_Offences_Rates.csv
The statewide file has the same layout without the Division column:
  https://open-crime-data.s3-ap-southeast-2.amazonaws.com/Crime%20Statistics/QLD_Reported_Offences_Rates.csv

Raw data structure:
  Columns: Division | Month Year | Homicide (Murder) | ... (94 columns)
  Values : rates per 100,000 persons, monthly

Methodology:
  Annual value. The annual rate is the sum of the 12 monthly rates (a
  rate per year), not their mean. Taking the mean understates every
  figure by a factor of about 12.

  Total. "Total offences (person, property, other)" is the sum of three
  raw columns only: Offences Against the Person + Offences Against
  Property + Other Offences. QPS publishes eight summary categories
  (Person, Property, Drug, Prostitution, Weapons, Good Order, Traffic,
  Other); summing all eight does not reproduce the reference workbook.

  Incomplete years are skipped, so the in-progress current year never
  produces a partial-year sum.

  With these rules the output matches the reference workbook to within
  about 0.2% (rounding) for the Chinchilla 2001 values and the Total
  row.

  Indicators. All 12 rows of the Crime sheet are produced per town: the
  11 raw columns in INDICATOR_COLS plus the computed total. The raw
  columns are independent top-level categories; none is a sub-item of
  another, and only person, property and other contribute to the total.

Output: cache/crime/<slug>_crime_qps.json per town, and
cache/crime/queensland_crime_qps.json for the state, each with the full
historical series. Values are rates per 1,000 persons (the raw CSV is
per 100,000; divided by 100).

Geography: each town maps to a QPS division through the qps_division
field in towns.toml. Chinchilla uses the Dalby division. The Toowoomba
sub-areas share the Toowoomba division.
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

# Queensland statewide rates. QPS publishes these directly in the same
# layout as the division file, so no population lookup or aggregation
# across divisions is needed.
QLD_RATES_URL = (
    "https://open-crime-data.s3-ap-southeast-2.amazonaws.com"
    "/Crime%20Statistics/QLD_Reported_Offences_Rates.csv"
)
QLD_CACHE_KEY = "qps_qld_statewide_offence_rates"

# Raw column names for the indicator rows of the Crime sheet. "total" is
# not a raw column; it is computed from TOTAL_COMPONENTS in the parsers.
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

# "Total offences (person, property, other)" sums these three columns
# only, not all eight summary categories.
TOTAL_COMPONENTS = ["offences_person", "offences_property", "other_offences"]

# Sheet-row label for each indicator, written into the JSON output so
# that update_crime.py reads labels from the fetcher output and needs no
# per-state label table. The raw column names are identical to the sheet
# row labels, so INDICATOR_COLS is reused here.
INDICATOR_LABELS = {**INDICATOR_COLS, "total": "Total offences (person, property, other)"}


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

        # Queensland state benchmark, from the statewide rates file. The result
        # matches the reference workbook's 2024 Queensland values to within
        # about 0.2-0.7%, the same margin as the town-level data.
        qld_path = self.download(QLD_RATES_URL, QLD_CACHE_KEY, suffix=".csv")
        if qld_path:
            self._extract_queensland(qld_path)
        else:
            self.log.warning("  Could not download QLD statewide offence rates -- skipping benchmark")

    # ── Parser ─────────────────────────────────────────────────────────────────

    def _parse_csv(self, path: Path) -> dict:
        """Parse the QPS division rates CSV.

        Returns:
          { "Roma": { 2022: {"drug": 27.6, ..., "unlawful_entry": ...}, ... }, ... }

        Values are annual sums of the 12 monthly rates, converted to per
        1,000 persons (the raw CSV is per 100,000; divided by 100). See the
        module docstring for the methodology.
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

            # Annual sum per division per year, converted from per 100,000 to
            # per 1,000. Years with fewer than 12 months are skipped: summing a
            # partial year (normally the in-progress current year) would
            # understate it and read as a sharp fall in crime. The same rule is
            # used in fetch_bom_rainfall.py.
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
                # Total = person + property + other only.
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

        # Each indicator carries its sheet-row label alongside its values,
        # so update_crime.py needs no per-state label table.
        indicators = {}
        for key in list(INDICATOR_COLS) + ["total"]:
            indicators[key] = {
                "label": INDICATOR_LABELS[key],
                "values": {str(yr): data[yr][key] for yr in years if key in data[yr]},
            }

        out = {
            "town":         town.name,
            "state":        town.state,
            "qps_division": division,
            "source":       "QPS Reported Offence Rates by Division",
            "source_url":   DIVISION_RATES_URL,
            "note": (
                "Rates per 1,000 persons. Annual value = sum of the 12 monthly "
                "rates (not the mean; see the module docstring). "
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

    # ── Queensland statewide benchmark ────────────────────────────────────────

    def _parse_statewide_csv(self, path: Path) -> dict:
        """Parse the QPS statewide rates CSV.

        Same methodology as _parse_csv (sum of 12 months, incomplete years
        skipped, total = person + property + other only). It is a separate
        method because the statewide file has no Division column to group
        by; the whole file is a single implicit division.

        Returns: { 2022: {"drug": 27.6, ..., "total": ...}, ... }
        """
        try:
            with open(path, encoding="utf-8", errors="replace") as f:
                reader = csv.DictReader(f)
                rows = list(reader)

            if not rows:
                self.log.error("QLD statewide CSV is empty")
                return {}

            monthly: dict[int, dict[str, list]] = defaultdict(lambda: defaultdict(list))

            for row in rows:
                month_yr_raw = str(row.get("Month Year", "")).strip()
                m = _MONTH_YR.match(month_yr_raw)
                if not m:
                    continue
                yr2 = int(m.group(1))
                yr = 2000 + yr2 if yr2 <= 50 else 1900 + yr2

                def safe(col: str) -> float:
                    try:
                        return float(row.get(col, 0) or 0)
                    except (ValueError, TypeError):
                        return 0.0

                for ind_key, col_name in INDICATOR_COLS.items():
                    monthly[yr][ind_key].append(safe(col_name))

            result: dict[int, dict] = {}
            incomplete_years = []
            for yr, indicators in monthly.items():
                months_present = max((len(v) for v in indicators.values()), default=0)
                if months_present < 12:
                    incomplete_years.append((yr, months_present))
                    continue

                annual = {}
                for ind_key, values in indicators.items():
                    if values:
                        annual[ind_key] = round(sum(values) / 100, 6)
                if all(k in annual for k in TOTAL_COMPONENTS):
                    annual["total"] = round(sum(annual[k] for k in TOTAL_COMPONENTS), 6)
                result[yr] = annual

            if incomplete_years:
                self.log.info(f"  QLD statewide: skipped incomplete years {incomplete_years}")

            return result

        except Exception as exc:
            self.log.error(f"QLD statewide CSV parse error: {exc}", exc_info=True)
            return {}

    def _extract_queensland(self, path: Path):
        self.log.info(f"  Parsing {path.name} ({path.stat().st_size // 1024} KB)")
        data = self._parse_statewide_csv(path)
        if not data:
            self.result.add_error("Queensland", "QLD statewide CSV parse returned no data")
            return

        years = sorted(data.keys())
        latest_yr = years[-1]
        latest = data[latest_yr]

        # Each indicator carries its sheet-row label alongside its values,
        # so update_crime.py needs no per-state label table.
        indicators = {}
        for key in list(INDICATOR_COLS) + ["total"]:
            indicators[key] = {
                "label": INDICATOR_LABELS[key],
                "values": {str(yr): data[yr][key] for yr in years if key in data[yr]},
            }

        out = {
            "town":         "Queensland",
            "state":        "QLD",
            "qps_division": None,
            "source":       "QPS Reported Offence Rates, Queensland statewide",
            "source_url":   QLD_RATES_URL,
            "note": (
                "Rates per 1,000 persons. Same methodology as individual "
                "division rows: annual value = SUM of 12 monthly rates, "
                "total = Offences Against the Person + Offences Against "
                "Property + Other Offences only."
            ),
            "indicators": indicators,
        }

        out_dir = Path(__file__).parent.parent / "cache" / "crime"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / "queensland_crime_qps.json"

        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        self.log.info(
            f"  Queensland (statewide): {len(years)} years, "
            f"latest ({latest_yr}) total={latest.get('total', 0):.3f}"
        )
        self.result.towns_ok.append("Queensland")


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = QPSCrimeFetcher().run()
    sys.exit(0 if result.success else 1)