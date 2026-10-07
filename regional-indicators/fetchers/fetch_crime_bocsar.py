"""
regional-indicators/fetchers/fetch_crime_bocsar.py

Fetches NSW recorded crime incident data from the NSW Bureau of Crime
Statistics and Research (BOCSAR) and converts it to annual rates per
1,000 persons.

Source: bocsar.nsw.gov.au/statistics-dashboards/open-datasets/criminal-offences-data.html
Direct downloads (no authentication, updated quarterly):
  NSW-wide: https://bocsar.nsw.gov.au/content/dam/dcj/bocsar/documents/open-datasets/Incident_by_NSW.xlsx
  By LGA:   https://bocsarblob.blob.core.windows.net/bocsar-open-data/RCI_offencebymonth.xlsm

Population for LGA-level rates:
  ABS Regional population by LGA, 2001-2025 (annual, 30 June reference
  date):
  https://www.abs.gov.au/statistics/people/population/regional-population/2024-25/32180DS0004_2001-25.xlsx

File layout:
  Both BOCSAR files are wide format: one row per (geography, offence
  category, subcategory) and one column per month from 1995 onward. The
  NSW-wide file carries population in each row ("2025 population" and
  "2026 population" columns). The LGA file does not, so LGA rates use
  the ABS population lookup above.

Methodology:
  Rate = sum of the 12 monthly incident counts across all subcategories
  within a category for the year, divided by (population / 1000). The
  annual figure is a sum, not a monthly mean, as for the QPS fetcher.
  Checked against the reference workbook: NSW-wide Assault 2024 computes
  to 8.98 against 8.85, and Narrabri LGA Assault 2024 to 13.71 against
  13.91 (both about 1.5%), the same margin as the other states' crime
  data.

  Incomplete years are skipped; BOCSAR's most recent year is normally
  partial. Years with no matching population figure are also skipped
  (the ABS series currently ends at 2025). A year needs both a complete
  12 months of incidents and a population figure to produce a rate.

Categories: the eight sheet categories (Assault, Drug offences,
Malicious damage to property, Other offences, Other offences against
the person, Robbery, Theft, Transport regulatory offences) are exact
"Offence category" values in the BOCSAR data, so no grouping or
combining is needed.

Template for other states: this fetcher is the model for future state
crime fetchers. The NSW-specific parts are the two BOCSAR URLs, the
wide-format parser with one column per month, the eight-category label
list, and the split between files with and without population. The
state-agnostic parts, which a new fetcher should reuse, are the output
JSON schema (indicators[key] = {"label": ..., "values": {...}}, read
generically by update_crime.py), the sum-not-mean annual aggregation
and the skipping of incomplete years. A fetcher that writes the same
JSON shape needs no change to update_crime.py.

Geography: the existing `lga` field of NSW towns in towns.toml is used
as the BOCSAR LGA name. It matches the BOCSAR LGA column verbatim (for
example `lga = "Narrabri"`), so no extra towns.toml field is needed.

Output: cache/crime/<slug>_crime_bocsar.json per town and
cache/crime/nsw_crime_bocsar.json for the state.
"""

from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher

try:
    import openpyxl
    import requests
except ImportError:
    raise ImportError("pip install openpyxl requests")


NSW_WIDE_URL = (
    "https://bocsar.nsw.gov.au/content/dam/dcj/bocsar/documents"
    "/open-datasets/Incident_by_NSW.xlsx"
)
LGA_URL = "https://bocsarblob.blob.core.windows.net/bocsar-open-data/RCI_offencebymonth.xlsm"
ABS_LGA_POP_BASE = "https://www.abs.gov.au/statistics/people/population/regional-population"
ABS_LGA_POP_RELEASE = "2024-25"   # update each cycle if the automatic advance below fails
ABS_LGA_POP_URL = (
    f"{ABS_LGA_POP_BASE}/{ABS_LGA_POP_RELEASE}/32180DS0004_2001-25.xlsx"
)
# Update each cycle: both the "/2024-25/" release folder and the
# "2001-25" file name suffix advance each year ABS republishes this
# series. A retired URL returns 404 rather than serving stale data.


def _guess_next_lga_pop_release(release: str) -> tuple[str, str] | None:
    """Return (release_folder, filename_suffix) for the release after the
    given one, for example "2024-25" -> ("2025-26", "26").

    Returns None if the release does not have the "YYYY-YY" shape, so
    that the caller falls back to the configured release.
    """
    m = re.match(r"^(\d{4})-(\d{2})$", release)
    if not m:
        return None
    start = int(m.group(1))
    new_start, new_end2 = start + 1, (start + 2) % 100
    return f"{new_start}-{new_end2:02d}", f"{new_end2:02d}"


def _discover_current_lga_pop_url(log) -> str:
    """Return the URL of the current ABS LGA population workbook.

    Uses the same guess-verify-fallback approach as
    _discover_current_release() in fetch_business.py: the next release
    is guessed from the configured one and checked with a HEAD request.
    The guessed URL is used only if it returns 200; otherwise the
    configured ABS_LGA_POP_URL is returned.
    """
    guess = _guess_next_lga_pop_release(ABS_LGA_POP_RELEASE)
    if not guess:
        log.warning(f"  ABS release {ABS_LGA_POP_RELEASE!r} doesn't match the "
                    f"expected 'YYYY-YY' shape -- using hardcoded URL")
        return ABS_LGA_POP_URL
    folder, suffix = guess
    guess_url = f"{ABS_LGA_POP_BASE}/{folder}/32180DS0004_2001-{suffix}.xlsx"
    try:
        resp = requests.head(guess_url, timeout=15, allow_redirects=True)
        if resp.status_code == 200:
            log.info(f"  Auto-advanced ABS LGA population release: "
                     f"{ABS_LGA_POP_RELEASE} -> {folder} (verified live)")
            return guess_url
    except requests.RequestException as exc:
        log.warning(f"  Could not check guessed release {folder!r}: {exc}")
    log.info(f"  Guessed release {folder!r} not live yet -- using hardcoded URL")
    return ABS_LGA_POP_URL

NSW_WIDE_CACHE_KEY = "bocsar_nsw_wide"
LGA_CACHE_KEY = "bocsar_lga"
ABS_POP_CACHE_KEY = "abs_lga_population_2001_2025"   # cache file name only; need not track the release

# Exact "Offence category" values in both BOCSAR files, in the row order
# of the Crime sheet.
CATEGORIES = [
    "Assault",
    "Drug offences",
    "Malicious damage to property",
    "Other offences",
    "Other offences against the person",
    "Robbery",
    "Theft",
    "Transport regulatory offences",
]


class BOCSARCrimeFetcher(BaseFetcher):

    SOURCE_NAME      = "crime_bocsar"
    SUPPORTED_STATES = ["NSW"]

    def fetch_all(self):
        pop_url = _discover_current_lga_pop_url(self.log)
        pop_path = self.download(pop_url, ABS_POP_CACHE_KEY, suffix=".xlsx")
        if not pop_path:
            self.result.add_error("ALL", "Could not download ABS LGA population data")
            return
        lga_population = self._parse_abs_lga_population(pop_path)
        if not lga_population:
            self.result.add_error("ALL", "ABS LGA population parse returned no data")
            return

        nsw_path = self.download(NSW_WIDE_URL, NSW_WIDE_CACHE_KEY, suffix=".xlsx")
        if nsw_path:
            self._extract_nsw_wide(nsw_path)
        else:
            self.log.warning("  Could not download NSW-wide BOCSAR data -- skipping benchmark")

        lga_path = self.download(LGA_URL, LGA_CACHE_KEY, suffix=".xlsm")
        if not lga_path:
            self.result.add_error("ALL", "Could not download BOCSAR LGA data")
            return

        lga_data = self._parse_lga_csv(lga_path, lga_population)
        if not lga_data:
            self.result.add_error("ALL", "BOCSAR LGA parse returned no data")
            return

        for town in self.applicable_towns():
            self._extract_town(town, lga_data)

    # ── ABS LGA population ────────────────────────────────────────────────────

    def _parse_abs_lga_population(self, path: Path) -> dict:
        """Return { "Narrabri": {2001: 14422, ..., 2025: 12797}, ... }.

        Workbook layout: sheet "Table 1", year headers (2001-2025) in
        row 5, data from row 7, column A = LGA code, column B = LGA name,
        columns C onward = one per year.
        """
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            ws = wb["Table 1"]
            years = list(range(2001, 2026))
            result = {}
            for row in ws.iter_rows(min_row=7, values_only=True):
                lga_name = row[1]
                if not lga_name:
                    continue
                result[lga_name] = {
                    yr: val for yr, val in zip(years, row[2:2 + len(years)])
                    if val is not None
                }
            self.log.info(f"  Parsed ABS population for {len(result)} LGAs")
            return result
        except Exception as exc:
            self.log.error(f"ABS LGA population parse error: {exc}", exc_info=True)
            return {}

    # ── Shared wide-format parser ─────────────────────────────────────────────

    def _sum_category_by_year(self, rows, geography_col_value, geography_col_idx,
                               month_cols: dict) -> dict:
        """Sum monthly incident counts by year and category for one
        geography.

        Takes raw rows (geography, offence category, subcategory, then
        the monthly columns) and returns { year: { category: sum } },
        summing all subcategories within each category. Only years with
        a complete 12 months are returned.
        """
        monthly: dict[int, dict[str, list]] = defaultdict(lambda: defaultdict(list))
        for row in rows:
            if geography_col_idx is not None and row[geography_col_idx] != geography_col_value:
                continue
            category = row[1]
            if category not in CATEGORIES:
                continue
            for yr, cols in month_cols.items():
                for c in cols:
                    val = row[c]
                    monthly[yr][category].append(val or 0)

        result = {}
        for yr, cats in monthly.items():
            # Completeness is judged from the header's month columns, not
            # from the number of values collected: a category with several
            # subcategory rows contributes one value per row per month, so
            # a count of values can reach 12 for a part year.
            if len(month_cols.get(yr, [])) != 12:
                continue
            result[yr] = {cat: sum(vals) for cat, vals in cats.items()}
        return result

    def _month_columns_by_year(self, header_row) -> dict:
        result = defaultdict(list)
        for c, v in enumerate(header_row):
            if isinstance(v, datetime):
                result[v.year].append(c)
        return dict(result)

    # ── NSW-wide (population built in) ────────────────────────────────────────

    def _extract_nsw_wide(self, path: Path):
        self.log.info(f"  Parsing {path.name} ({path.stat().st_size // 1024} KB)")
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            ws = wb["Data"]
            rows = list(ws.iter_rows(values_only=True))
        except Exception as exc:
            self.log.error(f"NSW-wide parse error: {exc}", exc_info=True)
            self.result.add_error("NSW", f"parse error: {exc}")
            return

        header = rows[0]
        month_cols = self._month_columns_by_year(header)
        data_rows = rows[1:]

        yearly_sums = self._sum_category_by_year(data_rows, None, None, month_cols)

        # Population is carried in every row and is the same in each (it
        # is the state population, not a per-category figure), so any row
        # supplies it.
        population_by_year: dict[int, int] = {}
        for row in data_rows:
            if row[3]:
                population_by_year[2025] = row[3]
            if row[4]:
                population_by_year[2026] = row[4]

        # In the current release only 2025 has both a complete year and a
        # population figure. Later years are picked up in the same way as
        # the source file grows.
        indicators = {}
        for cat in CATEGORIES:
            indicators[cat] = {"label": cat, "values": {}}

        for yr, cats in yearly_sums.items():
            pop = population_by_year.get(yr)
            if not pop:
                continue
            for cat, raw_sum in cats.items():
                rate = round(raw_sum / (pop / 1000), 6)
                indicators[cat]["values"][str(yr)] = rate

        out = {
            "town":       "NSW",
            "state":      "NSW",
            "lga":        None,
            "source":     "BOCSAR Recorded Criminal Incidents, NSW statewide",
            "source_url": NSW_WIDE_URL,
            "note": (
                "Rates per 1,000 persons. Annual value = SUM of 12 monthly "
                "incident counts across all subcategories within each "
                "category, divided by NSW population for that year."
            ),
            "indicators": indicators,
        }

        out_dir = Path(__file__).parent.parent / "cache" / "crime"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / "nsw_crime_bocsar.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        latest_yr = max((int(y) for cat in indicators.values() for y in cat["values"]), default=None)
        self.log.info(f"  NSW (statewide): latest complete year with population = {latest_yr}")
        self.result.towns_ok.append("NSW")

    # ── LGA-level (population looked up separately) ──────────────────────────

    def _parse_lga_csv(self, path: Path, lga_population: dict) -> dict:
        self.log.info(f"  Parsing {path.name} ({path.stat().st_size // 1024} KB)")
        try:
            wb = openpyxl.load_workbook(path, read_only=True, keep_vba=True)
            ws = wb["Data"]
            rows = list(ws.iter_rows(values_only=True))
        except Exception as exc:
            self.log.error(f"BOCSAR LGA parse error: {exc}", exc_info=True)
            return {}

        header = rows[0]
        month_cols = self._month_columns_by_year(header)
        data_rows = rows[1:]

        lgas_present = {row[0] for row in data_rows if row[0]}
        result = {}
        for lga in lgas_present:
            yearly_sums = self._sum_category_by_year(data_rows, lga, 0, month_cols)
            pop_series = lga_population.get(lga, {})

            annual = {}
            for yr, cats in yearly_sums.items():
                pop = pop_series.get(yr)
                if not pop:
                    continue
                annual[yr] = {
                    cat: round(raw_sum / (pop / 1000), 6) for cat, raw_sum in cats.items()
                }
            if annual:
                result[lga] = annual

        self.log.info(f"  Parsed {len(result)} LGAs with usable data")
        return result

    def _extract_town(self, town, lga_data: dict):
        lga = getattr(town, "lga", None)
        if not lga:
            self.log.warning(f"  [{town.name}] no lga in towns.toml — skipping")
            self.result.towns_skipped.append(town.name)
            return

        data = lga_data.get(lga)
        if not data:
            self.log.warning(f"  [{town.name}] LGA '{lga}' not found in BOCSAR data")
            self.result.towns_failed.append(town.name)
            return

        years = sorted(data.keys())
        latest_yr = years[-1]

        indicators = {}
        for cat in CATEGORIES:
            indicators[cat] = {
                "label": cat,
                "values": {str(yr): data[yr][cat] for yr in years if cat in data[yr]},
            }

        out = {
            "town":       town.name,
            "state":      town.state,
            "lga":        lga,
            "source":     "BOCSAR Recorded Criminal Incidents by LGA",
            "source_url": LGA_URL,
            "note": (
                "Rates per 1,000 persons. Annual value = SUM of 12 monthly "
                "incident counts across all subcategories within each "
                "category, divided by ABS LGA population for that year "
                "(30 June reference date)."
            ),
            "indicators": indicators,
        }

        out_dir = Path(__file__).parent.parent / "cache" / "crime"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{town.slug}_crime_bocsar.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        self.log.info(
            f"  {town.name} (lga='{lga}'): {len(years)} years, "
            f"latest ({latest_yr}) assault={data[latest_yr].get('Assault', 0):.3f}"
        )
        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = BOCSARCrimeFetcher().run()
    sys.exit(0 if result.success else 1)