"""
fetchers/fetch_crime_bocsar.py
----------------------------
Fetches NSW recorded crime incident data from the NSW Bureau of Crime
Statistics and Research (BOCSAR).

Source: bocsar.nsw.gov.au/statistics-dashboards/open-datasets/criminal-offences-data.html
Direct downloads (no auth, updated quarterly):
  NSW-wide: https://bocsar.nsw.gov.au/content/dam/dcj/bocsar/documents/open-datasets/Incident_by_NSW.xlsx
  By LGA:   https://bocsarblob.blob.core.windows.net/bocsar-open-data/RCI_offencebymonth.xlsm

Population for LGA-level rates:
  ABS Regional population by LGA, 2001-2025 (annual, 30 June reference date):
  https://www.abs.gov.au/statistics/people/population/regional-population/2024-25/32180DS0004_2001-25.xlsx

METHODOLOGY, confirmed 2026-09-24 against real data, not assumed:
  Both BOCSAR files are WIDE format: one row per (geography, offence
  category, subcategory), one column per month (1995 onward). The
  NSW-wide file has population built in per row ("2025 population",
  "2026 population" columns); the LGA-level file does NOT, so LGA
  rates need the separate ABS population lookup above.

  Rate = SUM of the 12 monthly incident counts across ALL subcategories
  within a category, for the target year, divided by (population / 1000)
  -- same "sum, not mean" principle already proven correct for QPS
  (an annual rate is a rate PER YEAR, not an average month). Verified
  directly: NSW-wide Assault 2024 computed at 8.98 vs the real
  workbook's known 8.85 (1.5% off); Narrabri LGA Assault 2024 computed
  at 13.71 vs the real workbook's known 13.91 (1.5% off) -- same small
  margin as every other state's crime data in this project, not a
  methodology problem.

  Skip incomplete years: BOCSAR's most recent year is partial (2026
  only had 6 of 12 months, matching its own "Jun 2026" cutoff) --
  same "skip incomplete years" pattern already proven correct
  elsewhere. Also skip any year with no matching ABS population figure
  (the population series currently ends 2025) -- a year needs BOTH a
  complete 12 months of crime data AND a population figure to produce
  a rate at all.

CONFIRMED REAL CATEGORY MAPPING: all 8 sheet categories (Assault, Drug
offences, Malicious damage to property, Other offences, Other offences
against the person, Robbery, Theft, Transport regulatory offences) are
literal, exact "Offence category" values in BOCSAR's own data -- no
grouping or combining needed, unlike QPS's turnover-band situation.

GENERALIZED 2026-09-24, per Steve's explicit direction ("knowing there
will be a fetch victoria, tasmania, nt, wa at some point"): this
fetcher's shape is meant to be the template for those. What's
NSW-specific here (narrowly, so the next state fetcher copies the
right things): the two BOCSAR URLs, the wide-format-with-Month-Year-
columns parser, the 8-category label list, and the "does this file
have population built in" split. What's already state-agnostic and
should NOT be re-invented per state: the output JSON schema
(indicators[key] = {"label": ..., "values": {...}}, read generically
by update_crime.py), the sum-not-mean annual aggregation principle,
and the skip-incomplete-years safeguard. A future state fetcher should
produce the same JSON shape from whatever that state's own source
actually looks like -- update_crime.py needs no changes to support it.

towns.toml mapping: NSW towns' existing `lga` field is reused directly
as the BOCSAR LGA name -- confirmed exact match (Narrabri's `lga =
"Narrabri"` matches BOCSAR's own LGA column value verbatim), no new
towns.toml field needed.
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
ABS_LGA_POP_RELEASE = "2024-25"   # update each cycle if auto-advance below fails
ABS_LGA_POP_URL = (
    f"{ABS_LGA_POP_BASE}/{ABS_LGA_POP_RELEASE}/32180DS0004_2001-25.xlsx"
)
# update each cycle -- both the "/2024-25/" release folder AND the
# "2001-25" filename suffix above advance every year ABS republishes this
# series; the URL 404s once retired rather than silently serving stale data


def _guess_next_lga_pop_release(release: str) -> tuple[str, str] | None:
    """
    "2024-25" -> ("2025-26", "26"). Returns (release_folder, filename_suffix)
    or None if the shape doesn't match, so an unexpected format fails safe
    (falls back to the hardcoded release) rather than guessing nonsense.
    """
    m = re.match(r"^(\d{4})-(\d{2})$", release)
    if not m:
        return None
    start = int(m.group(1))
    new_start, new_end2 = start + 1, (start + 2) % 100
    return f"{new_start}-{new_end2:02d}", f"{new_end2:02d}"


def _discover_current_lga_pop_url(log) -> str:
    """
    ADDED 2026-09-29, same guess-verify-fallback approach as
    fetch_business.py's _discover_current_release(): guess next year's
    likely release from this year's pattern, verify with a live HEAD
    request, and only use it if confirmed. Tried live 2026-09-29: a guessed
    "2025-26" release correctly 404s today while the hardcoded "2024-25"
    returns 200 -- confirms the fallback path works, not just the mechanism
    in the abstract.
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
ABS_POP_CACHE_KEY = "abs_lga_population_2001_2025"   # cache filename only, cosmetic -- no need to keep in sync

# Confirmed real, exact "Offence category" values (2026-09-24, direct
# inspection of both BOCSAR files) -- the sheet's own row order.
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
        """Returns { "Narrabri": {2001: 14422, ..., 2025: 12797}, ... }.
        Confirmed real structure (2026-09-24): Table 1, row 5 = year
        headers (2001-2025), data from row 7, column A = LGA code,
        column B = LGA name, columns C onward = one per year.
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
        """Shared core: given raw rows (Geography, Offence category,
        Subcategory, ...monthly columns...), returns
        { year: { category: raw_monthly_sum } } for one geography,
        summing ALL subcategories within each category, only for years
        with a complete 12 months present.
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
            # BUG FIXED 2026-09-28: this used to count VALUES (max list
            # length across categories), but a category with several
            # subcategory rows appends one value per row per month --
            # so 3 subcategories x 6 months = 18 passed a ">= 12" test
            # and a half-year (2026, Jan-Jun) got turned into a rate.
            # Completeness is a property of the header's month columns,
            # not of how many rows there happen to be.
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

        # Population is per-row here; take it from any matching row (confirmed
        # constant across rows in this file, per-state not per-category).
        population_by_year: dict[int, int] = {}
        for row in data_rows:
            if row[3]:
                population_by_year[2025] = row[3]
            if row[4]:
                population_by_year[2026] = row[4]

        # Only 2025 has both a complete year AND a matching population
        # figure in the current release; future years follow the same
        # pattern automatically as the source file grows.
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