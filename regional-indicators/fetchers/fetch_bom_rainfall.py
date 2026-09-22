"""
regional-indicators/fetchers/fetch_bom_rainfall.py
--------------------------------
Fetches annual + seasonal rainfall totals, primarily from the SILO API
(Queensland Government), falling back to manually-entered data for
towns whose TRUE, published BOM station isn't in SILO's dataset.

Source: SILO Patched Point Dataset
API:    https://www.longpaddock.qld.gov.au/cgi-bin/silo/PatchedPointDataset.php
Docs:   https://www.longpaddock.qld.gov.au/silo/

SILO hosts BOM station data through a clean HTTP API — no session cookies, no FTP.
Requires only an email address as username (no registration).

Config required in config.py:
    SILO_EMAIL = "uqsken12@uq.edu.au"

REQUEST FORMAT
  URL: PatchedPointDataset.php?format=csv&comment=R&station={N}&start=YYYYMMDD&finish=YYYYMMDD&username={email}
  format=csv + comment=R returns daily rainfall only (smaller response than alldata)

ACTUAL CSV FORMAT (verified from real response 2026-04-07):
  station,YYYY-MM-DD,daily_rain,daily_rain_source,metadata
  41240,2001-01-01,    0.0,0,"name=HEREWARD"
  41240,2001-01-02,    0.0,0,"latitude= -27.1858"
  ...
  Note: metadata column carries station info in first ~8 rows, then empty

SOURCE CODES for daily_rain_source (verified):
  0  = target station observed (ideal)
  25 = nearby station observed (still real obs, not the target station)
  15 = synthetic / interpolated from grid
  Any year where >10% of days are code=15 is flagged in output notes.
  Codes 0 and 25 are both treated as observed data.

POLICY, CHANGED 2026-09-21 -- NO MORE AUTO-SUBSTITUTION
  Confirmed real, published-record continuity matters more than
  automated data availability: towns.toml's bom_station is meant to be
  the ACTUAL station used in the published workbook history (e.g.
  Dalby Airport 041522, Moranbah Airport 034035), not whichever nearby
  station happens to be easiest to fetch. An earlier version of this
  fetcher auto-substituted a different SILO-available station when the
  configured one wasn't found (via _find_alternative_station, now
  removed from the main flow) -- this silently broke continuity with
  years of already-published data using the true station. Confirmed
  real example: Yarram's "Airport" station (85151) is real and
  currently active per BOM's own climate archive, but simply isn't in
  SILO's curated Patched Point Dataset -- a coverage gap, not a wrong
  number, and auto-substituting a different station would have been
  the wrong fix even though it "worked."

  New towns being added with no prior published history are a
  different case -- picking the nearest SILO-available station for
  those is legitimate (there's no continuity to preserve yet). That
  decision happens once, by hand, when the town is added to
  towns.toml -- not automatically at fetch time.

  When the TRUE station isn't in SILO: checks manual_rainfall_data.toml
  for manually-entered monthly figures (a human reads them off BOM's
  Climate Data Online page directly, since that page blocks automated
  access but not ordinary browsing -- confirmed 2026-09-21). If found,
  those run through the EXACT SAME total/summer/winter aggregation
  logic as SILO data, just source-labelled as manual. If not found,
  fails with a direct clickable link to that station's BOM CDO page.

BOM ANONYMOUS FTP -- CONFIRMED NOT A VIABLE SOURCE (2026-09-21)
  Investigated as an alternative to SILO. BOM's own README at
  ftp2.bom.gov.au/anon/gen/README states this service carries CURRENT
  forecasts/warnings/observations/charts only -- historical station
  climate records are explicitly a separate, PAID "registered user"
  product, with only non-real samples available free. Not a path to
  historical rainfall data. Don't revisit this without a real reason
  to believe it's changed.

AGGREGATION
  Daily/monthly mm → monthly totals → annual total and summer/winter split
  Summer = Jan, Feb, Mar, Oct, Nov, Dec
  Winter = Apr, May, Jun, Jul, Aug, Sep
  Historic average = mean of all complete years in the full record
  (All years with 12 complete months, not just YEAR_START-YEAR_END --
  applies identically whether the record came from SILO or manual entry)

DATA VALIDATION NOTE
  Dalby (41240) 2024: SILO = 972.7mm, manual QGSO+BoM xlsx = 947.4mm
  Difference (~2.5%) is expected — the manual xlsx had some missing days (None values).
  SILO fills gaps from nearby stations; this is the preferred data source.
  NOTE: 41240 was the auto-substituted station under the old policy --
  towns.toml now correctly points at 041522 (Dalby Airport), the real
  published station, which is NOT in SILO. This note is kept for
  historical context, not as current guidance.

Website CSVs produced:
  Environment - total rainfall.csv
  Environment - summer rainfall.csv
  Environment - winter rainfall.csv
"""

from __future__ import annotations

import csv
import json
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR, YEAR_START, YEAR_END

try:
    import requests
except ImportError:
    raise ImportError("pip install requests")

try:
    from bs4 import BeautifulSoup
except ImportError:
    raise ImportError("pip install beautifulsoup4 lxml")

# ── Configuration ──────────────────────────────────────────────────────────────

SILO_BASE        = "https://www.longpaddock.qld.gov.au/cgi-bin/silo"
SILO_STATION_URL = f"{SILO_BASE}/PatchedPointDataset.php"
SILO_SOURCE      = "SILO Patched Point Dataset (Queensland Government / BOM)"
SILO_SOURCE_URL  = "https://www.longpaddock.qld.gov.au/silo/"

MANUAL_SOURCE     = "Manually entered (BOM Climate Data Online)"
MANUAL_DATA_PATH  = Path(__file__).parent.parent / "manual_rainfall_data.toml"
BOM_CDO_URL        = (
    "https://www.bom.gov.au/jsp/ncc/cdio/weatherData/av"
    "?p_nccObsCode=139&p_display_type=dataFile&p_stn_num={station}"
)

BOM_AVERAGES_URL = "https://www.bom.gov.au/climate/averages/tables/cw_{station}.shtml"
# Confirmed genuinely accessible (2026-09-22, verified independently via
# both web_fetch and a raw curl request -- real HTTP 200, real content).
# A DIFFERENT BOM product from the interactive Climate Data Online portal
# that blocks automated access -- this is their static "Climate Averages"
# tables, not subject to the same block. Confirmed real row shape: label
# cell, 12 month cells, Annual cell, years-of-data cell, date-range cell
# (16 cells total for the "Mean rainfall (mm)" row specifically).

# Source codes in daily_rain_source column
CODE_SYNTHETIC   = 15   # interpolated/patched — flag if >10% of year
CODE_MANUAL      = -1   # marker for manually-entered data (never "synthetic")
# Codes 0 and 25 are both real observations (0=target station, 25=nearby station)

PATCH_THRESHOLD  = 0.10   # flag year if >this fraction of days are code=15

SUMMER_MONTHS = frozenset({1, 2, 3, 10, 11, 12})
WINTER_MONTHS = frozenset({4, 5, 6, 7, 8, 9})

# Earliest year to fetch — gives full historic record for average calculation
FETCH_FROM_YEAR = 2001

_manual_data_cache: dict | None = None


def _load_manual_rainfall_data() -> dict:
    """Load and cache manual_rainfall_data.toml -> {town: {(year, month): value_mm}}.
    A missing file just means no manual entries exist yet, not an error."""
    global _manual_data_cache
    if _manual_data_cache is not None:
        return _manual_data_cache

    if not MANUAL_DATA_PATH.exists():
        _manual_data_cache = {}
        return _manual_data_cache

    import tomllib
    with open(MANUAL_DATA_PATH, "rb") as f:
        data = tomllib.load(f)

    result: dict[str, dict] = {}
    for entry in data.get("month", []):
        town = entry.get("town")
        year = entry.get("year")
        month = entry.get("month")
        value = entry.get("value_mm")
        if town is None or year is None or month is None or value is None:
            continue
        result.setdefault(town, {})[(year, month)] = value
    _manual_data_cache = result
    return result


class BOMRainfallFetcher(BaseFetcher):

    SOURCE_NAME      = "bom_rainfall"
    SUPPORTED_STATES = []   # national

    def fetch_all(self):
        try:
            from config import SILO_EMAIL
        except ImportError:
            SILO_EMAIL = None

        if not SILO_EMAIL:
            self.result.add_error(
                "ALL",
                "SILO_EMAIL not set in config.py — add: SILO_EMAIL = 'your.email@uq.edu.au'"
            )
            return

        towns_with_station = [t for t in self.applicable_towns() if t.bom_station]
        towns_without      = [t for t in self.applicable_towns() if not t.bom_station]

        for t in towns_without:
            self.log.warning(f"  [{t.name}] no bom_station in towns.toml — skipping")
            self.result.towns_skipped.append(t.name)

        for town in towns_with_station:
            self._fetch_town(town, SILO_EMAIL)

    # ── Per-town fetch ─────────────────────────────────────────────────────────

    def _fetch_town(self, town, email: str):
        station    = str(town.bom_station)
        cache_path = CACHE_DIR / f"silo_rainfall_{station}.csv"
        source_label = SILO_SOURCE
        source_url = SILO_SOURCE_URL

        if not cache_path.exists() or self.force:
            ok, _ = self._download_station(station, cache_path, town.name, email)
            if not ok:
                # NOT auto-substituting a different station -- continuity
                # with the published record matters more than automated
                # availability (see module docstring). Check manual
                # entries instead.
                monthly = self._monthly_from_manual_data(town.name)
                if monthly:
                    self.log.info(
                        f"  [{town.name}] Station {station} not in SILO -- "
                        f"using {len(monthly)} manually-entered month(s)."
                    )
                    source_label = MANUAL_SOURCE
                    source_url = BOM_CDO_URL.format(station=station)
                    self._finish(town, station, monthly, source_label, source_url)
                    return

                link = BOM_CDO_URL.format(station=station)
                self.log.warning(
                    f"  [{town.name}] Station {station} not in SILO Patched Point "
                    f"Dataset, and no manual entries found in "
                    f"manual_rainfall_data.toml. This is the TRUE published "
                    f"station -- not auto-substituting a different one. "
                    f"Open this in a browser (not automatable, BOM blocks "
                    f"scraping here) to read the monthly figures and add them "
                    f"to manual_rainfall_data.toml: {link}"
                )
                self.result.towns_failed.append(town.name)
                return
        else:
            self.log.info(f"  [{town.name}] Using cached SILO data for station {station}")

        # Parse SILO CSV → same aggregation path as manual data
        monthly = self._parse_csv(cache_path, town.name)
        if monthly is None:
            self.result.towns_failed.append(town.name)
            return

        self._finish(town, station, monthly, source_label, source_url)

    def _finish(self, town, station: str, monthly: dict, source_label: str, source_url: str):
        annual, patched_years, historic_avg = self._aggregate(monthly, town.name)
        if not annual:
            self.result.towns_failed.append(town.name)
            return
        bom_official_avg, bom_official_note = self._fetch_official_historic_average(station, town.name)
        self._write_cache(
            town, station, annual, patched_years, historic_avg, source_label, source_url,
            bom_official_avg, bom_official_note,
        )

    def _fetch_official_historic_average(self, station: str, town_name: str) -> tuple[float | None, str]:
        """Fetch BOM's own official 'Mean rainfall (mm) Annual' figure
        for this station from their Climate Averages page -- confirmed
        genuinely accessible (2026-09-22), a different BOM product from
        the interactive portal that blocks automated access.

        Returns (value, note). value is None if genuinely unavailable
        (network error, station not on this page, unexpected page
        structure) -- the caller falls back to keeping whatever's
        already in the workbook rather than blanking it or guessing,
        per Steve's explicit instruction, but the note always explains
        what happened so it's visible either way.
        """
        url = BOM_AVERAGES_URL.format(station=station)
        try:
            resp = requests.get(url, timeout=30,
                                 headers={"User-Agent": "boomtown-indicators/1.0 (UQ research pipeline)"})
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "lxml")

            target_row = None
            for row in soup.find_all("tr"):
                first_cell = row.find("td")
                if first_cell and "Mean rainfall" in first_cell.get_text():
                    target_row = row
                    break

            if target_row is None:
                return None, (
                    f"BOM Climate Averages page for station {station} loaded, but no "
                    f"'Mean rainfall' row was found -- page structure may have changed. "
                    f"Kept the existing Historic Average; worth checking manually: {url}"
                )

            cells = target_row.find_all("td")
            if len(cells) < 15:
                return None, (
                    f"BOM Climate Averages page for station {station} loaded, but the "
                    f"rainfall row has an unexpected shape ({len(cells)} cells, expected "
                    f"18: label, 12 months, Annual, years, dates, plot, map). Kept the "
                    f"existing Historic Average; worth checking manually: {url}"
                )

            # Confirmed real row shape (2026-09-22, verified against the live
            # page, not just an isolated snippet): [0]=label, [1..12]=Jan..Dec,
            # [13]=Annual, [14]=years-of-data, [15]=date-range, [16..17]=empty
            # plot/map icon cells. Indexed from the START, not the end -- an
            # earlier version used cells[-3]/cells[-2] and silently grabbed the
            # date-range cell instead of Annual, because the real page has 18
            # cells (with trailing plot/map icons) not the 16 an isolated test
            # snippet assumed. Fixed positions from the start are stable
            # regardless of how many decorative cells follow.
            annual_text = cells[13].get_text(strip=True)
            annual_value = float(annual_text)
            years_text = cells[14].get_text(strip=True)
            self.log.info(
                f"  [{town_name}] BOM official historic average: {annual_value} mm "
                f"(station {station}, {years_text} years of record)"
            )
            return annual_value, (
                f"BOM's own official Mean Annual Rainfall for station {station}, "
                f"{years_text} years of record: {url}"
            )

        except Exception as exc:
            return None, (
                f"Could not fetch BOM's official historic average for station "
                f"{station} ({exc}). Kept the existing Historic Average; worth "
                f"checking manually: {BOM_AVERAGES_URL.format(station=station)}"
            )

    def _monthly_from_manual_data(self, town_name: str) -> dict:
        """Convert manual_rainfall_data.toml entries for this town into the
        same {(year, month): [(rain_mm, source_code), ...]} shape
        _parse_csv() produces from SILO CSV, so _aggregate() can be
        reused completely unchanged. Uses a single synthetic "day" per
        month at CODE_MANUAL (never counted as CODE_SYNTHETIC, so
        manual entries never get flagged as interpolated/patched --
        they're real numbers a human read directly, not the grid-
        interpolated data that flag is for)."""
        manual = _load_manual_rainfall_data().get(town_name, {})
        monthly: dict[tuple, list] = defaultdict(list)
        for (year, month), value_mm in manual.items():
            monthly[(year, month)].append((float(value_mm), CODE_MANUAL))
        return monthly

    # ── Download ───────────────────────────────────────────────────────────────

    def _download_station(
        self, station: str, cache_path: Path, town_name: str, email: str
    ) -> tuple[bool, str]:
        """Download SILO CSV for a station. Returns (success, station_used)."""
        params = {
            "format":   "csv",
            "comment":  "R",
            "station":  station,
            "start":    f"{FETCH_FROM_YEAR}0101",
            "finish":   f"{YEAR_END}1231",
            "username": email,
        }
        self.log.info(f"  [{town_name}] Downloading SILO station {station}...")
        try:
            resp = requests.get(
                SILO_STATION_URL,
                params=params,
                timeout=60,
                headers={"User-Agent": "boomtown-indicators/1.0 (UQ research pipeline)"},
            )
            resp.raise_for_status()
            content = resp.text

            # SILO returns a plain-text error (not HTTP error) for invalid stations
            if "Invalid station" in content or "Sorry station" in content:
                self.log.warning(
                    f"  [{town_name}] Station {station} not in SILO: "
                    f"{content[:120].strip()}"
                )
                return False, station

            # Sanity check — should be CSV not HTML
            if "<html" in content.lower()[:100]:
                self.log.error(f"  [{town_name}] Got HTML response — unexpected error")
                return False, station

            cache_path.write_text(content, encoding="utf-8")
            kb = cache_path.stat().st_size // 1024
            lines = len(content.splitlines())
            self.log.info(f"  [{town_name}] Saved {kb} KB ({lines} lines)")
            return True, station

        except Exception as exc:
            self.log.error(f"  [{town_name}] SILO download failed: {exc}")
            return False, station

    # ── Parse ──────────────────────────────────────────────────────────────────

    def _parse_csv(self, path: Path, town_name: str) -> dict | None:
        """
        Parse SILO CSV into:
          { (year, month): [(rain_mm, source_code), ...] }

        Actual SILO CSV format (comment=R):
          station,YYYY-MM-DD,daily_rain,daily_rain_source,metadata
          41240,2001-01-01,    0.0,0,"name=HEREWARD"
          ...
        """
        try:
            text  = path.read_text(encoding="utf-8", errors="replace")
            lines = text.splitlines()

            if not lines:
                self.log.error(f"  [{town_name}] Empty SILO CSV")
                return None

            # The first line IS the header — verify it looks right
            header = lines[0].strip()
            if "YYYY-MM-DD" not in header and "daily_rain" not in header:
                self.log.error(
                    f"  [{town_name}] Unexpected CSV header: {header[:100]}"
                )
                return None

            reader  = csv.DictReader(lines)
            monthly: dict[tuple, list] = defaultdict(list)
            skipped = 0

            for row in reader:
                try:
                    date_str = row["YYYY-MM-DD"].strip()
                    # date format: YYYY-MM-DD
                    year  = int(date_str[:4])
                    month = int(date_str[5:7])
                    rain  = float(row["daily_rain"].strip())
                    src   = int(row["daily_rain_source"].strip())
                    monthly[(year, month)].append((rain, src))
                except (ValueError, KeyError):
                    skipped += 1
                    continue

            if skipped:
                self.log.debug(f"  [{town_name}] Skipped {skipped} unparseable rows")

            if not monthly:
                self.log.error(f"  [{town_name}] No data rows parsed from SILO CSV")
                return None

            years = sorted(set(k[0] for k in monthly))
            self.log.info(
                f"  [{town_name}] Parsed {len(monthly)} month-records "
                f"({years[0]}–{years[-1]})"
            )
            return monthly

        except Exception as exc:
            self.log.error(f"  [{town_name}] CSV parse error: {exc}", exc_info=True)
            return None

    # ── Aggregate ──────────────────────────────────────────────────────────────

    def _aggregate(
        self,
        monthly: dict,
        town_name: str,
    ) -> tuple[dict, set, float | None]:
        """
        Aggregate daily/monthly records into annual totals -- identical
        logic regardless of whether `monthly` came from SILO's daily CSV
        or manual_rainfall_data.toml's monthly entries.

        Returns:
          annual        – { year: {"total": mm, "summer": mm, "winter": mm} }
          patched_years – set of years where >PATCH_THRESHOLD days have code=15
          historic_avg  – mean annual total across ALL complete years in record
        """
        # Summarise each (year, month) bucket
        month_totals: dict[tuple, dict] = {}
        for (year, month), days in monthly.items():
            total_mm     = sum(d[0] for d in days)
            synthetic    = sum(1 for d in days if d[1] == CODE_SYNTHETIC)
            total_days   = len(days)
            month_totals[(year, month)] = {
                "mm":          round(total_mm, 1),
                "synthetic":   synthetic,
                "total_days":  total_days,
            }

        annual:         dict[int, dict] = {}
        patched_years:  set[int]        = set()
        all_year_totals: list[float]    = []

        all_years = sorted(set(k[0] for k in month_totals))

        for year in all_years:
            months = {m: month_totals[(year, m)]
                      for m in range(1, 13)
                      if (year, m) in month_totals}

            # Skip incomplete years (don't include in historic average either)
            if len(months) < 12:
                continue

            total_mm  = sum(m["mm"] for m in months.values())
            summer_mm = sum(m["mm"] for mo, m in months.items() if mo in SUMMER_MONTHS)
            winter_mm = sum(m["mm"] for mo, m in months.items() if mo in WINTER_MONTHS)

            total_synthetic = sum(m["synthetic"]  for m in months.values())
            total_days      = sum(m["total_days"] for m in months.values())
            patch_frac      = total_synthetic / total_days if total_days else 0

            if patch_frac > PATCH_THRESHOLD:
                patched_years.add(year)

            all_year_totals.append(total_mm)
            annual[year] = {
                "total":  round(total_mm, 1),
                "summer": round(summer_mm, 1),
                "winter": round(winter_mm, 1),
            }

        historic_avg = (
            round(sum(all_year_totals) / len(all_year_totals), 1)
            if all_year_totals else None
        )

        if patched_years:
            self.log.warning(
                f"  [{town_name}] {len(patched_years)} years >10% synthetic data: "
                f"{sorted(patched_years)}"
            )

        return annual, patched_years, historic_avg

    # ── Write cache ────────────────────────────────────────────────────────────

    def _write_cache(
        self,
        town,
        station: str,
        annual:        dict,
        patched_years: set,
        historic_avg:  float | None,
        source_label:  str,
        source_url:    str,
        bom_official_avg:  float | None = None,
        bom_official_note: str = "",
    ):
        in_range = {yr: v for yr, v in annual.items() if YEAR_START <= yr <= YEAR_END}

        if not in_range:
            self.log.warning(
                f"  [{town.name}] no complete years in range {YEAR_START}–{YEAR_END}"
            )
            self.result.towns_failed.append(town.name)
            return

        # Build data quality note
        patch_note = ""
        flagged_in_range = sorted(y for y in patched_years if YEAR_START <= y <= YEAR_END)
        if flagged_in_range:
            patch_note = (
                f" NOTE: >10% of daily values are SILO-interpolated (not direct station "
                f"observations) in year(s): {flagged_in_range}."
            )

        manual_note = ""
        if source_label == MANUAL_SOURCE:
            manual_note = (
                f" Manually entered from BOM Climate Data Online (station {station} "
                f"is not in SILO's Patched Point Dataset) -- verify against "
                f"manual_rainfall_data.toml's entered_by/entered_date for provenance."
            )

        out = {
            "town":            town.name,
            "state":           town.state,
            "bom_station":     station,
            "source":          source_label,
            "source_url":      source_url,
            "note": (
                f"Rainfall from {source_label}, station {station}, aggregated to "
                f"annual totals. Summer = Jan–Mar + Oct–Dec; Winter = Apr–Sep. "
                f"Historic average (self-computed) = mean of {len(annual)} complete "
                f"years in record (from {min(annual)} to {max(annual)})."
                f"{manual_note}{patch_note}"
            ),
            "historic_avg_mm": historic_avg,
            "bom_official_historic_avg_mm":  bom_official_avg,
            "bom_official_historic_avg_note": bom_official_note,
            "indicators": {
                "rainfall":        {str(yr): v["total"]  for yr, v in in_range.items()},
                "rainfall_summer": {str(yr): v["summer"] for yr, v in in_range.items()},
                "rainfall_winter": {str(yr): v["winter"] for yr, v in in_range.items()},
            },
        }

        out_dir  = CACHE_DIR / "rainfall"
        out_dir.mkdir(exist_ok=True)
        out_path = out_dir / f"{town.slug}_bom_rainfall.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        latest_yr = max(in_range)
        self.log.info(
            f"  {town.name} (station {station}, {source_label}): "
            f"{len(in_range)} years in range, "
            f"latest ({latest_yr}) = {in_range[latest_yr]['total']} mm, "
            f"historic avg = {historic_avg} mm"
        )
        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = BOMRainfallFetcher().run()
    sys.exit(0 if result.success else 1)