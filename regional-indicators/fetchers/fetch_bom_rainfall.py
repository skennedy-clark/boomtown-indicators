"""
regional-indicators/fetchers/fetch_bom_rainfall.py

Fetches annual and seasonal rainfall totals for each town's BOM station,
primarily from the SILO API (Queensland Government), with manually
entered monthly figures as the fallback for stations that SILO does not
carry.

Source:
  SILO Patched Point Dataset
  API:   https://www.longpaddock.qld.gov.au/cgi-bin/silo/PatchedPointDataset.php
  Docs:  https://www.longpaddock.qld.gov.au/silo/

  SILO serves BOM station data over a plain HTTP API (no session
  cookies, no FTP). It requires only an email address as the username;
  no registration is needed.

  Configuration required in config.py:
      SILO_EMAIL = "<email address>"

Request format:
  PatchedPointDataset.php?format=csv&comment=R&station={N}&start=YYYYMMDD&finish=YYYYMMDD&username={email}
  format=csv with comment=R returns daily rainfall only, a smaller
  response than alldata.

CSV layout:
  station,YYYY-MM-DD,daily_rain,daily_rain_source,metadata
  41240,2001-01-01,    0.0,0,"name=HEREWARD"
  41240,2001-01-02,    0.0,0,"latitude= -27.1858"
  ...
  The metadata column carries station information in the first ~8 rows
  and is empty after that.

Source codes in daily_rain_source:
  0  = observed at the target station
  25 = observed at a nearby station
  15 = synthetic / interpolated from the grid
  Codes 0 and 25 are both treated as observed data. Any year in which
  more than 10% of days are code 15 is flagged in the output note.

Station policy (no automatic substitution):
  Continuity with the published record takes priority over automated
  data availability. towns.toml's bom_station is the station used in
  the published workbook history (e.g. Dalby Airport 041522, Moranbah
  Airport 034035), not whichever nearby station is easiest to fetch.
  When the configured station is not in SILO, the fetcher does not
  substitute another SILO station, because that would silently break
  continuity with the years already published. Example: Yarram Airport
  (85151) is an active station in BOM's climate archive but is not in
  SILO's Patched Point Dataset; that is a coverage gap, and a
  different station would give a different series.

  A new town with no published history is a different case: choosing
  the nearest SILO-available station is legitimate because there is no
  continuity to preserve. That choice is made once, manually, when the
  town is added to towns.toml, not at fetch time.

  When the configured station is not in SILO, manual_rainfall_data.toml
  is checked for manually entered monthly figures. These are read from
  BOM's Climate Data Online page in a browser, because that page blocks
  automated access. Manual figures go through the same
  total/summer/winter aggregation as SILO data and are labelled as
  manual in the output. If there are none, the town fails with a
  warning that links to the station's Climate Data Online page.

BOM anonymous FTP is not a source:
  BOM's README at ftp2.bom.gov.au/anon/gen/README states that the
  service carries current forecasts, warnings, observations and charts
  only. Historical station climate records are a separate, paid
  "registered user" product, with only sample data available free.

Aggregation:
  Daily or monthly mm -> monthly totals -> annual total and
  summer/winter split.
  Summer = Jan, Feb, Mar, Oct, Nov, Dec
  Winter = Apr, May, Jun, Jul, Aug, Sep
  Historic average = mean of all complete years in the full record
  (every year with 12 complete months, not only YEAR_START-YEAR_END).
  The same rule applies to SILO and manually entered records.

BOM official average:
  BOM's "Mean rainfall (mm)" annual figure for the station is also read
  from its Climate Averages table (BOM_AVERAGES_URL) and written
  alongside the computed average. Not every station has such a page.

Notes:
  Data comparison for SILO station 41240 (near Dalby), 2024:
  SILO = 972.7 mm; the manually compiled QGSO+BoM xlsx = 947.4 mm. The
  difference (~2.5%) is expected: the manual xlsx had some missing days,
  whereas SILO fills gaps from nearby stations. Station 41240 is not
  the published Dalby station. towns.toml points at 041522 (Dalby
  Airport), which is not in SILO, so this comparison illustrates SILO's
  gap filling and is not guidance for Dalby.

Output:
  cache/silo_rainfall_<station>.csv            raw SILO download
  cache/rainfall/<slug>_bom_rainfall.json      one per town

Website CSVs produced from this data:
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
# BOM's static "Climate Averages" tables. This is a different product
# from the interactive Climate Data Online portal and does not block
# automated requests. {station} is the station number zero-padded to
# six digits. The "Mean rainfall (mm)" row layout is described in
# _fetch_official_historic_average().

# Source codes in the daily_rain_source column.
CODE_SYNTHETIC   = 15   # interpolated/patched; a year is flagged if >10% of days
CODE_MANUAL      = -1   # marker for manually entered data (never counted as synthetic)
# Codes 0 (target station) and 25 (nearby station) are both observations.

PATCH_THRESHOLD  = 0.10   # flag a year if more than this fraction of days are code 15

SUMMER_MONTHS = frozenset({1, 2, 3, 10, 11, 12})
WINTER_MONTHS = frozenset({4, 5, 6, 7, 8, 9})

# Earliest year to fetch. A long record is needed for the historic
# average.
FETCH_FROM_YEAR = 2001

_manual_data_cache: dict | None = None


def _load_manual_rainfall_data() -> dict:
    """Load and cache manual_rainfall_data.toml as
    {town: {(year, month): value_mm}}.

    A missing file means no manual entries exist and is not an error.
    """
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
                # No other station is substituted (see "Station policy" in the
                # module docstring). Manual entries are used instead, if any.
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
                    f"manual_rainfall_data.toml. This is the published "
                    f"station, so no other station is substituted. "
                    f"Open this link in a browser (BOM blocks automated "
                    f"access) to read the monthly figures and add them "
                    f"to manual_rainfall_data.toml: {link}"
                )
                self.result.towns_failed.append(town.name)
                return
        else:
            self.log.info(f"  [{town.name}] Using cached SILO data for station {station}")

        # Parse the SILO CSV into the same structure as manual data.
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
        """Fetch BOM's official 'Mean rainfall (mm)' annual figure for this
        station from its Climate Averages page.

        Returns (value, note). value is None when the figure is
        unavailable (network error, no page for the station, unexpected
        page structure). In that case the caller keeps the Historic
        Average already in the workbook; it is neither blanked nor
        estimated. The note always states what happened.
        """
        # The page name uses the station number zero-padded to six digits
        # (cw_041522.shtml), whereas towns.toml holds the 5-digit form
        # (41522); the unpadded form returns 404. Of the 12 workbook
        # stations, 6 have a page (Dalby, Miles, Moranbah, Narrabri, Roma,
        # Toowoomba). The other six have none (404 when padded as well),
        # so they take the fallback and the returned note says so.
        station_padded = str(station).zfill(6)
        url = BOM_AVERAGES_URL.format(station=station_padded)
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

            # Row layout on the live page (18 cells): [0]=label,
            # [1..12]=Jan..Dec, [13]=Annual, [14]=years of data,
            # [15]=date range, [16..17]=empty plot/map icon cells. Cells are
            # indexed from the start because the number of trailing icon
            # cells is not fixed; indexing from the end would pick up the
            # date-range cell instead of Annual.
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
                f"checking manually: {url}"
            )

    def _monthly_from_manual_data(self, town_name: str) -> dict:
        """Convert this town's manual_rainfall_data.toml entries into the
        {(year, month): [(rain_mm, source_code), ...]} structure that
        _parse_csv() produces, so that _aggregate() handles both sources.

        Each month is represented by a single record with CODE_MANUAL.
        That code is never counted as CODE_SYNTHETIC, so manual entries
        are not flagged as interpolated: they are observed figures
        transcribed from BOM, not grid-interpolated data.
        """
        manual = _load_manual_rainfall_data().get(town_name, {})
        monthly: dict[tuple, list] = defaultdict(list)
        for (year, month), value_mm in manual.items():
            monthly[(year, month)].append((float(value_mm), CODE_MANUAL))
        return monthly

    # ── Download ───────────────────────────────────────────────────────────────

    def _download_station(
        self, station: str, cache_path: Path, town_name: str, email: str
    ) -> tuple[bool, str]:
        """Download the SILO CSV for a station. Returns (success, station)."""
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

            # SILO reports an invalid station as plain text, not as an HTTP error.
            if "Invalid station" in content or "Sorry station" in content:
                self.log.warning(
                    f"  [{town_name}] Station {station} not in SILO: "
                    f"{content[:120].strip()}"
                )
                return False, station

            # The response should be CSV, not HTML.
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
        """Parse a SILO CSV into
          { (year, month): [(rain_mm, source_code), ...] }

        SILO CSV layout (comment=R):
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

            # The first line is the header; check that it has the expected columns.
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
        """Aggregate daily or monthly records into annual totals. The logic
        is the same whether `monthly` came from SILO's daily CSV or from
        manual_rainfall_data.toml's monthly entries.

        Returns:
          annual        - { year: {"total": mm, "summer": mm, "winter": mm} }
          patched_years - set of years where >PATCH_THRESHOLD days have code=15
          historic_avg  - mean annual total across all complete years in record
        """
        # Summarise each (year, month) bucket.
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

            # Skip incomplete years; they are also excluded from the historic
            # average.
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

        # Build the data quality note.
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