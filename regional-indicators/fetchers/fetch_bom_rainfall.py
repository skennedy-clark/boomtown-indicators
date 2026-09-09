"""
fetchers/fetch_bom_rainfall.py
--------------------------------
Fetches annual rainfall totals from the SILO API (Queensland Government).

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

INVALID STATIONS
  Some station numbers are not in the SILO Patched Point Dataset.
  Confirmed invalid (2026-04-07): 42104 (Tara/Woodlea), 41559 (Goondiwindi WTP),
  34035 (Moranbah Airport), 85151 (Yarram).
  For these towns the fetcher attempts a SILO name search to find an alternative
  station number in the SILO dataset, then retries.
  If still not found, the town is marked failed with a clear message.
  It does NOT fall back to DataDrill (grid interpolation) as that would produce
  data from a different source that cannot be compared to booklet historical data
  without disclosure — park it instead.

AGGREGATION
  Daily mm → monthly totals → annual total and summer/winter split
  Summer = Jan, Feb, Mar, Oct, Nov, Dec
  Winter = Apr, May, Jun, Jul, Aug, Sep
  Historic average = mean of all complete years in the full SILO record
  (All years with 12 complete months, not just YEAR_START–YEAR_END)

DATA VALIDATION NOTE
  Dalby (41240) 2024: SILO = 972.7mm, manual QGSO+BoM xlsx = 947.4mm
  Difference (~2.5%) is expected — the manual xlsx had some missing days (None values).
  SILO fills gaps from nearby stations; this is the preferred data source.

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

# ── Configuration ──────────────────────────────────────────────────────────────

SILO_BASE        = "https://www.longpaddock.qld.gov.au/cgi-bin/silo"
SILO_STATION_URL = f"{SILO_BASE}/PatchedPointDataset.php"
SILO_SOURCE      = "SILO Patched Point Dataset (Queensland Government / BOM)"
SILO_SOURCE_URL  = "https://www.longpaddock.qld.gov.au/silo/"

# Source codes in daily_rain_source column
CODE_SYNTHETIC   = 15   # interpolated/patched — flag if >10% of year
# Codes 0 and 25 are both real observations (0=target station, 25=nearby station)

PATCH_THRESHOLD  = 0.10   # flag year if >this fraction of days are code=15

SUMMER_MONTHS = frozenset({1, 2, 3, 10, 11, 12})
WINTER_MONTHS = frozenset({4, 5, 6, 7, 8, 9})

# Earliest year to fetch — gives full historic record for average calculation
FETCH_FROM_YEAR = 2001


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

        if not cache_path.exists() or self.force:
            ok, used_station = self._download_station(station, cache_path, town.name, email)
            if not ok:
                # Try SILO name search for an alternative station number
                alt_station = self._find_alternative_station(town.name, email)
                if alt_station and alt_station != station:
                    self.log.info(f"  [{town.name}] Retrying with alternative station {alt_station}")
                    alt_cache = CACHE_DIR / f"silo_rainfall_{alt_station}.csv"
                    ok, used_station = self._download_station(alt_station, alt_cache, town.name, email)
                    if ok:
                        cache_path = alt_cache
                        station    = alt_station

                if not ok:
                    self.log.warning(
                        f"  [{town.name}] Station {town.bom_station} not in SILO Patched Point Dataset. "
                        f"No valid alternative found. "
                        f"Check https://www.longpaddock.qld.gov.au/cgi-bin/silo/PatchedPointDataset.php"
                        f"?format=name&nameFrag={town.name.lower().split()[0]}"
                    )
                    self.result.towns_failed.append(town.name)
                    return
        else:
            self.log.info(f"  [{town.name}] Using cached SILO data for station {station}")

        # Parse → aggregate → write
        monthly = self._parse_csv(cache_path, town.name)
        if monthly is None:
            self.result.towns_failed.append(town.name)
            return

        annual, patched_years, historic_avg = self._aggregate(monthly, town.name)
        if not annual:
            self.result.towns_failed.append(town.name)
            return

        self._write_cache(town, station, annual, patched_years, historic_avg)

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

    def _find_alternative_station(self, town_name: str, email: str) -> str | None:
        """
        Search SILO by name fragment to find a station number that is in the
        Patched Point Dataset.
        """
        # Use first word of town name as search fragment (e.g. "Goondiwindi" not "Goondiwindi WTP")
        frag = town_name.lower().split()[0]
        url  = f"{SILO_STATION_URL}?format=name&nameFrag={frag}"
        try:
            resp = requests.get(url, timeout=20,
                                headers={"User-Agent": "boomtown-indicators/1.0"})
            resp.raise_for_status()
            lines = [l.strip() for l in resp.text.splitlines() if l.strip()]
            if lines:
                self.log.info(
                    f"  [{town_name}] SILO name search for '{frag}' returned "
                    f"{len(lines)} results:"
                )
                for line in lines[:5]:
                    self.log.info(f"    {line}")
                # Return the first result's station number (pipe-delimited)
                first = lines[0].split("|")[0].strip()
                if first.isdigit():
                    return first
        except Exception as exc:
            self.log.warning(f"  [{town_name}] SILO name search failed: {exc}")
        return None

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
        Aggregate daily records into annual totals.

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

        original_station = str(town.bom_station)
        station_note = ""
        if station != original_station:
            station_note = (
                f" Data retrieved from alternative SILO station {station} "
                f"(configured station {original_station} not in SILO dataset)."
            )

        out = {
            "town":            town.name,
            "state":           town.state,
            "bom_station":     station,
            "source":          SILO_SOURCE,
            "source_url":      SILO_SOURCE_URL,
            "note": (
                f"Daily rainfall from SILO station {station}, aggregated to annual totals. "
                f"Summer = Jan–Mar + Oct–Dec; Winter = Apr–Sep. "
                f"Historic average = mean of {len(annual)} complete years in SILO record "
                f"(from {min(annual)} to {max(annual)}). "
                f"SILO source codes: 0=station obs, 25=nearby station obs, 15=interpolated."
                f"{station_note}{patch_note}"
            ),
            "historic_avg_mm": historic_avg,
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
            f"  {town.name} (station {station}): "
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