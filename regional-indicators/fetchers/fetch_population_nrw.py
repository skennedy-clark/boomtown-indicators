"""
regional-indicators/fetchers/fetch_population_nrw.py
------------------------------------------------------
Fetches non-resident worker (NRW) population data for Surat Basin and
Bowen Basin resource-region towns, at TWO geography levels:

  - LGA-level "Non-resident workers on-shift" (full 2008-2025 series) --
    maps to the LGA-section rows in the workbook (Isaac, Maranoa,
    Western Downs).
  - LGA + selected-UCL "Full-time equivalent (FTE)" estimates (2024-2025
    only) -- maps to the per-town UCL-section NRW rows.

Source: QGSO Surat Basin / Bowen Basin population reports.
Confirmed real URLs (Surat Basin, from live site, 2026-09-15):
  NRW on-shift, LGA, 2008-2025:
    https://www.qgso.qld.gov.au/issues/6606/surat-basin-population-report-tables-non-resident-workers-on-shift-local-government-area-lga-2008-2025.xlsx
  FTE population estimates, LGA + selected UCLs, 2024-2025:
    https://www.qgso.qld.gov.au/issues/6606/surat-basin-population-report-tables-full-time-equivalent-fte-population-estimates-local-government-area-lga-selected-urban-centres-localities-ucls-2024-2025.xlsx
  (A third file, Worker Accommodation Village bed capacity, exists at
  the same issue number but isn't consumed here -- not a currently
  tracked indicator. Worth a TODO note if that becomes relevant later.)

Confirmed real URLs (Bowen Basin, from live site, 2026-09-15):
  NRW on-shift, LGA, 2006-2025 (note: different start year than Surat's
  2008 -- year columns are detected dynamically, not hardcoded, so this
  doesn't need special-casing):
    https://www.qgso.qld.gov.au/issues/3341/bowen-basin-population-report-tables-non-resident-workers-on-shift-local-government-area-lga-2006-2025.xlsx
  FTE population estimates, LGA + selected UCLs, 2024-2025:
    https://www.qgso.qld.gov.au/issues/3341/bowen-basin-population-report-tables-full-time-equivalent-fte-population-estimates-local-government-area-lga-selected-urban-centres-localities-ucls-2024-2025.xlsx
  (Same third WAV-bed-capacity file exists here too, same "not tracked
  yet" note applies.)

Confirms what the module docstring below originally only guessed at:
issue numbers are per-region AND per-report, not shared or predictable
-- Surat's tables are issue 6606, Bowen's are issue 3341, and neither
matches either region's own PDF report issue number (3186 for Surat,
3366 for Bowen). Never assume a new issue number by incrementing a
known one.

*** PARSING NOT YET VALIDATED against a real downloaded file -- column
header search is defensive (searches for header row rather than
assuming fixed position, same approach used elsewhere in this project),
but the actual column names/order in these specific files haven't been
confirmed. First real run against an actual downloaded file is the
real test. ***

Writes cache/population/{slug}_population_nrw.json:
  {
    "town": ..., "lga": ...,
    "lga_nrw_on_shift_by_year": {"2008": ..., ..., "2025": ...},
    "ucl_nrw_latest": {"year": 2025, "value": ...} or null if no
        UCL-level match for this town in the FTE file
  }

Downstream wiring (writing these into the workbook) is deliberately NOT
built yet -- see TODO.md. Better to confirm this fetcher's parsing
against real files first than build the write side against unverified
column assumptions too.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher

try:
    import requests
    import openpyxl
except ImportError:
    raise ImportError("pip install requests openpyxl")


# ── Configuration ──────────────────────────────────────────────────────────

REGIONS = {
    "surat_basin": {
        "theme_url": (
            "https://www.qgso.qld.gov.au/statistics/theme/population/"
            "non-resident-population-queensland-resource-regions/surat-basin"
        ),
        "lgas": ["Western Downs", "Maranoa", "Toowoomba"],  # confirmed from
                                                              # real file: Toowoomba
                                                              # LGA has its own row
                                                              # in this Surat Basin
                                                              # data, not just
                                                              # Western Downs/Maranoa
        "nrw_lga_url": (
            "https://www.qgso.qld.gov.au/issues/6606/"
            "surat-basin-population-report-tables-non-resident-workers-"
            "on-shift-local-government-area-lga-2008-2025.xlsx"
        ),
        "fte_lga_ucl_url": (
            "https://www.qgso.qld.gov.au/issues/6606/"
            "surat-basin-population-report-tables-full-time-equivalent-"
            "fte-population-estimates-local-government-area-lga-selected-"
            "urban-centres-localities-ucls-2024-2025.xlsx"
        ),
    },
    "bowen_basin": {
        "theme_url": (
            "https://www.qgso.qld.gov.au/statistics/theme/population/"
            "non-resident-population-queensland-resource-regions/"
            "bowen-galilee-basins"
        ),
        "lgas": ["Isaac"],
        "nrw_lga_url": (
            "https://www.qgso.qld.gov.au/issues/3341/"
            "bowen-basin-population-report-tables-non-resident-workers-"
            "on-shift-local-government-area-lga-2006-2025.xlsx"
        ),
        "fte_lga_ucl_url": (
            "https://www.qgso.qld.gov.au/issues/3341/"
            "bowen-basin-population-report-tables-full-time-equivalent-"
            "fte-population-estimates-local-government-area-lga-selected-"
            "urban-centres-localities-ucls-2024-2025.xlsx"
        ),
    },
}


class QGSOPopulationNRWFetcher(BaseFetcher):

    SOURCE_NAME      = "population_nrw"
    SUPPORTED_STATES = ["QLD"]

    def fetch_all(self):
        region_data: dict[str, dict] = {}

        for region_key, region in REGIONS.items():
            towns_in_region = [
                t for t in self.applicable_towns() if t.lga in region["lgas"]
            ]
            if not towns_in_region:
                continue

            nrw_lga = self._fetch_and_parse_nrw_lga(region_key, region)
            fte_lga_ucl = self._fetch_and_parse_fte(region_key, region)

            if not nrw_lga and not fte_lga_ucl:
                self.result.add_error(
                    region_key,
                    f"Could not get either NRW-LGA or FTE-LGA-UCL data for "
                    f"{region_key}. See earlier warnings for the specific "
                    f"URL(s) that failed."
                )
                continue

            region_data[region_key] = {"nrw_lga": nrw_lga, "fte_lga_ucl": fte_lga_ucl}

        for town in self.applicable_towns():
            region_key = next(
                (k for k, r in REGIONS.items() if town.lga in r["lgas"]), None
            )
            if region_key is None or region_key not in region_data:
                continue
            self._extract_town(town, region_key, region_data[region_key])

    # ── Fetch + parse: LGA-level NRW on-shift, full history ─────────────────

    def _fetch_and_parse_nrw_lga(self, region_key: str, region: dict) -> dict:
        url = region["nrw_lga_url"]
        if not url:
            url = self._scrape_for_url(region_key, region["theme_url"], "non-resident-workers-on-shift")
        if not url:
            self.result.add_warning(
                region_key,
                f"No NRW-LGA URL (confirmed or scraped) for {region_key} -- "
                f"manual fix: visit {region['theme_url']}, open the "
                f"'population report (table)' tab, download the "
                f"'Non-resident workers on-shift LGA' file, save as "
                f"cache/{region_key}_nrw_lga.xlsx"
            )
            return {}

        path = self.download(url, f"{region_key}_nrw_lga", suffix=".xlsx")
        if not path:
            return {}

        return self._parse_multi_year_lga_sheet(path, value_label_hint="nrw")

    # ── Fetch + parse: LGA + selected UCL FTE, 2024-2025 ─────────────────────

    def _fetch_and_parse_fte(self, region_key: str, region: dict) -> dict:
        url = region["fte_lga_ucl_url"]
        if not url:
            url = self._scrape_for_url(region_key, region["theme_url"], "full-time-equivalent")
        if not url:
            self.result.add_warning(
                region_key,
                f"No FTE-LGA-UCL URL (confirmed or scraped) for {region_key} -- "
                f"manual fix: visit {region['theme_url']}, open the "
                f"'population report (table)' tab, download the FTE "
                f"population estimates file, save as "
                f"cache/{region_key}_fte_lga_ucl.xlsx"
            )
            return {}

        path = self.download(url, f"{region_key}_fte_lga_ucl", suffix=".xlsx")
        if not path:
            return {}

        return self._parse_latest_year_sheet(path, value_label_hint="fte")

    def _scrape_for_url(self, region_key: str, theme_url: str, slug_fragment: str) -> str | None:
        """Fallback for when a direct URL isn't hardcoded (Bowen Basin
        currently) or a hardcoded one 404s (issue number rolled over).
        Unverified against the live site -- the direct Surat Basin URLs
        above were confirmed by a human visiting the page; this scrape
        has not been. If the theme page's file links are loaded via
        JavaScript rather than present in the raw HTML, this will find
        nothing no matter how correct the regex is -- worth checking
        "view page source" vs. the visible tab content if this keeps
        failing."""
        try:
            resp = requests.get(theme_url, timeout=30)
            resp.raise_for_status()
            region_word = region_key.split("_")[0]
            pattern = (
                rf'https?://www\.qgso\.qld\.gov\.au/issues/\d+/'
                rf'{region_word}-basin-population-report-tables-'
                rf'[^"\s]*{slug_fragment}[^"\s]*\.xlsx'
            )
            urls = re.findall(pattern, resp.text, re.IGNORECASE)
            if urls:
                self.log.info(f"  [{region_key}] Found URL on page: {urls[0]}")
                return urls[0]
        except Exception as exc:
            self.log.warning(f"  [{region_key}] Page scrape failed: {exc}")
        return None

    # ── Parsers ────────────────────────────────────────────────────────────

    def _find_lga_and_year_rows(self, ws, max_search_rows: int = 15):
        """QGSO NRW/FTE files put the 'LGA' column label and the actual
        year headers on TWO SEPARATE rows (confirmed from a real file):
          Row N:   LGA(a) | Non-resident workers on-shift(b) | ...   <- group label row
          Row N+1: None   | 2008 | 2009 | 2010 | ...                  <- real year row
        Searching for years in the same row as 'LGA' (what the earlier
        version of this file did) finds nothing, since the LGA row's
        other cells are a group title, not individual years.

        Returns (lga_row_idx, lga_col, year_row_idx, {col: year}), or
        (None, None, None, {}) if the LGA label can't be found at all.
        year_row_idx/year_cols can still be empty even if lga_row_idx
        is found, if no year row turns up within the search window --
        that's a real "couldn't find years" case, not this bug.
        """
        lga_row_idx = None
        lga_col = None
        for r in range(1, min(max_search_rows, ws.max_row) + 1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                # Require the cell to START WITH "lga" and be short --
                # a real header cell is "LGA" or "LGA(a)", not a full
                # sentence that happens to mention "(LGA)" as an
                # abbreviation (confirmed: the report's own title row
                # does exactly that and was wrongly matched before this
                # fix -- "contains lga anywhere" is too loose).
                if isinstance(v, str) and v.strip().lower().startswith("lga") and len(v.strip()) <= 15:
                    lga_row_idx, lga_col = r, c
                    break
            if lga_row_idx:
                break

        if lga_row_idx is None:
            return None, None, None, {}

        # Year row: search from the LGA row itself forward a few rows --
        # covers both "years on the same row" and "years on the next
        # row" layouts without needing to know which in advance.
        for r in range(lga_row_idx, min(lga_row_idx + 4, ws.max_row) + 1):
            year_cols = {}
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                year = None
                if isinstance(v, (int, float)) and 1990 <= v <= 2100:
                    year = int(v)
                elif isinstance(v, str) and v.strip().isdigit() and 1990 <= int(v.strip()) <= 2100:
                    year = int(v.strip())
                elif hasattr(v, "year") and isinstance(getattr(v, "year", None), int):
                    # Excel sometimes stores a year header as an actual
                    # date (e.g. 1/07/2008) -- openpyxl hands that back
                    # as a datetime, not a plain number.
                    if 1990 <= v.year <= 2100:
                        year = v.year
                if year:
                    year_cols[c] = year
            if year_cols:
                return lga_row_idx, lga_col, r, year_cols

        return lga_row_idx, lga_col, None, {}

    def _parse_multi_year_lga_sheet(self, path: Path, value_label_hint: str) -> dict:
        """Return {lga_name: {year_str: value}} from a sheet shaped like
        LGA(a) | Non-resident workers on-shift(b)
               | 2008 | 2009 | ... | 2025
        — LGA-label row and year row are found separately, not assumed
        to be the same row (see _find_lga_and_year_rows). Rows with no
        numeric value in any year column (e.g. a "— persons —" units
        label row) are naturally skipped, not string-matched -- more
        robust than matching a specific em-dash character.
        """
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.worksheets[0]

            lga_row_idx, lga_col, year_row_idx, year_cols = self._find_lga_and_year_rows(ws)
            if lga_row_idx is None:
                self.log.error(f"Could not find an 'LGA' label anywhere in {path.name}")
                return {}
            if not year_cols:
                self.log.error(f"Found 'LGA' label but no year row nearby in {path.name}")
                return {}

            result = {}
            for r in range(year_row_idx + 1, ws.max_row + 1):
                lga_val = ws.cell(r, lga_col).value
                if not lga_val:
                    continue
                year_vals = {}
                for c, yr in year_cols.items():
                    v = ws.cell(r, c).value
                    if isinstance(v, (int, float)):
                        year_vals[str(yr)] = int(v)
                if year_vals:
                    result[str(lga_val).strip()] = year_vals

            self.log.info(f"  Parsed {len(result)} LGA entries, {len(year_cols)} years, from {path.name}")
            return result

        except Exception as exc:
            self.log.error(f"Multi-year LGA sheet parse error ({path.name}): {exc}", exc_info=True)
            return {}

    def _parse_latest_year_sheet(self, path: Path, value_label_hint: str) -> dict:
        """Return {ucl_name: {"lga": ..., "year": ..., "value": ...}}
        -- the "value" is specifically the Non-resident workers on-shift
        figure for the latest year, not ERP or the combined FTE figure.

        Confirmed real structure (differs from the simpler LGA-only
        file): a 2-row header where row N has LGA(a) | Location(b) |
        UCL(a) | 2024 | (blank) | (blank) | 2025 | (blank) | (blank),
        and row N+1 sub-labels each year's 3-column group as Estimated
        resident population | Non-resident workers on-shift | FTE
        population estimate. Matching on "Location" instead of "UCL"
        (an earlier version of this code did) finds the wrong column --
        Location holds broad categories ("In town", "Rural areas"), UCL
        holds the actual place names ("Injune", "Roma", "Toowoomba").
        LGA and Location cells use Excel's usual "blank means same as
        the row above" convention (only the first row of each LGA group
        has a value) -- LGA is forward-filled here for context; not
        needed for matching, since matching is by UCL name.
        """
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.worksheets[0]

            label_row, lga_col = None, None
            for r in range(1, min(15, ws.max_row) + 1):
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and v.strip().lower().startswith("lga") and len(v.strip()) <= 15:
                        label_row, lga_col = r, c
                        break
                if label_row:
                    break
            if label_row is None:
                self.log.error(f"Could not find an 'LGA' label row in {path.name}")
                return {}

            ucl_col = None
            for c in range(1, ws.max_column + 1):
                v = ws.cell(label_row, c).value
                if isinstance(v, str) and v.strip().lower().startswith("ucl") and len(v.strip()) <= 15:
                    ucl_col = c
                    break
            if ucl_col is None:
                self.log.error(f"Could not find a 'UCL' column in {path.name} (found LGA row but no UCL column)")
                return {}

            year_cols = {}
            for c in range(1, ws.max_column + 1):
                v = ws.cell(label_row, c).value
                if isinstance(v, (int, float)) and 1990 <= v <= 2100:
                    year_cols[c] = int(v)
            if not year_cols:
                self.log.error(f"Could not find any year columns in {path.name}")
                return {}

            latest_year = max(year_cols.values())
            latest_start_col = next(c for c, y in year_cols.items() if y == latest_year)
            sorted_cols = sorted(year_cols.keys())
            idx = sorted_cols.index(latest_start_col)
            group_end_col = sorted_cols[idx + 1] - 1 if idx + 1 < len(sorted_cols) else ws.max_column

            sub_row = label_row + 1
            nrw_col = None
            for c in range(latest_start_col, group_end_col + 1):
                v = ws.cell(sub_row, c).value
                if isinstance(v, str) and "non-resident workers" in v.lower():
                    nrw_col = c
                    break
            if nrw_col is None:
                self.log.error(
                    f"Could not find a 'Non-resident workers on-shift' sub-column "
                    f"for {latest_year} in {path.name}"
                )
                return {}

            result = {}
            current_lga = None
            for r in range(sub_row + 1, ws.max_row + 1):
                lga_val = ws.cell(r, lga_col).value
                if lga_val:
                    current_lga = str(lga_val).strip()
                ucl_val = ws.cell(r, ucl_col).value
                value = ws.cell(r, nrw_col).value
                if ucl_val and isinstance(value, (int, float)):
                    result[str(ucl_val).strip()] = {
                        "lga": current_lga,
                        "year": latest_year,
                        "value": int(value),
                    }

            self.log.info(f"  Parsed {len(result)} UCL entries for {latest_year} from {path.name}")
            return result

        except Exception as exc:
            self.log.error(f"FTE sheet parse error ({path.name}): {exc}", exc_info=True)
            return {}

    # ── Per-town extraction ──────────────────────────────────────────────────

    def _extract_town(self, town, region_key: str, region_result: dict):
        nrw_lga = region_result.get("nrw_lga", {})
        fte_lga_ucl = region_result.get("fte_lga_ucl", {})

        lga_series = None
        for key in nrw_lga:
            if key.lower() == town.lga.lower():
                lga_series = nrw_lga[key]
                break

        ucl_match = None
        for candidate in (town.name, town.sa2_name):
            if not candidate:
                continue
            for key in fte_lga_ucl:
                if key.lower() == candidate.lower():
                    ucl_match = fte_lga_ucl[key]
                    break
            if ucl_match:
                break

        if lga_series is None and ucl_match is None:
            self.log.warning(f"  [{town.name}] no NRW/FTE match in {region_key} data -- skipping")
            self.result.towns_skipped.append(town.name)
            return

        out = {
            "town":                      town.name,
            "state":                     town.state,
            "lga":                       town.lga,
            "source_region":             region_key,
            "lga_nrw_on_shift_by_year":  lga_series,
            "ucl_nrw_latest":            ucl_match,
        }

        out_dir = Path(__file__).parent.parent / "cache" / "population"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{town.slug}_population_nrw.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        lga_note = f"LGA series: {len(lga_series)} years" if lga_series else "LGA series: none"
        ucl_note = f"UCL {ucl_match['year']} = {ucl_match['value']:,}" if ucl_match else "UCL: none"
        self.log.info(f"  {town.name}: {lga_note}, {ucl_note}")
        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = QGSOPopulationNRWFetcher().run()
    sys.exit(0 if result.success else 1)