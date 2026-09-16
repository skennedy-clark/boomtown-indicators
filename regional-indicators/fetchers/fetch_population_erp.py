"""
regional-indicators/fetchers/fetch_population_erp.py
-------------------------------------------------------
Fetches the main SA2/LGA-level "Population (ERP)" figure -- the primary
indicator row in each town's block on the Population sheet, distinct
from fetch_population_ucl.py's UCL-level figure and
fetch_population_nrw.py's non-resident-worker figure.

CURRENT STATUS: no live-API path implemented yet. This fetcher only
implements the WORKING fallback -- reading a manually-downloaded QGSO
Regional Database export if the researcher has placed one at
cache/qgso_and_bom_{YEAR}.xlsx -- plus clear, actionable reporting when
that file isn't there. See "PATH TO FULL AUTOMATION" below before
building a live-API version from scratch.

WHY NOT A LIVE API CALL, LIKE fetch_qgso_housing.py HAS:
fetch_qgso_housing.py already implements real, working QRSIS API
automation (the same underlying system this data comes from) -- see its
BASE_URL/PUBLIC_USER/ACCESS_LEVEL/COLLGRP_ID constants and the
COLLECTIONS dict mapping each housing series to a specific QRSIS
collection id (e.g. 1925 = residential sales, 1929 = median rent).
Population (ERP) is almost certainly ALSO a QRSIS collection with its
own id -- it just hasn't been identified yet, and qgso_housing itself
is currently reported at "0 ok" (not confirmed working end-to-end), so
building a second, independent QRSIS integration for population risks
duplicating a problem that needs fixing once, not twice.

PATH TO FULL AUTOMATION (next real step, not done here):
  1. Fix/confirm fetch_qgso_housing.py's QRSIS calls actually work
     end-to-end (it's currently 0 ok per project status notes) --
     whatever's broken there likely affects any population collection
     built the same way.
  2. Discover the Population (ERP) collection id: go to
     https://statistics.qgso.qld.gov.au/ (or the Regional Database
     front door at http://www.qgso.qld.gov.au/products/tables/
     qld-regional-database/index.php), open browser devtools' Network
     tab, manually select "Population (ERP)" data for one town, and
     read the collection id out of the POST request the same way
     fetch_qgso_housing.py's docstring documents doing for housing
     (interleaved p_names/p_values, look for the id parameter).
  3. Add a "population_erp" entry to a COLLECTIONS-style dict mirroring
     fetch_qgso_housing.py's, reusing its _q() POST-encoding helper and
     HEADERS/BASE_URL constants rather than duplicating them.

WORKING FALLBACK (what this file actually does):
Per the project's own documented pipeline approach for QGSO Regional
Database sources (llm_context.md, "GROUP B"): attempt automation,
and where that's not possible, read a locally-placed manual export and
report clearly on what's missing. This fetcher:
  1. Looks for cache/qgso_and_bom_{YEAR}.xlsx, trying the current year
     and the year before (in case the researcher's file lags).
  2. If found, parses its "Pop" sheet.
  3. If not found, produces a clear, actionable error naming exactly
     where to get the file and what to name it.

"Pop" sheet format (confirmed from a real prior-year file, per project
notes -- structure may shift year to year, so this searches for the
header row rather than assuming a fixed position):
  Collection header: "Population (ERP)(a) persons only"
  Row 5 (approx) = header: Region | [town name] | [boundary type] | None | [year]
  SINGLE YEAR SNAPSHOT ONLY -- this file does not contain history, only
  whatever year QGSO's export currently covers. That's fine for this
  project's actual need (one more year at a time), not a limitation to
  work around.

*** NOT YET TESTED against a real cache/qgso_and_bom_{YEAR}.xlsx file --
the exact "Pop" sheet layout is taken from project notes describing a
2025 file, not independently re-verified. First real run against an
actual file is the real test. ***
"""

from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR

try:
    import openpyxl
except ImportError:
    raise ImportError("pip install openpyxl")


ASSEMBLED_FILE_PATTERN = "qgso_and_bom_{year}.xlsx"
REGIONAL_DATABASE_URL = (
    "http://www.qgso.qld.gov.au/products/tables/qld-regional-database/index.php"
)
POP_SHEET_NAME_CANDIDATES = ["Pop", "Population", "Pop sheet"]


class QGSOPopulationERPFetcher(BaseFetcher):

    SOURCE_NAME      = "population_erp"
    SUPPORTED_STATES = ["QLD"]  # NSW/VIC towns need a separate ABS-based fetcher --
                                 # see TODO.md, this is QLD/QRSIS-sourced data only

    def fetch_all(self):
        path = self._find_assembled_file()
        if not path:
            current_year = datetime.now().year
            self.result.add_error(
                "ALL",
                f"No manually-downloaded QGSO Regional Database export found. "
                f"This indicator has no working automated source yet (see this "
                f"file's docstring for the path to building one).\n"
                f"  Manual fix: visit {REGIONAL_DATABASE_URL}, download the "
                f"Population (ERP) data for this year's towns, assemble it "
                f"into a workbook with a 'Pop' sheet (same format as prior "
                f"years -- ask if unsure), and save it as one of:\n"
                f"    cache/{ASSEMBLED_FILE_PATTERN.format(year=current_year)}\n"
                f"    cache/{ASSEMBLED_FILE_PATTERN.format(year=current_year - 1)}\n"
                f"  Then re-run this fetcher."
            )
            return

        self.log.info(f"  Found manually-assembled file: {path.name}")
        pop_data = self._parse_pop_sheet(path)
        if not pop_data:
            self.result.add_error(
                "ALL",
                f"Found {path.name} but could not locate/parse its 'Pop' sheet. "
                f"Check the sheet is named one of {POP_SHEET_NAME_CANDIDATES} "
                f"and still has a 'Region' column and a year column."
            )
            return

        for town in self.applicable_towns():
            self._extract_town(town, pop_data)

    def _find_assembled_file(self) -> Path | None:
        """Look for this year's or last year's manually-assembled QGSO
        export. Year is never hardcoded -- computed from the current
        date, so this keeps working correctly in future years without
        a code change."""
        current_year = datetime.now().year
        for year in (current_year, current_year - 1):
            candidate = CACHE_DIR / ASSEMBLED_FILE_PATTERN.format(year=year)
            if candidate.exists():
                return candidate
        return None

    # ── Parser ─────────────────────────────────────────────────────────────

    def _parse_pop_sheet(self, path: Path) -> dict:
        """Return { "<region name>": {"value": int, "year": int} }.

        Searches for the header row (containing "Region") rather than
        assuming a fixed row number -- this file is hand-assembled by a
        researcher each year and its exact layout has already been
        observed to shift.
        """
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            sheet_name = next(
                (n for n in wb.sheetnames if n in POP_SHEET_NAME_CANDIDATES), None
            )
            if sheet_name is None:
                self.log.error(
                    f"No sheet named any of {POP_SHEET_NAME_CANDIDATES} found "
                    f"in {path.name} (sheets present: {wb.sheetnames})"
                )
                return {}

            ws = wb[sheet_name]

            header_row_idx = None
            region_col = None
            year_col = None
            year_val = None
            for r in range(1, min(15, ws.max_row) + 1):
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and v.strip().lower() == "region":
                        header_row_idx = r
                        region_col = c
                        break
                if header_row_idx:
                    break

            if header_row_idx is None:
                self.log.error(f"Could not find a 'Region' header cell in {sheet_name}")
                return {}

            # Year column: scan the header row for a 4-digit year value.
            for c in range(1, ws.max_column + 1):
                v = ws.cell(header_row_idx, c).value
                if isinstance(v, (int, float)) and 1990 <= v <= 2100:
                    year_col = c
                    year_val = int(v)
                    break
                if isinstance(v, str) and v.strip().isdigit() and 1990 <= int(v.strip()) <= 2100:
                    year_col = c
                    year_val = int(v.strip())
                    break

            if year_col is None:
                self.log.error(f"Could not find a year column in {sheet_name}'s header row")
                return {}

            result = {}
            for r in range(header_row_idx + 1, ws.max_row + 1):
                region = ws.cell(r, region_col).value
                value = ws.cell(r, year_col).value
                if region and isinstance(value, (int, float)):
                    result[str(region).strip()] = {"value": int(value), "year": year_val}

            self.log.info(f"  Parsed {len(result)} region entries for year {year_val}")
            return result

        except Exception as exc:
            self.log.error(f"Pop sheet parse error: {exc}", exc_info=True)
            return {}

    # ── Per-town extraction ────────────────────────────────────────────────

    def _extract_town(self, town, pop_data: dict):
        # SA2 name is the more specific match for most towns; LGA name
        # matters for towns whose Population (ERP) row is at LGA level
        # (e.g. Brisbane) rather than SA2 (most study towns) -- both are
        # tried since we don't yet have a per-town "which level" flag.
        candidates = [town.sa2_name, town.lga, town.name]
        match = None
        matched_as = None
        for candidate in candidates:
            if not candidate:
                continue
            if candidate in pop_data:
                match, matched_as = pop_data[candidate], candidate
                break
            for key in pop_data:
                if key.lower() == candidate.lower():
                    match, matched_as = pop_data[key], key
                    break
            if match:
                break

        if not match:
            self.log.warning(
                f"  [{town.name}] no Population (ERP) match "
                f"(tried: {[c for c in candidates if c]}) -- skipping"
            )
            self.result.towns_skipped.append(town.name)
            return

        out = {
            "town":        town.name,
            "state":       town.state,
            "source":      "QGSO Regional Database (manually assembled export)",
            "matched_as":  matched_as,
            "year":        match["year"],
            "value":       match["value"],
        }

        out_dir = Path(__file__).parent.parent / "cache" / "population"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{town.slug}_population_erp.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        self.log.info(f"  {town.name} ('{matched_as}'): {match['year']} = {match['value']:,}")
        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = QGSOPopulationERPFetcher().run()
    sys.exit(0 if result.success else 1)