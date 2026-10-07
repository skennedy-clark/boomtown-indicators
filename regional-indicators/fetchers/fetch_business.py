"""
regional-indicators/fetchers/fetch_business.py

Fetches ABS business counts by SA2, industry division and turnover size
range from "Counts of Australian Businesses, including Entries and
Exits", Data cube 9.

Source:
  https://www.abs.gov.au/statistics/economy/business-indicators/counts-australian-businesses-including-entries-and-exits/{release}/8165DC09.xlsx
  Release in ABS_RELEASE: jul2021-jun2025 (published 16 Dec 2025).

File layout:
  The workbook has one table per year, not one combined time series:
    Table 1 = June 2025 (the latest year)
    Table 2 = June 2024
    Table 3 = June 2023
  Only Table 1 is read. The indicators workbook already holds earlier
  years from past updates, so each run adds only the year that Table 1
  represents.

  Columns (headers in rows 6-7, data from row 8):
    A: Industry Code    B: Industry Label
    C: SA2 Code          D: SA2 Label
    E: 0-<$50k no.        F: $50k-<$200k no.
    G: $200k-<$2m no.     H: $2m-<$5m no.
    I: $5m-<$10m no.      J: $10m+ no.
    K: Total no.

  Every industry code appears twice: once as per-SA2 detail rows, and
  once as a "Total <industry>" row with SA2 Code and SA2 Label both
  blank (a state/national aggregate). The aggregate rows are excluded
  by requiring a non-blank SA2 code; including them would double-count
  the totals.

Method (as documented in Notes_DD.docx and checked against the data):
  - Industry code "A" (Agriculture, Forestry and Fishing) is Primary
    Production. All other codes, including "X" (Currently Unknown), are
    Non-Primary Production and are summed together. The notes say to
    "group all OTHER industries as non-primary production", so X is
    deliberately included in NPP.
  - Turnover bands: the first three are kept as published (0-50k,
    50k-200k, 200k-2m); 2m-5m, 5m-10m and 10m+ are combined into a
    single "2m+" category. This matches the workbook's four bands.
  - SA2 rows are summed within each production category for a single
    SA2. A composite area (Toowoomba) sums across all its component
    SA2s for each turnover band.

SA2 mapping (checked against the Business sheet and towns.toml):
  Chinchilla=307011172  Goondiwindi=307011173  Moranbah=312011341
  Tara=307011178  Wambo(Dalby)=307021183  Miles-Wandoan=307011175
  Roma=307011176  Roma Surrounds(Wallumbilla)=307011177
  Broadsound-Nebo(Dysart)=312011338  Narrabri=110031197
  Narrabri Surrounds=110031198 (no corresponding project town)
  Toowoomba - SA2 Composite = the ten urban Toowoomba SA2s (317011446,
    -47, -52, -53, -54, -55, -56, -57, -58, -59): Darling Heights,
    Drayton - Harristown, Middle Ridge, Newtown (Qld), North Toowoomba -
    Harlaxton, Rangeville, Toowoomba - Central, - East, - West,
    Wilsonton.

    Basis for the composite: the previous cycle's raw download ("ABS SA2
    Turnover Businesses.xlsx") lists exactly these SA2s on its
    "Extracted" sheet. The ten reproduce the workbook's 2023/24
    composite exactly on all 8 values (NPP 2042/2855/3602/987, PP
    223/184/156/37), and no other subset of the ten does. Smaller
    definitions (three or four SA2s, or excluding Toowoomba - East) do
    not reproduce the published figures.

    Known discrepancy: the reference workbook's 2024/25 composite (NPP
    1865/2745/3463/964, PP 212/169/138/52) equals these ten minus
    Rangeville (again the only subset that matches exactly; Rangeville's
    126 accounts for the whole gap in the first band). The 2023/24
    column included Rangeville, so the omission appears to be an
    oversight in the reference workbook's newest column, not a change of
    definition; this is unconfirmed. This fetcher uses all ten, so its
    2024/25 composite exceeds the reference workbook's by Rangeville's
    contribution.

    Validation: the other 11 regions reproduce the reference workbook's
    2024/25 column exactly (22 of 22 NPP/PP blocks).

Output: cache/business/<region slug>_business.json, one per region in
SA2_REGIONS.
"""

from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher

try:
    import openpyxl
    import requests
except ImportError:
    raise ImportError("pip install openpyxl requests")


ABS_RELEASE = "jul2021-jun2025"   # update each cycle if the auto-advance
                                  # fails; see _discover_current_release()
ABS_DC9_BASE = (
    "https://www.abs.gov.au/statistics/economy/business-indicators/"
    "counts-australian-businesses-including-entries-and-exits"
)
ABS_DC9_URL = f"{ABS_DC9_BASE}/{ABS_RELEASE}/8165DC09.xlsx"
PRIMARY_PRODUCTION_CODE = "A"


def _guess_next_release(release: str) -> str | None:
    """Return the release name one year on from `release`, for example
    "jul2021-jun2025" -> "jul2022-jun2026".

    The ABS release name for this publication is a rolling 5-year window
    that has advanced by one year with each release. Returns None if the
    string does not match the "julYYYY-junYYYY" shape, so that a change
    of naming scheme falls back to the hardcoded release.
    """
    m = re.match(r"^jul(\d{4})-jun(\d{4})$", release)
    if not m:
        return None
    start, end = int(m.group(1)), int(m.group(2))
    return f"jul{start + 1}-jun{end + 1}"


def _discover_current_release(log) -> str:
    """Return the release name to download.

    Allows the fetcher to run in the following cycle without a code
    change: the next release name is derived from ABS_RELEASE, checked
    with a HEAD request, and used only if that request returns 200.
    Otherwise the hardcoded ABS_RELEASE is used. An unpublished release
    returns 404, which takes the fallback path.

    ABS_RELEASE still needs a manual update if the derivation stops
    working, for example if ABS changes the naming scheme, retires the
    5-year rolling window, or skips a year.
    """
    guess = _guess_next_release(ABS_RELEASE)
    if not guess:
        log.warning(f"  Release name {ABS_RELEASE!r} doesn't match the expected "
                    f"'julYYYY-junYYYY' shape -- cannot guess next year's, using hardcoded value")
        return ABS_RELEASE
    guess_url = f"{ABS_DC9_BASE}/{guess}/8165DC09.xlsx"
    try:
        resp = requests.head(guess_url, timeout=15, allow_redirects=True)
        if resp.status_code == 200:
            log.info(f"  Auto-advanced ABS release: {ABS_RELEASE} -> {guess} (verified live)")
            return guess
    except requests.RequestException as exc:
        log.warning(f"  Could not check guessed release {guess!r}: {exc}")
    log.info(f"  Guessed release {guess!r} not live yet -- using {ABS_RELEASE!r}")
    return ABS_RELEASE

# SA2 codes for every region the Business sheet tracks (see the module
# docstring for how they were established).
SA2_REGIONS = {
    "Broadsound-Nebo":            [312011338],
    "Chinchilla":                 [307011172],
    "Goondiwindi":                [307011173],
    "Miles-Wandoan":              [307011175],
    "Moranbah":                   [312011341],
    "Narrabri":                   [110031197],
    "Narrabri Surrounds":         [110031198],
    "Roma":                       [307011176],
    "Roma Surrounds":             [307011177],
    "Tara":                       [307011178],
    "Toowoomba - SA2 Composite":  [
        317011446,  # Darling Heights
        317011447,  # Drayton - Harristown
        317011452,  # Middle Ridge
        317011453,  # Newtown (Qld)
        317011454,  # North Toowoomba - Harlaxton
        317011455,  # Rangeville (see "Known discrepancy" in the module docstring)
        317011456,  # Toowoomba - Central
        317011457,  # Toowoomba - East
        317011458,  # Toowoomba - West
        317011459,  # Wilsonton
    ],
    "Wambo":                      [307021183],
}


class ABSBusinessFetcher(BaseFetcher):

    SOURCE_NAME      = "business"
    SUPPORTED_STATES = []

    def fetch_all(self):
        release = _discover_current_release(self.log)
        url = f"{ABS_DC9_BASE}/{release}/8165DC09.xlsx"
        self._source_url = url   # recorded by _write_region as the source URL
        cache_key = f"abs_dc9_{release}"
        path = self.download(url, cache_key, suffix=".xlsx")
        if not path:
            self.result.add_error("ALL", "Data cube 9 download failed")
            return

        self.log.info(f"  Parsing {path.name} ({path.stat().st_size // 1024} KB)")

        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            ws = wb["Table 1"]  # the latest year; see module docstring
        except Exception as exc:
            self.log.error(f"Could not open Table 1: {exc}", exc_info=True)
            self.result.add_error("ALL", f"Could not open Table 1: {exc}")
            return

        # The year label is in row 4, e.g.
        # "Businesses by Industry Division by SA2 by Turnover Size Ranges, June 2025 (a) (b)"
        title_cell = ws.cell(4, 1).value or ""
        self.log.info(f"  Table 1 title: {title_cell}")

        by_sa2 = self._parse_table(ws)
        if not by_sa2:
            self.result.add_error("ALL", "No data rows parsed from Table 1")
            return

        for region_name, sa2_codes in SA2_REGIONS.items():
            self._write_region(region_name, sa2_codes, by_sa2)

    def _parse_table(self, ws) -> dict:
        """Return {sa2_code: {"NPP": {band: count}, "PP": {band: count}}}.

        Aggregation happens here, directly from the raw rows: turnover
        bands are combined into the workbook's four categories and
        industry rows are split into NPP/PP, so that writing a region is
        a plain sum.
        """
        by_sa2: dict = defaultdict(lambda: {
            "NPP": defaultdict(int), "PP": defaultdict(int),
        })
        rows_used, rows_skipped_total, rows_skipped_other = 0, 0, 0

        for row in ws.iter_rows(min_row=8, values_only=True):
            industry_code = row[0]
            if industry_code is None:
                break  # end of data block

            sa2_code = row[2]
            if sa2_code is None:
                rows_skipped_total += 1
                continue  # "Total <industry>" aggregate row, not per-SA2

            try:
                sa2_code = int(sa2_code)
            except (TypeError, ValueError):
                rows_skipped_other += 1
                continue

            band = "PP" if industry_code == PRIMARY_PRODUCTION_CODE else "NPP"
            b_0_50k, b_50_200k, b_200k_2m, b_2_5m, b_5_10m, b_10m_plus = (
                row[4] or 0, row[5] or 0, row[6] or 0, row[7] or 0, row[8] or 0, row[9] or 0
            )

            entry = by_sa2[sa2_code][band]
            entry["0k-50k"]   += b_0_50k
            entry["50k-200k"] += b_50_200k
            entry["200k-2m"]  += b_200k_2m
            entry["2m+"]      += (b_2_5m + b_5_10m + b_10m_plus)
            rows_used += 1

        self.log.info(
            f"  Parsed {rows_used} rows ({len(by_sa2)} SA2s), "
            f"skipped {rows_skipped_total} state/national totals, "
            f"{rows_skipped_other} unparseable"
        )
        return dict(by_sa2)

    def _write_region(self, region_name: str, sa2_codes: list, by_sa2: dict):
        combined = {"NPP": defaultdict(int), "PP": defaultdict(int)}
        found_any = False

        for sa2_code in sa2_codes:
            if sa2_code not in by_sa2:
                self.log.warning(f"  [{region_name}] SA2 {sa2_code} not found in Table 1")
                continue
            found_any = True
            for production in ("NPP", "PP"):
                for band, count in by_sa2[sa2_code][production].items():
                    combined[production][band] += count

        if not found_any:
            self.log.warning(f"  [{region_name}] no SA2 data found at all -- skipping")
            self.result.towns_failed.append(region_name)
            return

        out = {
            "region": region_name,
            "sa2_codes": sa2_codes,
            "source": "ABS Counts of Australian Businesses, Data cube 9",
            "source_url": self._source_url,
            "note": (
                f"Industry code 'A' (Agriculture, Forestry and Fishing) = Primary "
                f"Production; all other industry codes (including 'X', Currently "
                f"Unknown) = Non-Primary Production, summed. Turnover bands "
                f"2m-5m/5m-10m/10m+ combined into a single '2m+' category. "
                f"{'Composite of ' + str(len(sa2_codes)) + ' SA2s.' if len(sa2_codes) > 1 else ''}"
            ),
            "indicators": {
                "npp": {band: count for band, count in combined["NPP"].items()},
                "pp":  {band: count for band, count in combined["PP"].items()},
            },
        }

        out_dir = Path(__file__).parent.parent / "cache" / "business"
        out_dir.mkdir(parents=True, exist_ok=True)
        slug = region_name.lower().replace(" ", "_").replace("-", "_")
        out_path = out_dir / f"{slug}_business.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        self.log.info(
            f"  {region_name}: NPP total={sum(combined['NPP'].values())}, "
            f"PP total={sum(combined['PP'].values())} -> {out_path.name}"
        )
        self.result.towns_ok.append(region_name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = ABSBusinessFetcher().run()
    sys.exit(0 if result.success else 1)