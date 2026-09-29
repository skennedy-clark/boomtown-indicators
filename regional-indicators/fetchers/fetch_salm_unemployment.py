"""
fetchers/fetch_salm_unemployment.py
-------------------------------------
Fetches Small Area Labour Markets (SALM) unemployment rates from the
Department of Employment and Workplace Relations (DEWR).

Sources (both are quarterly, published ~3 months after the reference quarter):
  SA2: "SALM Smoothed SA2 Datafiles (ASGS 2021)"  -- cache/salm_smoothed_sa2.csv
  LGA: "SALM Smoothed LGA Datafiles (ASGS 2025)"  -- cache/salm_smoothed_lga.csv
  Page: https://www.dewr.gov.au/employment-research/small-area-labour-markets

Both CSVs share a layout: a 3-row block per region (unemployment level,
labour force, unemployment RATE), quarterly columns "Dec-10", "Mar-11" ...

ANNUAL VALUE = plain arithmetic mean of the 4 quarterly smoothed rates in the
calendar year (Mar, Jun, Sep, Dec), UNROUNDED. VERIFIED 2026-09-28 against the
real workbook, not assumed: e.g. Chinchilla 2012's quarters are 1.8/1.7/1.7/1.5
= 1.675 and the workbook holds exactly 1.675; Tara 2025 = 11.2 matches to the
digit. Tested alternatives -- December-quarter only matched 12% of history,
mean-of-4 matched 68-69%. The residual mismatches are VINTAGE, not method:
worst in 2019-2023 (~0.15pp average) where SALM re-estimated its history at the
2016->2021 ASGS changeover, back to 13-14 of 15 matching in 2024-25. (The
"double-smoothing" worry -- averaging already-4-quarter-averaged values -- is
real in principle but is simply the workbook's own long-standing convention.)

COMPLETE YEARS ONLY (added 2026-09-28): a year needs ALL FOUR quarters. Before
this the fetcher averaged however many quarters existed, so a partial current
year (e.g. 2026 with only Mar-26) or the first year of the series (Dec-10
alone) would silently become a one-quarter "annual" figure -- same class of bug
as the partial-year problem in Crime.

TWO OUTPUTS:
  1. Per-TOWN files  cache/unemployment/{town.slug}_salm.json  (SA2 level).
     FORMAT AND NAMES UNCHANGED -- transform/to_csv.py and
     transform/booklet/pages/data_page.py read these directly
     (indicators.unemployment = flat {year: value}). Do not change them
     without changing those consumers.
  2. Per-REGION files cache/unemployment/regions/{sa2|lga}_{slug}.json, one
     for every row of the workbook's Employment sheet that SALM can supply
     (15 SA2 rows incl. Toowoomba - East and Narrabri Surrounds, which are
     not "towns", and 5 LGA rows). Read by update_employment.py.

REGIONS ARE MATCHED BY CODE, THEN NAME-CHECKED. If SALM's own name for a code
does not match the workbook row's label, that region is REFUSED rather than
written -- a drifted/reassigned code must fail loudly, not put the wrong
region's numbers in a row (cf. the Wallumbilla QPS division lesson).

NOT COVERED HERE: the NSW and "Queensland (benchmark)" state rows. SALM does not
publish states. Their values (NSW: long decimals like 5.2707 = a monthly
average; Queensland from 2010: multiples of 0.025 = a mean of four 1-decimal
quarters) look like ABS Labour Force Survey figures -- source still to be
confirmed with Steve.

CACHE-ONLY LGA FILE: the DEWR site returned 503 to automated requests during
development, so the LGA CSV is expected to be placed manually at
cache/salm_smoothed_lga.csv (download from the SALM page, "Smoothed LGA
Datafiles"). The fetcher logs which quarter each file ends at so a stale cache
is visible. The SA2 file keeps its original scrape-then-fallback download.

SA2 matching for towns uses towns.toml sa2_code (ASGS 2021 Edition 3).
Known towns.toml problems found 2026-09-28: Shepparton's sa2_code is not a
valid ASGS 2021 code (the town spans "Shepparton - North" 216031416 and
"Shepparton - South East" 216031594 -- needs a decision), and Yarram's was
wrong (correct: 205051104). These are why the VIC towns "fail" below.
"""

from __future__ import annotations

import csv
import io
import json
import re
import statistics
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR

try:
    import requests
except ImportError:
    raise ImportError("pip install requests")


# ── Configuration ──────────────────────────────────────────────────────────────

# Stable resource page — scrape this to find current CSV URL
# Main SALM page (accessible) and resource page (may be blocked)
SALM_MAIN_PAGE = "https://www.dewr.gov.au/employment-research/small-area-labour-markets"
RESOURCE_PAGE  = (
    "https://www.dewr.gov.au/employment-research/resources/"
    "salm-smoothed-sa2-datafiles-asgs-2021"
)

# Known-good URL from December quarter 2025 release (try first, fallback if 404)
FALLBACK_CSV_URL = (
    "https://www.dewr.gov.au/download/17068/"
    "salm-smoothed-sa2-datafiles-asgs-2021-december-quarter-2025/"
    "42403/salm-smoothed-sa2-datafiles-asgs-2021-december-quarter-2025/csv"
)

CACHE_KEY = "salm_smoothed_sa2"
LGA_CACHE_KEY = "salm_smoothed_lga"
LGA_PAGE = SALM_MAIN_PAGE

# Workbook Employment-sheet row label (column A) -> ASGS 2021 SA2 code.
# Same 15 regions verified for the Business sheet (2026-09-23).
SA2_REGIONS = {
    "Broadsound-Nebo":             "312011338",
    "Chinchilla":                  "307011172",
    "Goondiwindi":                 "307011173",
    "Miles-Wandoan":               "307011175",
    "Moranbah":                    "312011341",
    "Narrabri":                    "110031197",
    "Narrabri Surrounds":          "110031198",
    "North Toowoomba - Harlaxton": "317011454",
    "Roma":                        "307011176",
    "Roma Surrounds":              "307011177",
    "Tara":                        "307011178",
    "Toowoomba - Central":         "317011456",
    "Toowoomba - East":            "317011457",
    "Toowoomba - West":            "317011458",
    "Wambo":                       "307021183",
}

# Workbook label -> ASGS 2025 LGA code (verified in the March-2026 LGA file).
LGA_REGIONS = {
    "Goondiwindi":   "33610",
    "Isaac":         "33980",
    "Maranoa":       "34860",
    "Toowoomba LGA": "36910",
    "Western Downs": "37310",
}

QUARTERS = ("Mar", "Jun", "Sep", "Dec")
INDICATOR_LABEL = "Smoothed Unemployment rate (%)"

# CSV column format: "Data Item", "SA2 name", "SA2 Code (2021 ASGS)", "Mar-10", "Jun-10", ...
# Rows alternate between unemployment level and unemployment rate
RATE_ROW_LABEL = "Smoothed unemployment rate"


class SALMUnemploymentFetcher(BaseFetcher):

    SOURCE_NAME      = "salm_unemployment"
    SUPPORTED_STATES = []   # national

    def fetch_all(self):
        self._file_quarters = {}

        # ── SA2 file (scrape-then-fallback, unchanged) ────────────────────────
        url  = self._find_csv_url()
        path = self._download_with_browser_ua(url, CACHE_KEY)

        if not path:
            self.result.add_error("ALL", "Could not download SALM SA2 CSV")
            return

        self.log.info(f"  Parsing {path.name} ({path.stat().st_size // 1024} KB)")

        sa2_data = self._parse_csv(path, kind="SA2")
        if not sa2_data:
            self.result.add_error("ALL", "SALM CSV parse returned no data")
            return

        # Per-town files: format/names unchanged (consumed by to_csv.py and
        # the booklet's data_page.py).
        for town in self.applicable_towns():
            self._extract_town(town, sa2_data)

        # Per-region files for the Employment sheet: SA2 rows ...
        for label, code in SA2_REGIONS.items():
            self._write_region("SA2", label, code, sa2_data, path.name)

        # ... and LGA rows (cache-only file, see module docstring).
        lga_path = CACHE_DIR / f"{LGA_CACHE_KEY}.csv"
        if not lga_path.exists():
            msg = (
                f"LGA file not found at {lga_path}. Download 'SALM Smoothed LGA "
                f"Datafiles' (CSV) from {LGA_PAGE} and save it there -- the DEWR "
                f"site blocks automated downloads, so this one is manual."
            )
            self.log.error("  " + msg)
            self.result.add_error("LGA", msg)
            return
        self.log.info(f"  Parsing {lga_path.name} ({lga_path.stat().st_size // 1024} KB)")
        lga_data = self._parse_csv(lga_path, kind="LGA")
        if not lga_data:
            self.result.add_error("LGA", "SALM LGA CSV parse returned no data")
            return
        for label, code in LGA_REGIONS.items():
            self._write_region("LGA", label, code, lga_data, lga_path.name)

        self.log.info(
            "  File vintages: "
            + ", ".join(f"{k} ends {v}" for k, v in self._file_quarters.items())
        )

    def _download_with_browser_ua(self, url: str, cache_key: str) -> Path | None:
        """Download with browser User-Agent to bypass bot detection."""
        from config import CACHE_DIR
        out_path = CACHE_DIR / f"{cache_key}.csv"
        if out_path.exists() and not self.force:
            self.log.info(f"  Using cached: {out_path.name}")
            return out_path
        try:
            headers = {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                )
            }
            self.log.info(f"  Downloading {url} → {out_path.name}")
            resp = requests.get(url, headers=headers, timeout=60, stream=True)
            resp.raise_for_status()
            content_type = resp.headers.get("Content-Type", "")
            if "html" in content_type.lower():
                self.log.error(f"  Got HTML response (not CSV) — URL may be wrong")
                return None
            with open(out_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=65536):
                    f.write(chunk)
            size_kb = out_path.stat().st_size // 1024
            self.log.info(f"  Saved {size_kb} KB")
            if size_kb < 500:
                self.log.warning(f"  File is small ({size_kb} KB) — expected ~2000+ KB. May be an error page.")
            return out_path
        except Exception as exc:
            self.log.error(f"  Download failed: {exc}")
            return None

    def _find_csv_url(self) -> str:
        """
        Scrape resource page to find current CSV download URL.
        Falls back to the known-good URL if the page can't be reached.
        """
        try:
            headers = {"User-Agent": "Mozilla/5.0 (research pipeline; contact uq.edu.au)"}
            resp = requests.get(RESOURCE_PAGE, headers=headers, timeout=20)
            resp.raise_for_status()

            # Find CSV download link — pattern: /download/{id}/salm-smoothed-sa2...
            pattern = r'(https://www\.dewr\.gov\.au/download/\d+/salm-smoothed-sa2[^"\'>\s]+/csv)'
            matches = re.findall(pattern, resp.text)
            if matches:
                url = matches[0]
                self.log.info(f"  Found CSV URL on resource page: {url}")
                return url

            self.log.warning("  Could not find CSV link on resource page — using fallback URL")
        except Exception as exc:
            self.log.warning(f"  Resource page scrape failed ({exc}) — using fallback URL")

        return FALLBACK_CSV_URL

    def _parse_csv(self, path: Path, kind: str = "SA2") -> dict:
        """
        Parse a SALM smoothed-rate CSV (SA2 or LGA layout) into:
          { code: (salm_name, { year: annual_rate }) }

        Annual rate = plain mean of the 4 quarterly smoothed rates (Mar, Jun,
        Sep, Dec) in the calendar year -- VERIFIED against the workbook, see
        module docstring. A year is only produced if ALL FOUR quarters are
        present (no partial years). Records the file's last quarter in
        self._file_quarters[kind] so a stale cache is visible in the log.
        """
        try:
            with open(path, encoding="utf-8-sig", errors="replace") as f:
                lines = f.readlines()

            header_idx = None
            for i, line in enumerate(lines):
                if line.lstrip().startswith("Data Item"):
                    header_idx = i
                    break
            if header_idx is None:
                self.log.error(f"Could not find header row in SALM {kind} CSV")
                return {}

            reader = csv.reader(lines[header_idx:])
            headers = next(reader)

            col_to_yq = {}
            last_label = None
            for i, h in enumerate(headers[3:], start=3):
                m = re.match(r"^(Mar|Jun|Sep|Dec)-(\d{2})$", h.strip())
                if m:
                    yr2 = int(m.group(2))
                    year = 2000 + yr2 if yr2 <= 50 else 1900 + yr2
                    col_to_yq[i] = (year, m.group(1))
                    last_label = h.strip()
            if hasattr(self, "_file_quarters") and last_label:
                self._file_quarters[kind] = last_label

            result: dict[str, tuple] = {}
            for row in reader:
                if len(row) < 3:
                    continue
                if RATE_ROW_LABEL not in row[0].strip():
                    continue
                code = str(row[2]).strip().replace(".0", "")
                if not code.isdigit():
                    continue

                by_year: dict[int, dict[str, float]] = {}
                for col_i, (year, q) in col_to_yq.items():
                    if col_i >= len(row):
                        continue
                    val = row[col_i].strip().replace(",", "")
                    if val and val != "-":
                        try:
                            by_year.setdefault(year, {})[q] = float(val)
                        except ValueError:
                            pass

                annual = {
                    yr: round(statistics.mean(qs[q] for q in QUARTERS), 4)
                    for yr, qs in by_year.items()
                    if all(q in qs for q in QUARTERS)
                }
                result[code] = (row[1].strip(), annual)

            years = [y for _, d in result.values() for y in d]
            if years:
                self.log.info(
                    f"  Parsed {len(result)} {kind}s, complete years "
                    f"{min(years)} to {max(years)}"
                )
            return result

        except Exception as exc:
            self.log.error(f"SALM {kind} CSV parse error: {exc}", exc_info=True)
            return {}

    @staticmethod
    def _norm_name(name: str) -> str:
        """Compare spellings loosely: 'Broadsound - Nebo' == 'Broadsound-Nebo',
        'Toowoomba' == 'Toowoomba LGA'."""
        n = name.lower()
        n = re.sub(r"\blga\b", "", n)
        return re.sub(r"[^a-z0-9]", "", n)

    def _write_region(self, section: str, label: str, code: str, data: dict, source_file: str):
        tag = f"{section}:{label}"
        rec = data.get(code)
        if not rec:
            self.log.warning(f"  [{tag}] code {code} not found in SALM {section} data")
            self.result.towns_failed.append(tag)
            return
        salm_name, annual = rec

        # Refuse a region whose code now belongs to something else.
        if self._norm_name(salm_name) != self._norm_name(label):
            self.log.error(
                f"  [{tag}] NAME MISMATCH: code {code} is '{salm_name}' in SALM but the "
                f"workbook row is '{label}' -- refusing to write (code may have been "
                f"reassigned between ASGS editions)"
            )
            self.result.towns_failed.append(tag)
            return
        if not annual:
            self.log.warning(f"  [{tag}] no complete years in SALM data")
            self.result.towns_failed.append(tag)
            return

        from config import YEAR_START, YEAR_END
        values = {str(y): v for y, v in sorted(annual.items()) if YEAR_START <= y <= YEAR_END}

        out = {
            "region":     label,
            "section":    section,
            "code":       code,
            "salm_name":  salm_name,
            "source":     f"DEWR Small Area Labour Markets (SALM) — smoothed {section}",
            "source_file": source_file,
            "file_ends":  self._file_quarters.get(section),
            "note": (
                "Annual value = plain mean of the 4 quarterly smoothed rates in the "
                "calendar year (complete years only). Verified against the workbook."
            ),
            "indicators": {"unemployment": {"label": INDICATOR_LABEL, "values": values}},
        }

        slug = re.sub(r"[^a-z0-9]+", "_", label.lower()).strip("_")
        out_dir = CACHE_DIR / "unemployment" / "regions"
        out_dir.mkdir(parents=True, exist_ok=True)
        with open(out_dir / f"{section.lower()}_{slug}.json", "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        latest = max(annual)
        self.log.info(f"  [{tag}] {len(values)} years, latest ({latest}) = {annual[latest]:.3f}%")
        self.result.towns_ok.append(tag)

    def _extract_town(self, town, data: dict):
        """Match town SA2 code and write cache JSON."""
        sa2 = town.sa2_code
        if not sa2:
            self.log.warning(f"  [{town.name}] no sa2_code in towns.toml")
            self.result.towns_skipped.append(town.name)
            return

        rec = data.get(sa2)
        annual = rec[1] if rec else None
        if not annual:
            self.log.warning(f"  [{town.name}] SA2 {sa2} not found in SALM data")
            self.result.towns_failed.append(town.name)
            return

        from config import YEAR_START, YEAR_END
        values = {
            str(yr): val for yr, val in annual.items()
            if YEAR_START <= yr <= YEAR_END
        }

        latest_yr = max(annual)
        out = {
            "town":       town.name,
            "state":      town.state,
            "sa2_code":   sa2,
            "source":     "DEWR Small Area Labour Markets (SALM) — smoothed SA2",
            "source_url": RESOURCE_PAGE,
            "note":       "Annual value = mean of 4 quarterly smoothed unemployment rates",
            "indicators": {"unemployment": values},
        }

        out_dir = CACHE_DIR / "unemployment"
        out_dir.mkdir(exist_ok=True)
        out_path = out_dir / f"{town.slug}_salm.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        self.log.info(
            f"  {town.name} (SA2 {sa2}): {len(values)} years, "
            f"latest ({latest_yr}) = {annual[latest_yr]:.2f}%"
        )
        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = SALMUnemploymentFetcher().run()
    sys.exit(0 if result.success else 1)