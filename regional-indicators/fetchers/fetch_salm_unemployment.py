"""
regional-indicators/fetchers/fetch_salm_unemployment.py

Fetches Small Area Labour Markets (SALM) smoothed unemployment rates from
the Department of Employment and Workplace Relations (DEWR) and writes
annual rates per town and per Employment-sheet region.

Sources (both quarterly, published about three months after the
reference quarter):
  SA2: "SALM Smoothed SA2 Datafiles (ASGS 2021)" - cache/salm_smoothed_sa2.csv
  LGA: "SALM Smoothed LGA Datafiles (ASGS 2025)" - cache/salm_smoothed_lga.csv
  Page: https://www.dewr.gov.au/employment-research/small-area-labour-markets

File layout: both CSVs have a three-row block per region (unemployment
level, labour force, unemployment rate) and quarterly columns "Dec-10",
"Mar-11", ...

Annual value: the arithmetic mean of the four quarterly smoothed rates
in the calendar year (Mar, Jun, Sep, Dec). This reproduces the reference
workbook: Chinchilla 2012 has quarters 1.8/1.7/1.7/1.5 = 1.675 and the
workbook holds 1.675; Tara 2025 = 11.2 matches. The mean of four
quarters matches 68-69% of the workbook's history, against 12% for the
December quarter alone. The remaining differences are due to data
vintage, not method: they are largest in 2019-2023 (about 0.15
percentage points on average), where SALM re-estimated its history at
the ASGS 2016 to 2021 changeover, and 13-14 of 15 regions match in
2024-25. Averaging values that are already four-quarter smoothed is the
workbook's long-standing convention.

Complete years only: a year is produced only when all four quarters are
present. A partial current year, or the first year of the series (Dec-10
alone), would otherwise yield an annual figure from fewer quarters.

Output:
  1. Per-town files cache/unemployment/{town.slug}_salm.json (SA2
     level), with indicators.unemployment as a flat {year: value}
     mapping. transform/to_csv.py and
     transform/booklet/pages/data_page.py read these directly, so the
     format and names must not change without changing those consumers.
  2. Per-region files cache/unemployment/regions/{sa2|lga}_{slug}.json,
     one for every row of the workbook's Employment sheet that SALM can
     supply: 15 SA2 rows (including Toowoomba - East and Narrabri
     Surrounds, which are not towns) and 5 LGA rows. Read by
     update_employment.py.

Region matching: regions are matched by code and then checked by name.
If the SALM name for a code does not match the workbook row label, the
region is refused rather than written, so that a code reassigned between
ASGS editions fails visibly instead of placing another region's figures
in the row.

Not covered: the NSW and "Queensland (benchmark)" state rows. SALM does
not publish state figures. Those rows are supplied by
fetch_nsw_labour.py and fetch_qrsis_labour.py.

LGA file is cache-only: the DEWR site returns 503 to automated requests
for this file, so the LGA CSV must be placed manually at
cache/salm_smoothed_lga.csv (download "Smoothed LGA Datafiles" from the
SALM page). The fetcher logs the last quarter in each file so that a
stale cache is visible. The SA2 file is downloaded automatically, from a
scraped link with a fallback URL.

Notes: towns are matched on the sa2_code field in towns.toml (ASGS 2021
Edition 3). Shepparton's sa2_code is not a valid ASGS 2021 code; the
town spans "Shepparton - North" (216031416) and "Shepparton - South
East" (216031594) and has no single SA2, so it is reported as failed.
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

# RESOURCE_PAGE is scraped for the current CSV link; it may refuse
# automated requests. SALM_MAIN_PAGE is the main SALM page.
SALM_MAIN_PAGE = "https://www.dewr.gov.au/employment-research/small-area-labour-markets"
RESOURCE_PAGE  = (
    "https://www.dewr.gov.au/employment-research/resources/"
    "salm-smoothed-sa2-datafiles-asgs-2021"
)

# Update each cycle. Fallback only, used when scraping the resource page
# fails. This is the December quarter 2025 release.
FALLBACK_CSV_URL = (
    "https://www.dewr.gov.au/download/17068/"
    "salm-smoothed-sa2-datafiles-asgs-2021-december-quarter-2025/"
    "42403/salm-smoothed-sa2-datafiles-asgs-2021-december-quarter-2025/csv"
)

CACHE_KEY = "salm_smoothed_sa2"
LGA_CACHE_KEY = "salm_smoothed_lga"
LGA_PAGE = SALM_MAIN_PAGE

# Workbook Employment-sheet row label (column A) -> ASGS 2021 SA2 code.
# The same 15 regions are used for the Business sheet.
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

# Workbook label -> ASGS 2025 LGA code, as used in the SALM LGA file.
LGA_REGIONS = {
    "Goondiwindi":   "33610",
    "Isaac":         "33980",
    "Maranoa":       "34860",
    "Toowoomba LGA": "36910",
    "Western Downs": "37310",
}

QUARTERS = ("Mar", "Jun", "Sep", "Dec")
INDICATOR_LABEL = "Smoothed Unemployment rate (%)"

# CSV columns: "Data Item", "SA2 name", "SA2 Code (2021 ASGS)", "Mar-10",
# "Jun-10", ... Only rows whose "Data Item" contains this label are read.
RATE_ROW_LABEL = "Smoothed unemployment rate"


class SALMUnemploymentFetcher(BaseFetcher):

    SOURCE_NAME      = "salm_unemployment"
    SUPPORTED_STATES = []   # national

    def fetch_all(self):
        self._file_quarters = {}

        # ── SA2 file (scraped link, then fallback URL) ────────────────────────
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

        # Per-town files, read by to_csv.py and the booklet's data_page.py.
        for town in self.applicable_towns():
            self._extract_town(town, sa2_data)

        # Per-region files for the Employment sheet: SA2 rows ...
        for label, code in SA2_REGIONS.items():
            self._write_region("SA2", label, code, sa2_data, path.name)

        # ... and LGA rows (cache-only file; see the module docstring).
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
        """Download the CSV with a browser User-Agent; the DEWR site rejects
        requests it identifies as automated. Uses the cached file unless
        force is set.
        """
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
        """Return the current SA2 CSV download URL, scraped from the resource
        page. Falls back to FALLBACK_CSV_URL if the page cannot be read or
        has no matching link.
        """
        try:
            headers = {"User-Agent": "Mozilla/5.0 (research pipeline; contact uq.edu.au)"}
            resp = requests.get(RESOURCE_PAGE, headers=headers, timeout=20)
            resp.raise_for_status()

            # Download link pattern: /download/{id}/salm-smoothed-sa2.../csv
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
        """Parse a SALM smoothed-rate CSV (SA2 or LGA layout) into:
          { code: (salm_name, { year: annual_rate }) }

        The annual rate is the mean of the four quarterly smoothed rates
        (Mar, Jun, Sep, Dec) in the calendar year; see the module docstring.
        A year is produced only if all four quarters are present. The last
        quarter in the file is recorded in self._file_quarters[kind] so
        that a stale cache is visible in the log.
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
        """Normalise a region name for loose comparison, so that
        'Broadsound - Nebo' equals 'Broadsound-Nebo' and 'Toowoomba' equals
        'Toowoomba LGA'.
        """
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

        # Refuse a region whose code now belongs to a different region.
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
        """Write the cache JSON for a town, matched on its SA2 code."""
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