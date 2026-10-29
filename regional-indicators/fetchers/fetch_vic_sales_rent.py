"""
regional-indicators/fetchers/fetch_vic_sales_rent.py

Fetches the median house sale price and the median weekly rent of
3-bedroom houses for each Victorian town, for the town's "<LGA> (LGA)"
block of the Housing sheet (the block copied from Narrabri by add_town).
The output has the same file name and schema as fetch_nsw_sales_rent.py,
so update_housing.py writes it without change.

Sources
-------
Sales: Valuer-General Victoria (VGV), Victorian Property Sales Report,
"Houses by suburb" time series: the median house price for each
locality (suburb or town) and calendar year, ten years per file (for
example houses-by-suburb-2015-2025.xlsx). Older editions reach earlier
years. The files are listed in the Victorian Government data catalogue:
    https://discover.data.vic.gov.au/api/3/action/package_show?id=victorian-property-sales-report-median-house-by-suburb-time-series
and served from www.land.vic.gov.au.

Rent: Department of Families, Fairness and Housing (DFFH), Rental
Report, "Quarterly median rents by Local Government Area": the median
weekly rent of new lettings (bonds lodged), by LGA and quarter, from
June 1999 to the file's quarter. Each quarter's file holds the full
history, so only the latest is needed. Listed in the catalogue:
    https://discover.data.vic.gov.au/api/3/action/package_show?id=rental-report-quarterly-quarterly-median-rents-by-lga
The catalogue links DFFH pages, which link the workbook.

File layouts
------------
Houses by suburb (one sheet): a header row with "Locality" in column A
and the years in every second column from B; each year's value column
is followed by a flag column ("^", "*" or blank; VGV's notes in "A
Guide to Property Values" explain them). After the last year come a
"Prelim" column (the following year to date, not used), change and
growth columns. Values are text ("495000"); "NA" and "-" mean no
figure. Localities are in capitals ("SHEPPARTON").

Quarterly median rents by LGA: one sheet per dwelling type ("3br House"
is read). Row 2 holds the quarter ("Sep 2025") over each pair of
columns, row 3 "Count" and "Median"; column A holds the region (only on
its first row) and column B the LGA. "-" marks a quarter DFFH does not
report (too few lettings).

Geography and gaps (the Housing block is headed by the LGA)
-----------------------------------------------------------
  Median Sale Price   VGV publishes medians by locality, not by LGA,
                      and for houses only. The town's locality is used
                      (towns.toml field `locality`, for example
                      "Shepparton"); a town without `locality` gets no
                      sale price. This is a different geography and
                      dwelling type from the New South Wales rows (LGA,
                      all dwellings); the cache file's note says so.
  Mean Sale Price,    not published by VGV at suburb level: left blank.
  Sales No.
  House 3-Bed         DFFH 3-bedroom houses, the town's LGA.
  Total (3-bed, all   DFFH publishes 3-bedroom houses and 3-bedroom
  dwelling types)     flats separately; their medians cannot be
                      combined, so the row is left blank.

Aggregation: rent is the mean of the four quarterly medians of a
calendar year, complete years only, as for New South Wales. Sale prices
are VGV's calendar-year medians as published; a year is taken from the
newest file that holds it (VGV revises recent years).

Manual downloads: www.land.vic.gov.au and www.dffh.vic.gov.au refuse
some automated requests. If a file cannot be downloaded, the copy saved
by hand in the data folder ("1 Data/Housing/vgv_houses_by_suburb.xlsx"
and "1 Data/Housing/dffh_rents_by_lga.xlsx"; see transform/
data_folder.py) is used; the run_end_to_end step "Manual downloads"
copies them into the cache.

Output: cache/housing/regions/<town slug>_sales_rent.json, read by
update_housing.py.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR, YEAR_START, YEAR_END

CATALOGUE = "https://discover.data.vic.gov.au/api/3/action/package_show?id={}"
SALES_DATASET = "victorian-property-sales-report-median-house-by-suburb-time-series"
RENT_DATASET = "rental-report-quarterly-quarterly-median-rents-by-lga"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

# Copies saved by hand (see transform/data_folder.py, MANUAL_FILES).
MANUAL_SALES = "vgv_houses_by_suburb.xlsx"
MANUAL_RENT = "dffh_rents_by_lga.xlsx"

SALES_FILE = re.compile(r"houses-by-suburb-(\d{4})-(\d{4})\.xlsx", re.IGNORECASE)
RENT_SHEET = "3br House"
MONTHS = {"mar": 1, "jun": 2, "june": 2, "sep": 3, "sept": 3, "dec": 4}
QUARTER_LABEL = re.compile(r"^\s*([A-Za-z]+)\.?\s+(\d{4})\s*$")
XLSX_LINK = re.compile(r'href="([^"]+?\.xlsx)"', re.IGNORECASE)

# Cache key -> label, in the order of fetch_nsw_sales_rent.py.
LABELS = {
    "mean_sales_price":   "Mean Sales Price $",
    "median_sales_price": "Median Sales Price $ (houses, town locality; VGV)",
    "sales_no":           "Sales No.",
    "rent_house_3bed":    "Median Weekly Rent for New Lettings $ (House, 3 Bedrooms; DFFH)",
    "rent_total_3bed":    "Median Weekly Rent for New Lettings $ (Total, 3 Bedrooms)",
}
NOT_PUBLISHED = {
    "mean_sales_price": "VGV publishes medians only",
    "sales_no": "VGV's suburb time series has no sales counts",
    "rent_total_3bed": "DFFH publishes 3-bedroom houses and flats separately; medians cannot be combined",
}


def lga_key(name) -> str:
    """An LGA name as compared between files ("Wellington (Vic.)" and
    "wellington" both give "wellington")."""
    return re.sub(r"\s*\([^)]*\)\s*$", "", str(name or "")).strip().casefold()


def _number(value) -> float | None:
    text = str(value if value is not None else "").replace(",", "").strip()
    try:
        return float(text)
    except ValueError:
        return None


# ── Sales: VGV houses by suburb ───────────────────────────────────────────────────────

def sales_files(catalogue: dict) -> dict[int, str]:
    """{last year: url} of the "Houses by suburb" time series files in a
    catalogue entry (package_show result)."""
    found = {}
    for resource in catalogue.get("resources", []):
        url = resource.get("url") or ""
        match = SALES_FILE.search(url)
        if match:
            found[int(match.group(2))] = url
    return found


def parse_suburb_medians(rows, localities: set[str]) -> dict[str, dict[int, tuple[float, str]]]:
    """{locality (upper case): {year: (median, flag)}} from the rows of a
    Houses by suburb sheet, for the localities asked for."""
    years: dict[int, int] = {}
    out: dict[str, dict[int, tuple[float, str]]] = {}
    for row in rows:
        if not row:
            continue
        first = str(row[0] or "").strip()
        if not years:
            if first.casefold() == "locality":
                for index, value in enumerate(row):
                    if isinstance(value, int) or (isinstance(value, str) and value.strip().isdigit()):
                        years[index] = int(value)
                    elif isinstance(value, str) and value.strip().casefold() == "prelim":
                        break           # the following year to date: not a full year
            continue
        name = first.upper()
        if name not in localities:
            continue
        for index, year in years.items():
            value = _number(row[index]) if index < len(row) else None
            if value is None:
                continue
            flag = str(row[index + 1] or "").strip() if index + 1 < len(row) else ""
            out.setdefault(name, {})[year] = (value, flag)
    return out


# ── Rent: DFFH quarterly median rents by LGA ──────────────────────────────────────────

def quarter_of(label) -> tuple[int, int] | None:
    """("Sep 2025") -> (2025, 3); None for anything else."""
    match = QUARTER_LABEL.match(str(label or ""))
    if not match or match.group(1).casefold() not in MONTHS:
        return None
    return int(match.group(2)), MONTHS[match.group(1).casefold()]


def latest_rent_resource(catalogue: dict) -> str | None:
    """URL of the catalogue resource for the latest quarter, judged by
    the quarter named in the resource's title."""
    best, best_url = None, None
    for resource in catalogue.get("resources", []):
        match = re.search(r"\b(Mar|Jun|June|Sep|Sept|Dec)\w*\s+(\d{4})\b", resource.get("name") or "")
        if not match:
            continue
        key = (int(match.group(2)), MONTHS[match.group(1).casefold()])
        if best is None or key > best:
            best, best_url = key, resource.get("url")
    return best_url


def parse_lga_rents(rows, lgas: set[str]) -> dict[str, dict[tuple[int, int], float]]:
    """{LGA key: {(year, quarter): median weekly rent}} from the rows of
    one sheet of the DFFH workbook, for the LGAs asked for (lga_key
    form)."""
    rows = list(rows)
    header = None
    for number, row in enumerate(rows):
        if row and any(str(v or "").strip() == "Median" for v in row):
            header = number
            break
    if header is None or header == 0:
        return {}
    quarters = {}
    for index, value in enumerate(rows[header]):
        if str(value or "").strip() == "Median":
            quarter = quarter_of(rows[header - 1][index])
            if quarter:
                quarters[index] = quarter
    out: dict[str, dict[tuple[int, int], float]] = {}
    for row in rows[header + 1:]:
        if not row or len(row) < 2 or lga_key(row[1]) not in lgas:
            continue
        for index, quarter in quarters.items():
            value = _number(row[index]) if index < len(row) else None
            if value is not None:
                out.setdefault(lga_key(row[1]), {})[quarter] = value
    return out


def calendar_year_means(quarterly: dict[tuple[int, int], float]) -> dict[int, float]:
    """Mean of the four quarterly figures of each calendar year that has
    all four."""
    by_year: dict[int, list[float]] = {}
    for (year, _), value in quarterly.items():
        by_year.setdefault(year, []).append(value)
    return {year: round(sum(v) / 4, 2) for year, v in by_year.items() if len(v) == 4}


# ── Fetcher ───────────────────────────────────────────────────────────────────────────

class VICSalesRentFetcher(BaseFetcher):

    SOURCE_NAME      = "vic_sales_rent"
    SUPPORTED_STATES = ["VIC"]

    def _catalogue(self, dataset: str) -> dict:
        import requests
        try:
            response = requests.get(CATALOGUE.format(dataset), headers=HEADERS, timeout=60)
            response.raise_for_status()
            return response.json()["result"]
        except Exception as exc:
            self.log.warning(f"  catalogue entry '{dataset}' not read: {exc}")
            return {}

    def _xlsx(self, url: str, cache_key: str) -> Path | None:
        """Download a workbook; a page is searched for its first .xlsx
        link. None if no workbook is obtained."""
        import requests
        from urllib.parse import urljoin
        path = self.download(url, cache_key, suffix=".xlsx", headers=HEADERS)
        if path is None:
            return None
        if path.read_bytes()[:2] == b"PK":
            return path
        match = XLSX_LINK.search(path.read_text(encoding="utf-8", errors="ignore"))
        path.unlink(missing_ok=True)
        if not match:
            self.log.info(f"  no workbook link on {url}")
            return None
        try:
            data = requests.get(urljoin(url, match.group(1)), headers=HEADERS, timeout=300)
            data.raise_for_status()
        except Exception as exc:
            self.log.info(f"  workbook not downloaded ({exc})")
            return None
        if not data.content.startswith(b"PK"):
            return None
        path.write_bytes(data.content)
        return path

    def _manual(self, name: str) -> Path | None:
        path = CACHE_DIR / name
        if path.exists():
            self.log.info(f"  using the copy downloaded by hand: {name}")
            return path
        return None

    def _sales(self, localities: set[str]) -> tuple[dict, list[str]]:
        import openpyxl
        files = sales_files(self._catalogue(SALES_DATASET))
        paths = []
        for last_year in sorted(files, reverse=True):
            path = self._xlsx(files[last_year], f"vgv_houses_by_suburb_{last_year}")
            if path is not None:
                paths.append((files[last_year], path))
        if not paths:
            manual = self._manual(MANUAL_SALES)
            if manual is not None:
                paths.append(("downloaded by hand", manual))

        merged: dict[str, dict[int, tuple[float, str]]] = {}
        used = []
        for source, path in paths:
            book = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                found = parse_suburb_medians(book.worksheets[0].iter_rows(values_only=True), localities)
            finally:
                book.close()
            used.append(source)
            for name, years in found.items():
                for year, entry in years.items():
                    merged.setdefault(name, {}).setdefault(year, entry)   # newest file first
        return merged, used

    def _rents(self, lgas: set[str]) -> tuple[dict, str]:
        import openpyxl
        url = latest_rent_resource(self._catalogue(RENT_DATASET))
        # Cached under the page's name, which names the quarter, so that a
        # newer quarter is downloaded rather than the cached one reused.
        key = "dffh_" + re.sub(r"\W+", "_", url.rstrip("/").rsplit("/", 1)[-1]).strip("_") if url else ""
        path = self._xlsx(url, key) if url else None
        source = url or ""
        if path is None:
            path, source = self._manual(MANUAL_RENT), "downloaded by hand"
        if path is None:
            return {}, ""
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            if RENT_SHEET not in book.sheetnames:
                self.log.warning(f"  sheet '{RENT_SHEET}' not in {path.name}")
                return {}, source
            return parse_lga_rents(book[RENT_SHEET].iter_rows(values_only=True), lgas), source
        finally:
            book.close()

    def fetch_all(self):
        towns = self.applicable_towns()
        if not towns:
            self.log.info("  no Victorian towns in towns.toml")
            return
        localities = {t.locality.upper() for t in towns if t.locality}
        sales, sales_sources = self._sales(localities) if localities else ({}, [])
        rents, rent_source = self._rents({lga_key(t.lga) for t in towns})
        if not sales_sources:
            self.result.add_warning("ALL", f"VGV houses-by-suburb file not obtained: download it by hand "
                                           f"as 1 Data/Housing/{MANUAL_SALES}")
        if not rent_source:
            self.result.add_warning("ALL", f"DFFH median rents by LGA not obtained: download it by hand "
                                           f"as 1 Data/Housing/{MANUAL_RENT}")

        out_dir = CACHE_DIR / "housing" / "regions"
        out_dir.mkdir(parents=True, exist_ok=True)
        for town in towns:
            medians = sales.get(town.locality.upper(), {}) if town.locality else {}
            rent = calendar_year_means(rents.get(lga_key(town.lga), {}))
            values = {
                "median_sales_price": {str(y): v for y, (v, _) in sorted(medians.items()) if YEAR_START <= y <= YEAR_END},
                "rent_house_3bed": {str(y): v for y, v in sorted(rent.items()) if YEAR_START <= y <= YEAR_END},
            }
            if not any(values.values()):
                self.result.add_error(town.name, f"no sales or rent figures (locality '{town.locality}', "
                                                 f"LGA '{town.lga}')")
                self.result.towns_failed.append(town.name)
                continue
            indicators = {key: {"label": label, "values": values.get(key, {})} for key, label in LABELS.items()}
            flags = {str(y): f for y, (_, f) in sorted(medians.items()) if f and YEAR_START <= y <= YEAR_END}
            out = {
                "region": town.name,
                "lga": town.lga,
                "locality": town.locality,
                "section": f"{town.lga} (LGA)",
                "source": "Valuer-General Victoria, houses by suburb; DFFH Rental Report, quarterly median rents by LGA",
                "source_url": "; ".join(s for s in sales_sources[:1] + [rent_source] if s),
                "note": (
                    f"Median Sale Price: VGV calendar-year median house price for the locality "
                    f"{town.locality or '(none set)'} -- a town, not the LGA, and houses only (the New South "
                    f"Wales rows are LGA, all dwellings). Rent: mean of the four quarterly DFFH medians for "
                    f"3-bedroom houses (new lettings) in {town.lga} LGA, complete years only. Not published, "
                    f"left blank: " + "; ".join(f"{k} ({v})" for k, v in NOT_PUBLISHED.items()) + "."
                ),
                "vgv_flags": flags,
                "indicators": indicators,
            }
            with open(out_dir / f"{town.slug}_sales_rent.json", "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2)
            for key, series in values.items():
                if series:
                    latest = max(series, key=int)
                    self.log.info(f"  [{town.name}:{key}] {min(series, key=int)}-{latest}, latest = {series[latest]}")
                else:
                    self.log.info(f"  [{town.name}:{key}] no figures")
            self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = VICSalesRentFetcher().run()
    sys.exit(0 if result.success else 1)
