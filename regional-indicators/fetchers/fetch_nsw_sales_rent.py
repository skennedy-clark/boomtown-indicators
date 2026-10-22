"""
regional-indicators/fetchers/fetch_nsw_sales_rent.py

Fetches mean and median sales price, sales number, and median weekly
rent for new bonds (House 3-bedroom, and Total 3-bedroom) for each New
South Wales town's LGA, for the town's "<Town> (LGA)" block of the
Housing sheet. The towns are the NSW study towns in towns.toml, matched
to the report's LGA names by their `lga` field; a town added to
towns.toml is fetched without a code change.

Source: NSW Department of Communities and Justice (DCJ) Rent and Sales
Report: quarterly, LGA level, since 2011. Its "New Bonds" terminology
matches the sheet's row labels.
  https://dcj.nsw.gov.au/about-us/families-and-communities-statistics/housing-rent-and-sales/rent-and-sales-report.html

File layout:
  - Each quarter has two files, a "Sales tables" workbook and a "Rent
    tables" workbook, each with an "LGA" sheet.
  - Sales LGA sheet: long format, one row per (LGA, DwellingType).
    Columns "Median Sales Price $'000s", "Mean Sales Price $'000s" and
    "Sales No.". Prices are in thousands and are multiplied by 1000 to
    match the sheet, which holds whole dollars (282034.72, not 282.03).
    A town's row is the one with its LGA and DwellingType="Total" (all
    dwelling types); for example Narrabri Q1 2026 gives median $435k,
    mean $458k and 51 sales.
  - Rent LGA sheet: long format, one row per (LGA, DwellingType,
    Bedrooms). Column "Median Weekly Rent for New Bonds $". Two rows
    are read: DwellingType="House" with Bedrooms="3 Bedrooms" (the
    sheet's "House 3-Bed" row), and DwellingType="Total" with
    Bedrooms="3 Bedrooms" (the sheet's "Total" row, which is the
    all-types 3-bedroom figure, not an all-bedrooms figure).

File naming: file names are inconsistent between quarters, including
within one calendar year ("issue-152-sales-tables-mar-2025.xlsx"
against "sales_tables_june_2025_quarter.xlsx": underscores or hyphens,
with or without an issue number), so they cannot be constructed. They
are discovered instead: the landing page links the current quarter and
the "Previous rent and sales reports" page (ARCHIVE_PAGE) links every
earlier quarter published as a spreadsheet (from June 2017 for sales
and September 2017 for rent; earlier issues are PDF only). The kind and
quarter are read from each file name ("rent ... september 2017" is the
rent table for the September quarter of 2017). Rent and sales are not
published for the same quarter, because their reporting lags differ
(for example Rent = June 2026 with Sales = March 2026). If neither page
can be read, the files listed in QUARTER_FILES are used.

Aggregation: as for region-level Housing data in fetch_qgso_housing.py
(_fetch_regions). Price and rent are the mean of the four quarterly
values for the year; Sales No. is the sum of the four quarterly counts.
Complete years only (all four quarters present).

Output: cache/housing/regions/<town slug>_sales_rent.json (for example
narrabri_sales_rent.json), read by update_housing.py.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR, YEAR_START, YEAR_END

try:
    import requests
    import openpyxl
except ImportError:
    raise ImportError("pip install requests openpyxl")


DAM_BASE = (
    "https://dcj.nsw.gov.au/content/dam/dcj/dcj-website/documents/"
    "about-us/families-and-communities-statistics/housing-and-rent-sales"
)
LANDING_PAGE = (
    "https://dcj.nsw.gov.au/about-us/families-and-communities-statistics/"
    "housing-rent-and-sales/rent-and-sales-report.html"
)
ARCHIVE_PAGE = (
    "https://dcj.nsw.gov.au/about-us/families-and-communities-statistics/"
    "housing-rent-and-sales/previous-rent-and-sales-reports.html"
)
SITE = "https://dcj.nsw.gov.au"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
FILE_LINK = re.compile(r'href="([^"]*?((?:sales|rent)[-_]tables[-_][a-z]+[-_]\d{4}[^"/]*?\.xlsx))"', re.IGNORECASE)
FILE_NAME = re.compile(r"(sales|rent)[-_]tables[-_]([a-z]+)[-_](\d{4})", re.IGNORECASE)


# Used only if the DCJ pages cannot be read (the files are otherwise
# discovered; see the module docstring).
QUARTER_FILES = {
    # (year, quarter): (sales_filename, rent_filename)
    (2025, 1): ("issue-152-sales-tables-mar-2025.xlsx", "issue-151-rent-tables-mar-2025.xlsx"),
    (2025, 2): ("sales_tables_june_2025_quarter.xlsx", "Rent_tables_June_2025_quarter.xlsx"),
    (2025, 3): ("sales_tables_september_2025_quarter.xlsx", "rent_tables_september_2025_quarter.xlsx"),
    (2025, 4): ("sales-tables-december-2025-quarter.xlsx", "rent_tables_december_2025_quarter.xlsx"),
}

SALES_LABELS = {
    "mean_sales_price":   "Mean Sales Price $",
    "median_sales_price": "Median Sales Price $",
    "sales_no":           "Sales No.",
}
RENT_LABELS = {
    "rent_house_3bed":  "Median Weekly Rent for New Bonds $ (House, 3 Bedrooms)",
    "rent_total_3bed":  "Median Weekly Rent for New Bonds $ (Total, 3 Bedrooms)",
}


def report_files(html: str) -> dict[tuple[str, int, int], str]:
    """{(kind, year, quarter): absolute URL} for the sales and rent table
    spreadsheets linked from a DCJ page. The first link found for a
    quarter is kept."""
    found: dict[tuple[str, int, int], str] = {}
    for href, name in FILE_LINK.findall(html):
        match = FILE_NAME.search(name)
        quarter = _month_to_quarter(match.group(2)) if match else 0
        if not quarter:
            continue
        url = href if href.startswith("http") else SITE + href
        found.setdefault((match.group(1).lower(), int(match.group(3)), quarter), url)
    return found


def _discover_report_files(log) -> dict[tuple[str, int, int], str]:
    """The table spreadsheets linked from the landing page and the
    archive page ({} if neither can be read)."""
    found: dict[tuple[str, int, int], str] = {}
    for page in (LANDING_PAGE, ARCHIVE_PAGE):
        try:
            resp = requests.get(page, headers=HEADERS, timeout=60)
            resp.raise_for_status()
        except requests.RequestException as exc:
            log.warning(f"  Could not read {page}: {exc}")
            continue
        for key, url in report_files(resp.text).items():
            found.setdefault(key, url)
    if found:
        quarters = sorted({(y, q) for _, y, q in found})
        log.info(f"  Found {len(found)} table files, {quarters[0][0]} Q{quarters[0][1]} to "
                 f"{quarters[-1][0]} Q{quarters[-1][1]}")
    return found


def _month_to_quarter(month_name: str) -> int:
    m = month_name.lower()
    if m in ("march", "mar"): return 1
    if m in ("june", "jun"): return 2
    if m in ("september", "sep", "sept"): return 3
    if m in ("december", "dec"): return 4
    return 0


def lga_key(name) -> str:
    """An LGA name as compared with the report's ("Narrabri (NSW)" and
    "narrabri" both give "narrabri")."""
    return re.sub(r"\s*\([^)]*\)\s*$", "", str(name or "")).strip().casefold()


def _header(ws) -> tuple[dict[str, int], int] | None:
    """({column heading, normalised: index}, header row number) of an LGA
    table. The LGA column's heading has been worded "Local Government Area
    (LGA)", "Local Government Areas" and "LGA (Local Government Area)";
    headings are compared in lower case with spaces and line breaks
    collapsed."""
    for number, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), start=1):
        if len(row) > 3 and isinstance(row[3], str) and re.search(r"local government|\blga\b", row[3], re.IGNORECASE):
            return {re.sub(r"\s+", " ", str(v)).strip().lower(): i for i, v in enumerate(row) if v}, number
    return None


def _column(cols: dict[str, int], *starts: str) -> int | None:
    for heading, index in cols.items():
        if any(heading.startswith(start) for start in starts):
            return index
    return None


def _parse_sales_file(path: Path, log, lgas: set[str]) -> dict[str, dict]:
    """{LGA key: {"median_sales_price", "mean_sales_price", "sales_no"}}
    for the LGAs in `lgas` (lga_key form) with a full set of figures."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb["LGA"]
    except Exception as exc:
        log.warning(f"  Could not open/parse {path.name}: {exc}")
        return {}
    found = _header(ws)
    if not found:
        log.warning(f"  Could not find the header row in {path.name}")
        return {}
    cols, header_number = found
    median_col = _column(cols, "median sales price")
    mean_col = _column(cols, "mean sales price")
    no_col = _column(cols, "sales no")
    if median_col is None or mean_col is None or no_col is None:
        log.warning(f"  Column headers didn't match expected names in {path.name}")
        return {}

    def _num(v):
        if v in (None, "-", "s"):
            return None
        try:
            return float(str(v).replace(",", ""))
        except ValueError:
            return None

    out: dict[str, dict] = {}
    for row in ws.iter_rows(min_row=header_number + 1, values_only=True):
        if len(row) <= 4 or lga_key(row[3]) not in lgas or row[4] != "Total":
            continue
        # A figure DCJ suppresses ("s": 30 or fewer sales) is left out; the
        # other figures for the quarter are kept.
        median, mean, no = _num(row[median_col]), _num(row[mean_col]), _num(row[no_col])
        values = {}
        if median is not None:
            values["median_sales_price"] = median * 1000
        if mean is not None:
            values["mean_sales_price"] = mean * 1000
        if no is not None:
            values["sales_no"] = no
        if values:
            out[lga_key(row[3])] = values
    return out


def _parse_rent_file(path: Path, log, lgas: set[str]) -> dict[str, dict]:
    """{LGA key: {"rent_house_3bed", "rent_total_3bed"}} for the LGAs in
    `lgas` (lga_key form)."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb["LGA"]
    except Exception as exc:
        log.warning(f"  Could not open/parse {path.name}: {exc}")
        return {}
    found = _header(ws)
    if not found:
        log.warning(f"  Could not find the header row in {path.name}")
        return {}
    cols, header_number = found
    median_col = _column(cols, "median weekly rent")
    if median_col is None:
        log.warning(f"  Rent column header didn't match expected name in {path.name}")
        return {}

    def _num(v):
        if v in (None, "-", "s"):
            return None
        try:
            return float(str(v).replace(",", ""))
        except ValueError:
            return None

    out: dict[str, dict] = {}
    for row in ws.iter_rows(min_row=header_number + 1, values_only=True):
        if len(row) <= 5 or lga_key(row[3]) not in lgas or row[5] != "3 Bedrooms":
            continue
        if row[4] == "House":
            out.setdefault(lga_key(row[3]), {})["rent_house_3bed"] = _num(row[median_col])
        elif row[4] == "Total":
            out.setdefault(lga_key(row[3]), {})["rent_total_3bed"] = _num(row[median_col])
    return out


class NSWSalesRentFetcher(BaseFetcher):

    SOURCE_NAME      = "nsw_sales_rent"
    SUPPORTED_STATES = ["NSW"]

    def fetch_all(self):
        towns = self.applicable_towns()
        if not towns:
            self.log.info("  no New South Wales towns in towns.toml")
            return
        lgas = {lga_key(t.lga): t for t in towns if t.lga}
        files = _discover_report_files(self.log)
        if not files:
            self.log.warning("  Using the file list in QUARTER_FILES")
            for (yr, q), (sales_fn, rent_fn) in QUARTER_FILES.items():
                files[("sales", yr, q)] = f"{DAM_BASE}/{sales_fn}"
                files[("rent", yr, q)] = f"{DAM_BASE}/{rent_fn}"

        # (year, quarter) -> LGA key -> {indicator: value}
        by_quarter: dict[tuple[int, int], dict[str, dict]] = {}
        parsers = {"sales": _parse_sales_file, "rent": _parse_rent_file}
        for (kind, yr, q), url in sorted(files.items(), key=lambda kv: (kv[0][1], kv[0][2], kv[0][0])):
            if not (YEAR_START <= yr <= YEAR_END + 1):
                continue
            path = self.download(url, f"dcj_{kind}_{yr}q{q}", suffix=".xlsx", headers=HEADERS)
            if path is None:
                self.result.add_warning(f"{yr} Q{q}", f"could not download the {kind} tables ({url})")
                continue
            for lga, values in parsers[kind](path, self.log, set(lgas)).items():
                by_quarter.setdefault((yr, q), {}).setdefault(lga, {}).update(values)

        if not by_quarter:
            self.result.add_error("ALL", "No sales/rent data assembled from any quarter")
            return
        for lga, town in lgas.items():
            self._write_town(town, {quarter: data[lga] for quarter, data in by_quarter.items() if lga in data})

    def _write_town(self, town, by_quarter: dict[tuple[int, int], dict]) -> None:
        if not by_quarter:
            self.result.add_error(town.name, f"LGA '{town.lga}' not found in any quarter's tables")
            return

        # Aggregate by calendar year: mean for price and rent, sum for the
        # sales count.
        by_year: dict[int, dict[str, list]] = {}
        for (yr, q), vals in by_quarter.items():
            if not (YEAR_START <= yr <= YEAR_END):
                continue
            for key, v in vals.items():
                if v is not None:
                    by_year.setdefault(yr, {}).setdefault(key, []).append(v)

        indicators = {}
        for key in list(SALES_LABELS) + list(RENT_LABELS):
            label = SALES_LABELS.get(key) or RENT_LABELS.get(key)
            values = {}
            for yr, series in sorted(by_year.items()):
                vals = series.get(key, [])
                if len(vals) != 4:
                    continue  # complete years (4 quarters) only
                if key == "sales_no":
                    values[str(yr)] = int(sum(vals))
                else:
                    values[str(yr)] = round(sum(vals) / len(vals), 2)
            indicators[key] = {"label": label, "values": values}

        out = {
            "region":     town.name,
            "lga":        town.lga,
            "section":    f"{town.name} (LGA)",
            "source":     "NSW Department of Communities and Justice (DCJ), Rent and Sales Report — LGA tables",
            "source_url": LANDING_PAGE,
            "note": (
                "Calendar-year means of the 4 quarterly Median/Mean Sales Price and "
                "Median Weekly Rent readings, and SUM of the 4 quarterly Sales No. "
                "counts -- complete years only (all 4 quarters present). Sales Price "
                "converted from the source's $'000s to whole dollars to match this "
                "sheet's existing convention. Rent/Sales are reported with different "
                "lags and are not always on the same quarter -- merged independently "
                "by their own reporting quarter."
            ),
            "indicators": indicators,
        }
        out_dir = CACHE_DIR / "housing" / "regions"
        out_dir.mkdir(parents=True, exist_ok=True)
        with open(out_dir / f"{town.slug}_sales_rent.json", "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        for key, entry in indicators.items():
            vals = entry["values"]
            if vals:
                latest = max(vals, key=int)
                self.log.info(f"  [{town.name}:{key}] {len(vals)} years, latest ({latest}) = {vals[latest]}")
        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = NSWSalesRentFetcher().run()
    sys.exit(0 if result.success else 1)
