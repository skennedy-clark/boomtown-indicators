"""
fetchers/fetch_narrabri_sales_rent.py
--------------------------------------
Fetches Mean/Median Sales Price, Sales No., and Median Weekly Rent for New
Bonds (House 3-bed, and Total/all-types 3-bed) for Narrabri (LGA), for the
Housing sheet's Narrabri block (rows 132-141). Confirmed via direct
inspection of the real starting file that this block needs exactly these
five indicators, matching row labels and sub-rows exactly.

SOURCE: NSW Department of Communities and Justice (DCJ) Rent and Sales
Report -- quarterly, LGA-level, since 2011. Confirmed live 2026-10-01 as
the real, current, authoritative source (matches the sheet's own "New
Bonds" terminology exactly, almost certainly the same source the sheet
was originally built from).

STRUCTURE, confirmed by downloading real files and inspecting directly,
not assumed:
  - Each quarter publishes TWO separate files: a "Sales tables" workbook
    and a "Rent tables" workbook, both with an "LGA" sheet.
  - Sales LGA sheet: long/tidy format, one row per (LGA, DwellingType).
    Columns confirmed: "Median Sales Price $'000s", "Mean Sales Price
    $'000s", "Sales No." -- values in THOUSANDS, multiplied by 1000 below
    to match the sheet's existing full-dollar convention (confirmed: the
    sheet's own history is in whole dollars, e.g. 282034.72, not 282.03).
    Narrabri's relevant row has DwellingType="Total" (all dwelling types
    combined) -- confirmed live: Q1 2026 gives Median=$435k, Mean=$458k,
    Sales No.=51.
  - Rent LGA sheet: long/tidy format, one row per (LGA, DwellingType,
    Bedrooms). Column confirmed: "Median Weekly Rent for New Bonds $".
    TWO rows needed, confirmed live: DwellingType="House" + Bedrooms="3
    Bedrooms" (matches the sheet's "House 3-Bed" row exactly), and
    DwellingType="Total" + Bedrooms="3 Bedrooms" (matches the sheet's
    bare "Total" row -- confirmed by value: both read $500 in the same
    live quarter, consistent with the "Total" row being the ALL-types
    3-bedroom figure, not an all-bedrooms figure).

FILE NAMING, confirmed inconsistent across quarters -- live-tested
multiple real historical filenames, including several different naming
conventions in the same calendar year (e.g. "issue-152-sales-tables-
mar-2025.xlsx" vs "sales_tables_june_2025_quarter.xlsx" -- underscores vs
hyphens, issue-number-prefixed vs not). No single predictable pattern
exists, so this fetcher does NOT try to guess historical filenames the
way other fetchers in this project discover a current release slug.
Instead:
  - Known-good historical filenames (2025, all 4 quarters, both sales and
    rent) are hardcoded as a confirmed-working backfill -- every one of
    the 8 live-tested directly (HTTP 200), not assumed from the list they
    came from.
  - The CURRENT quarter's files are discovered live each run, by scraping
    the live "rent-and-sales-report.html" landing page, which reliably
    links the latest quarter under a predictable "{type}-tables-
    {month}-{year}-quarter.xlsx" pattern (confirmed live 2026-10-01: Rent
    = June 2026, Sales = March 2026 -- rent and sales are NOT
    synchronised to the same quarter, reflecting their different
    reporting lags, confirmed from the page itself).
  - update each cycle: once a full NEW calendar year's 4 quarters exist,
    add them to QUARTER_FILES the same way 2025's were added, by visiting
    the "Previous rent and sales reports" archive page and copying the
    real filenames -- there is no way to discover these automatically
    given the inconsistent naming, confirmed above.

AGGREGATION: matches this project's established convention for region-
level Housing data (see fetch_qgso_housing.py's _fetch_regions) -- price
and rent are the MEAN of the 4 quarterly values for the year; Sales No.
is the SUM of the 4 quarterly counts (a count, not a rate, so summing is
correct the same way it is for QLD's housing sales count and building
approvals elsewhere in this project). Complete years only (all 4 quarters
present).
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

NARRABRI_NAME = "Narrabri"

# update each cycle -- known-good historical filenames, live-tested
# individually (HTTP 200) on 2026-10-01, not assumed from wherever this
# list originally came from. Add new quarters here once a full calendar
# year is available; see the module docstring for why this can't be
# discovered automatically.
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


def _discover_current_quarter_files(log) -> dict[str, str] | None:
    """Scrape the live landing page for the current quarter's real
    filenames -- confirmed live 2026-10-01 these appear as direct /content/
    dam/... links under predictable button text ("Rent tables <Month>
    <Year> quarter" / "Sales tables <Month> <Year> quarter"). Returns
    {"sales": url, "rent": url} or None if the page structure doesn't
    match (falls back to the hardcoded QUARTER_FILES backfill only).
    """
    try:
        resp = requests.get(LANDING_PAGE, timeout=30)
        resp.raise_for_status()
    except requests.RequestException as exc:
        log.warning(f"  Could not fetch the landing page: {exc}")
        return None

    found = {}
    for kind in ("rent", "sales"):
        m = re.search(
            rf'/content/dam/dcj/dcj-website/documents/about-us/families-and-communities-statistics/'
            rf'housing-and-rent-sales/({kind}-tables-[a-z]+-\d{{4}}-quarter\.xlsx)',
            resp.text,
            re.IGNORECASE,
        )
        if m:
            found[kind] = f"{DAM_BASE}/{m.group(1)}"
    if len(found) == 2:
        log.info(f"  Discovered current quarter files live: {found}")
        return found
    log.warning("  Could not find both current-quarter links on the landing page")
    return None


def _month_to_quarter(month_name: str) -> int:
    m = month_name.lower()
    if m in ("march", "mar"): return 1
    if m in ("june", "jun"): return 2
    if m in ("september", "sep"): return 3
    if m in ("december", "dec"): return 4
    return 0


def _parse_sales_file(path: Path, log) -> dict | None:
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb["LGA"]
    except Exception as exc:
        log.warning(f"  Could not open/parse {path.name}: {exc}")
        return None
    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=15, values_only=True):
        if row[3] == "Local Government Area (LGA)":
            header_row = row
            break
    if not header_row:
        log.warning(f"  Could not find the header row in {path.name}")
        return None
    cols = {v: i for i, v in enumerate(header_row) if v}
    median_col = cols.get("Median Sales Price\n$'000s")
    mean_col = cols.get("Mean Sales Price\n$'000s")
    no_col = cols.get("Sales\nNo.")
    if median_col is None or mean_col is None or no_col is None:
        log.warning(f"  Column headers didn't match expected names in {path.name}")
        return None

    for row in ws.iter_rows(min_row=1, values_only=True):
        if row[3] == NARRABRI_NAME and row[4] == "Total":
            def _num(v):
                if v in (None, "-", "s"):
                    return None
                try:
                    return float(str(v).replace(",", ""))
                except ValueError:
                    return None
            median, mean, no = _num(row[median_col]), _num(row[mean_col]), _num(row[no_col])
            if median is None or mean is None or no is None:
                return None
            return {"median_sales_price": median * 1000, "mean_sales_price": mean * 1000, "sales_no": no}
    return None


def _parse_rent_file(path: Path, log) -> dict | None:
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb["LGA"]
    except Exception as exc:
        log.warning(f"  Could not open/parse {path.name}: {exc}")
        return None
    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=15, values_only=True):
        if row[3] == "Local Government Area (LGA)":
            header_row = row
            break
    if not header_row:
        log.warning(f"  Could not find the header row in {path.name}")
        return None
    cols = {v: i for i, v in enumerate(header_row) if v}
    median_col = cols.get("Median Weekly Rent for New Bonds\n$")
    if median_col is None:
        log.warning(f"  Rent column header didn't match expected name in {path.name}")
        return None

    def _num(v):
        if v in (None, "-", "s"):
            return None
        try:
            return float(str(v).replace(",", ""))
        except ValueError:
            return None

    out = {}
    for row in ws.iter_rows(min_row=1, values_only=True):
        if row[3] != NARRABRI_NAME or row[5] != "3 Bedrooms":
            continue
        if row[4] == "House":
            out["rent_house_3bed"] = _num(row[median_col])
        elif row[4] == "Total":
            out["rent_total_3bed"] = _num(row[median_col])
    return out if out else None


class NarrabriSalesRentFetcher(BaseFetcher):

    SOURCE_NAME      = "narrabri_sales_rent"
    SUPPORTED_STATES = ["NSW"]

    def fetch_all(self):
        quarters = dict(QUARTER_FILES)  # (year, quarter) -> (sales_fn, rent_fn)

        current = _discover_current_quarter_files(self.log)
        # The current page's sales/rent quarters aren't necessarily the
        # same quarter (confirmed live: rent lags less than sales) -- each
        # is merged in under its OWN reporting quarter, parsed from its
        # filename, rather than assumed to match.
        current_urls: dict[tuple[int, int], dict[str, str]] = {}
        if current:
            for kind, url in current.items():
                m = re.search(r'-([a-z]+)-(\d{4})-quarter\.xlsx', url, re.IGNORECASE)
                if m:
                    q = _month_to_quarter(m.group(1))
                    yr = int(m.group(2))
                    if q:
                        current_urls.setdefault((yr, q), {})[kind] = url

        by_quarter_data: dict[tuple[int, int], dict] = {}

        for (yr, q), (sales_fn, rent_fn) in quarters.items():
            sales_url = f"{DAM_BASE}/{sales_fn}"
            rent_url = f"{DAM_BASE}/{rent_fn}"
            sales_path = self.download(sales_url, f"dcj_sales_{yr}q{q}", suffix=".xlsx")
            rent_path = self.download(rent_url, f"dcj_rent_{yr}q{q}", suffix=".xlsx")
            sales_data = _parse_sales_file(sales_path, self.log) if sales_path else None
            rent_data = _parse_rent_file(rent_path, self.log) if rent_path else None
            if sales_data or rent_data:
                merged = {}
                merged.update(sales_data or {})
                merged.update(rent_data or {})
                by_quarter_data[(yr, q)] = merged
                self.log.info(f"  [{yr} Q{q}] parsed: {merged}")

        for (yr, q), urls in current_urls.items():
            if (yr, q) in by_quarter_data:
                continue  # already have this quarter from the hardcoded backfill
            merged = {}
            if "sales" in urls:
                p = self.download(urls["sales"], f"dcj_sales_{yr}q{q}", suffix=".xlsx")
                if p:
                    merged.update(_parse_sales_file(p, self.log) or {})
            if "rent" in urls:
                p = self.download(urls["rent"], f"dcj_rent_{yr}q{q}", suffix=".xlsx")
                if p:
                    merged.update(_parse_rent_file(p, self.log) or {})
            if merged:
                by_quarter_data[(yr, q)] = merged
                self.log.info(f"  [{yr} Q{q}] (live-discovered) parsed: {merged}")

        if not by_quarter_data:
            self.result.add_error("ALL", "No Narrabri sales/rent data assembled from any quarter")
            return

        # Aggregate by calendar year -- mean for price/rent, sum for sales count
        by_year: dict[int, dict[str, list]] = {}
        for (yr, q), vals in by_quarter_data.items():
            if not (YEAR_START <= yr <= YEAR_END):
                continue
            for key, v in vals.items():
                if v is not None:
                    by_year.setdefault(yr, {}).setdefault(key, []).append(v)

        all_keys = list(SALES_LABELS) + list(RENT_LABELS)
        indicators = {}
        for key in all_keys:
            label = SALES_LABELS.get(key) or RENT_LABELS.get(key)
            values = {}
            for yr, series in by_year.items():
                vals = series.get(key, [])
                if len(vals) != 4:
                    continue  # complete years (4 quarters) only
                if key == "sales_no":
                    values[str(yr)] = int(sum(vals))
                else:
                    values[str(yr)] = round(sum(vals) / len(vals), 2)
            indicators[key] = {"label": label, "values": values}

        out = {
            "region":     NARRABRI_NAME,
            "section":    "Narrabri",
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
        with open(out_dir / "narrabri_sales_rent.json", "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        for key, entry in indicators.items():
            vals = entry["values"]
            if vals:
                latest = max(vals, key=int)
                self.log.info(f"  [Narrabri:{key}] {len(vals)} years, latest ({latest}) = {vals[latest]}")
        self.result.towns_ok.append("Narrabri:sales_rent")


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = NarrabriSalesRentFetcher().run()
    sys.exit(0 if result.success else 1)
