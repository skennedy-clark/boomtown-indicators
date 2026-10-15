"""
regional-indicators/fetchers/fetch_fuel_qld.py

Fetches annual average unleaded petrol prices for Queensland towns from
the Queensland Government's fuel price reporting open data, for towns
that the RACQ Annual Fuel Price Report (fetch_fuel.py) does not name.
RACQ stays the source for every town it covers; the writer uses these
figures only for a Fuel block that RACQ has no location for.

Source: data.qld.gov.au datasets "fuel-price-reporting-2019-2022",
"fuel-price-reporting" (2023) and "fuel-price-reporting-<year>" (2024
onwards): one CSV per month, one row each time a site changes a price:
    SiteId, Site_Name, Site_Brand, Sites_Address_Line_1, Site_Suburb,
    Site_State, Site_Post_Code, Site_Latitude, Site_Longitude,
    Fuel_Type, Price, TransactionDateutc
Price is in tenths of a cent (1999 = 199.9 c/L). The file list is read
from each dataset (CKAN package_show); the datasets for future years are
found by name, so no code change is needed as years pass.

Method: fuel type "Unleaded", the sites in the town's postcodes. A
site's price holds until its next change, which gives each site a daily
price; a site with prices on at least half the days of a year counts for
that year, and the town's figure is the mean of those sites' yearly
means. A price more than 40% from the median of all the year's changes
in the towns fetched is ignored (keying errors such as 9999 occur and
would otherwise hold until the site's next change).

Checked against RACQ's published annual averages, 2019-2025:
Goondiwindi and Roma within 1.4 c/L in every year (0.0-0.2 in most),
Dalby within 2.1 except 4.9 in 2023, Emerald within 1 except 3.7 and
2.5 in 2023-24. Data start in 2019.

Only complete years are produced (a year whose December file has been
published). Monthly files are cached; each run downloads only months it
does not have.

The portal sits behind a web firewall that refuses some cloud servers;
from a desktop it downloads normally.

Output: cache/fuel/qld_open_data_ulp_annual.json
    {"source", "report_year", "locations": {town: {"2019": 150.7, ...}},
     "sites": {town: {"2019": 5, ...}}}
Read by transform/xlsx_update/update_fuel.py.
"""

from __future__ import annotations

import csv
import io
import json
import re
import statistics
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher

API = "https://www.data.qld.gov.au/api/3/action/package_show"
FIRST_YEAR = 2019
FUEL_TYPE = "unleaded"
OUTLIER_SHARE = 0.4
MIN_DAYS = 183
OUTPUT_NAME = "qld_open_data_ulp_annual.json"
SOURCE = "Queensland Government fuel price reporting (data.qld.gov.au)"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def package_names(last_year: int) -> list[str]:
    return ["fuel-price-reporting-2019-2022", "fuel-price-reporting"] + [
        f"fuel-price-reporting-{year}" for year in range(2024, last_year + 1)]


def file_month(name: str) -> tuple[int, int] | None:
    """(year, month) from a resource name such as "Queensland Fuel Prices
    December 2019" or a file name such as "fuel-prices-2025-06-changes-only"."""
    months = ["january", "february", "march", "april", "may", "june", "july",
              "august", "september", "october", "november", "december"]
    text = name.lower()
    year = re.search(r"(20\d\d)", text)
    if not year:
        return None
    for i, month in enumerate(months, start=1):
        if month in text or month[:3] + str(year.group(1))[2:] in text.replace(" ", ""):
            return int(year.group(1)), i
    numeric = re.search(r"20\d\d-(\d\d)", text)
    if numeric and 1 <= int(numeric.group(1)) <= 12:
        return int(year.group(1)), int(numeric.group(1))
    return None


def _column(header: list[str], *names: str) -> int:
    lowered = [h.strip().lower().replace(" ", "_") for h in header]
    for name in names:
        if name in lowered:
            return lowered.index(name)
    raise KeyError(f"none of {names} in the header")


def _time(text: str) -> datetime | None:
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%dT%H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(text.strip()[:19], fmt)
        except ValueError:
            continue
    return None


def read_changes(text: str, postcodes: set[str], changes: dict) -> None:
    """Add the unleaded price changes for `postcodes` in one monthly CSV to
    `changes`: {(postcode, site): [(time, cents per litre), ...]}."""
    rows = csv.reader(io.StringIO(text))
    header = next(rows, None)
    if not header:
        return
    c_site = _column(header, "siteid", "site_id")
    c_postcode = _column(header, "site_post_code", "site_postcode", "postcode")
    c_fuel = _column(header, "fuel_type", "fueltype")
    c_price = _column(header, "price")
    c_time = _column(header, "transactiondateutc", "transaction_date_utc", "transactiondate", "transaction_date")
    last = max(c_site, c_postcode, c_fuel, c_price, c_time)
    for row in rows:
        if len(row) <= last:
            continue
        postcode = row[c_postcode].strip().split(".")[0]
        if postcode not in postcodes or row[c_fuel].strip().lower() != FUEL_TYPE:
            continue
        when = _time(row[c_time])
        try:
            price = float(row[c_price])
        except ValueError:
            continue
        if when:
            changes[(postcode, row[c_site].strip())].append((when, price / 10 if price > 1000 else price))


def annual_means(changes: dict, years: range) -> dict[tuple[str, int], tuple[float, int]]:
    """{(postcode, year): (mean cents per litre, number of sites)}."""
    by_year = defaultdict(list)
    for points in changes.values():
        for when, price in points:
            by_year[when.year].append(price)
    median = {year: statistics.median(prices) for year, prices in by_year.items()}

    daily = defaultdict(list)
    for (postcode, site), points in changes.items():
        points = sorted((w, p) for w, p in points if abs(p / median[w.year] - 1) <= OUTLIER_SHARE)
        for i, (start, price) in enumerate(points):
            end = points[i + 1][0].date() if i + 1 < len(points) else date(start.year + 1, 1, 1)
            day = start.date()
            while day < end:
                if day.year in years:
                    daily[(postcode, site, day.year)].append(price)
                day += timedelta(days=1)

    site_means = defaultdict(list)
    for (postcode, site, year), prices in daily.items():
        if len(prices) >= MIN_DAYS:
            site_means[(postcode, year)].append(statistics.mean(prices))
    return {key: (statistics.mean(v), len(v)) for key, v in site_means.items()}


class QldFuelOpenDataFetcher(BaseFetcher):

    SOURCE_NAME      = "fuel_qld"
    SUPPORTED_STATES = ["QLD"]

    def fetch_all(self):
        import requests
        from config import YEAR_END

        towns = [t for t in self.applicable_towns() if t.postcodes or t.postcode]
        postcodes_of = {t.name: {str(p) for p in (t.postcodes or [t.postcode])} for t in towns}
        wanted = set().union(*postcodes_of.values()) if postcodes_of else set()
        if not wanted:
            self.log.info("  no Queensland towns with postcodes")
            return

        files: dict[tuple[int, int], str] = {}
        for package in package_names(YEAR_END + 1):
            try:
                response = requests.get(API, params={"id": package}, headers=HEADERS, timeout=60)
                if response.status_code == 404:
                    continue
                resources = response.json()["result"]["resources"]
            except Exception as exc:
                self.result.add_warning(package, f"could not read the dataset: {exc}")
                continue
            for resource in resources:
                if (resource.get("format") or "").lower() != "csv":
                    continue
                month = file_month(resource.get("name", "")) or file_month(resource.get("url", ""))
                if month and month[0] >= FIRST_YEAR:
                    files.setdefault(month, resource["url"])
        if not files:
            self.result.add_error("ALL", "no monthly files found on data.qld.gov.au")
            return

        complete = [y for y in range(FIRST_YEAR, YEAR_END + 1) if (y, 12) in files]
        if not complete:
            self.result.add_error("ALL", "no complete year of monthly files")
            return
        years = range(FIRST_YEAR, max(complete) + 1)

        changes: dict = defaultdict(list)
        for (year, month), url in sorted(files.items()):
            if year > years.stop - 1:
                continue
            path = self.download(url, f"qld_fuel_prices_{year}_{month:02d}", suffix=".csv", headers=HEADERS)
            if path is None:
                self.result.add_warning(f"{year}-{month:02d}", f"could not download {url}")
                continue
            try:
                read_changes(path.read_bytes().decode("utf-8-sig", errors="replace"), wanted, changes)
            except KeyError as exc:
                self.result.add_warning(f"{year}-{month:02d}", f"unexpected layout: {exc}")

        by_postcode = annual_means(changes, years)
        locations, sites = {}, {}
        for town, codes in postcodes_of.items():
            series, counts = {}, {}
            for year in years:
                parts = [by_postcode[(pc, year)] for pc in codes if (pc, year) in by_postcode]
                n = sum(count for _, count in parts)
                if n:
                    series[str(year)] = round(sum(mean * count for mean, count in parts) / n, 1)
                    counts[str(year)] = n
            if series:
                locations[town] = series
                sites[town] = counts
                self.result.towns_ok.append(town)
                self.log.info(f"  {town}: {min(series)}-{max(series)}, "
                              f"{series[max(series)]} c/L in {max(series)} ({counts[max(series)]} sites)")
            else:
                self.result.add_warning(town, "no unleaded prices reported in its postcodes")

        out_dir = Path(__file__).parent.parent / "cache" / "fuel"
        out_dir.mkdir(parents=True, exist_ok=True)
        with open(out_dir / OUTPUT_NAME, "w", encoding="utf-8") as f:
            json.dump({"source": SOURCE, "report_year": years.stop - 1,
                       "locations": locations, "sites": sites}, f, indent=2)
