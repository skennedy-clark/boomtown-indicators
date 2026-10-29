"""
regional-indicators/fetchers/fetch_crime_csa.py

Fetches Victorian recorded crime from the Crime Statistics Agency (CSA)
and converts it to the categories and units of the Crime sheet's New
South Wales part (rates per 1,000 persons, calendar years), so that
update_crime.py writes it without change.

Source: CSA "Criminal incidents by LGA" data tables, one workbook per
quarter, each holding the ten years ending in that quarter. The
December workbooks give calendar years. They are listed in the Victorian
Government data catalogue (dataset "criminal-incident"):
    https://discover.data.vic.gov.au/api/3/action/package_show?id=criminal-incident
and served from https://files.crimestatistics.vic.gov.au/ (for example
.../2026-03/Data_Tables_LGA_Criminal_Incidents_Year_Ending_December_2025.xlsx).
The CSA website itself refuses automated requests; the files host does
not.

Sheet "Table 02": one row per year, LGA and offence subgroup, with
"Incidents Recorded" and "LGA Rate per 100,000 population" (CSA's own
population figures). Rates of the subgroups in a category are added and
divided by 100 to give a rate per 1,000.

Incidents, not offences: BOCSAR counts criminal incidents, and so does
this table (an incident with several offences is counted once, under
its principal offence).

Categories. CSA uses the Victorian offence classification, BOCSAR its
own; the subdivisions are grouped into the eight New South Wales rows:
    Assault                              A20 Assault and related offences
    Robbery                              A50 Robbery
    Other offences against the person    A10, A30, A40, A60, A70, A80
    Theft                                B30 Burglary/break and enter, B40 Theft,
                                         B50 Deception, B60 Bribery
    Malicious damage to property         B10 Arson, B20 Property damage
    Drug offences                        C (all)
    Transport regulatory offences        F20 Transport regulation offences
    Other offences                       D, E, F10, F30, F90
The grouping is close but not identical to BOCSAR's (for example, BOCSAR
counts fraud separately from theft in its detailed categories); compare
Victorian towns with each other, and with New South Wales towns only
with that caveat.

History: the latest December workbook gives ten years. Older December
workbooks are added, newest first, to reach earlier years; a year is
taken from the newest workbook that holds it (CSA revises earlier
years). Workbooks that cannot be downloaded are skipped and reported.

Geography: the town's `lga` in towns.toml, as CSA names it ("Greater
Shepparton", "Wellington"). Yarram is a small part of Wellington LGA,
so its figures describe the whole LGA.

Output: cache/crime/<slug>_crime_csa.json per Victorian town, in the
schema of fetch_crime_bocsar.py. The downloaded workbooks are kept in
cache/crime/csa/.
"""

from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR, YEAR_START, YEAR_END

CATALOGUE = "https://discover.data.vic.gov.au/api/3/action/package_show?id=criminal-incident"
DECEMBER_FILE = re.compile(r"LGA_Criminal_Incidents_Year_Ending_December_(\d{4})\.xlsx", re.IGNORECASE)
SHEET = "Table 02"
SOURCE = "Crime Statistics Agency (Victoria), criminal incidents by LGA"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

# Sheet row label -> CSA offence subdivision codes (or a division letter).
CATEGORIES = {
    "Assault": ["A20"],
    "Robbery": ["A50"],
    "Other offences against the person": ["A10", "A30", "A40", "A60", "A70", "A80"],
    "Theft": ["B30", "B40", "B50", "B60"],
    "Malicious damage to property": ["B10", "B20"],
    "Drug offences": ["C"],
    "Transport regulatory offences": ["F20"],
    "Other offences": ["D", "E", "F10", "F30", "F90"],
}


def category_of(subdivision: str) -> str | None:
    """The sheet row for a CSA subdivision ("A20 Assault and related
    offences" -> "Assault"); None if it belongs to none."""
    code = (subdivision or "").strip().split(" ")[0].upper()
    for label, codes in CATEGORIES.items():
        if code in codes or code[:1] in codes:
            return label
    return None


def december_files(catalogue: dict) -> dict[int, str]:
    """{last year: url} of the December LGA workbooks in the catalogue
    entry (package_show result)."""
    found = {}
    for resource in catalogue.get("resources", []):
        url = resource.get("url") or ""
        match = DECEMBER_FILE.search(url)
        if match:
            found[int(match.group(1))] = url
    return found


def parse_rates(rows, lgas: set[str]) -> dict[str, dict[int, dict[str, float]]]:
    """{lga: {year: {sheet row label: rate per 1,000}}} from the rows of
    Table 02 (header first)."""
    rows = iter(rows)
    header = [str(h or "").strip() for h in next(rows)]
    col = {name: header.index(name) for name in ("Year", "Year ending", "Local Government Area",
                                                 "Offence Subdivision", "LGA Rate per 100,000 population")}
    out: dict[str, dict[int, dict[str, float]]] = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    for row in rows:
        if not row or row[col["Year"]] is None:
            continue
        lga = str(row[col["Local Government Area"]] or "").strip()
        if lga not in lgas or str(row[col["Year ending"]]).strip() != "December":
            continue
        label = category_of(str(row[col["Offence Subdivision"]] or ""))
        rate = row[col["LGA Rate per 100,000 population"]]
        if label and isinstance(rate, (int, float)):
            out[lga][int(row[col["Year"]])][label] += rate / 100
    return out


def candidates(available: dict[int, str], earliest: int, tried: set[int]) -> list[int]:
    """December workbooks that would add years before `earliest` (each
    covers its last year and the nine before), best first: the one
    reaching furthest back while still meeting the years already held."""
    return sorted(y for y in available if y not in tried and earliest - 1 <= y and y - 9 < earliest)


class CSACrimeFetcher(BaseFetcher):

    SOURCE_NAME      = "crime_csa"
    SUPPORTED_STATES = ["VIC"]

    def fetch_all(self):
        import openpyxl
        import requests

        towns = self.applicable_towns()
        if not towns:
            return
        try:
            response = requests.get(CATALOGUE, headers=HEADERS, timeout=60)
            response.raise_for_status()
            available = december_files(response.json()["result"])
        except Exception as exc:
            self.result.add_error("ALL", f"CSA catalogue not read: {exc}")
            return
        if not available:
            self.result.add_error("ALL", "no December LGA workbook listed in the CSA catalogue")
            return

        folder = CACHE_DIR / "crime" / "csa"
        folder.mkdir(parents=True, exist_ok=True)
        lgas = {re.sub(r"\s*\([^)]*\)\s*$", "", t.lga).strip() for t in towns}
        merged: dict[str, dict[int, dict[str, float]]] = defaultdict(dict)
        used = []
        def read(last_year: int) -> bool:
            url = available[last_year]
            path = folder / f"csa_lga_incidents_dec{last_year}.xlsx"
            if not path.exists() or self.force:
                try:
                    data = requests.get(url, headers=HEADERS, timeout=300)
                    data.raise_for_status()
                    if not data.content.startswith(b"PK"):
                        raise ValueError("not an .xlsx file")
                    path.write_bytes(data.content)
                except Exception as exc:
                    self.log.info(f"  CSA December {last_year} workbook not downloaded ({exc})")
                    return False
            book = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                rates = parse_rates(book[SHEET].iter_rows(values_only=True), lgas)
            finally:
                book.close()
            used.append(last_year)
            for lga, years in rates.items():
                for year, values in years.items():
                    merged[lga].setdefault(year, dict(values))     # newest workbook first: keep its figure
            self.log.info(f"  CSA December {last_year} workbook read ({last_year - 9}-{last_year})")
            return True

        latest = max(available)
        if not read(latest):
            self.result.add_error("ALL", f"latest CSA workbook ({latest}) not downloaded")
            return
        earliest, tried = latest - 9, {latest}
        while earliest > YEAR_START:
            for last_year in candidates(available, earliest, tried):
                tried.add(last_year)
                if read(last_year):
                    earliest = last_year - 9
                    break
            else:
                break

        out_dir = CACHE_DIR / "crime"
        for town in towns:
            lga = re.sub(r"\s*\([^)]*\)\s*$", "", town.lga).strip()
            years = {y: v for y, v in merged.get(lga, {}).items() if YEAR_START <= y <= YEAR_END}
            if not years:
                self.result.add_error(town.name, f"LGA '{lga}' not found in the CSA workbooks")
                self.result.towns_failed.append(town.name)
                continue
            indicators = {label: {"label": label,
                                  "values": {str(y): round(years[y].get(label, 0.0), 6) for y in sorted(years)}}
                          for label in CATEGORIES}
            out = {
                "town": town.name, "state": town.state, "lga": lga, "source": SOURCE,
                "source_url": available[max(used)] if used else "",
                "note": ("Rates per 1,000 persons, calendar years: CSA criminal incidents by LGA (Table 02), "
                         "subgroup rates per 100,000 added by category and divided by 100. Categories grouped "
                         "from the Victorian offence classification to the New South Wales rows; see "
                         "fetch_crime_csa.py."),
                "indicators": indicators,
            }
            with open(out_dir / f"{town.slug}_crime_csa.json", "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2)
            self.log.info(f"  {town.name} ({lga}): {min(years)}-{max(years)}")
            self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = CSACrimeFetcher().run()
    sys.exit(0 if result.success else 1)
