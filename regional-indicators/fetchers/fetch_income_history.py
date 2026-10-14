"""
regional-indicators/fetchers/fetch_income_history.py

Fetches the history of the postcode income figures on the Income sheet,
so that a town added to the workbook starts with its earlier years. The
annual update itself uses fetch_income_table6.py (latest year) and
fetch_income.py (Table 8).

Source: ATO Taxation Statistics, "Individuals Table 6" by postcode, one
file per income year from 2000-01, gathered in one data.gov.au package:
    https://data.gov.au/data/dataset/taxation-statistics-postcode-data
The file list is read from the package (CKAN package_show).

Figures per income year, for the town's primary postcode:
    avg_income_taxable      taxable income or loss ($) / number, taxable individuals
    earners_no, wages_total salary or wages, number and $
    abn_npp_income_*,       total business income (non-primary production,
    abn_pp_income_*,        primary production, all), number and $
    abn_total_income_*
Up to 2009-10 the ATO published taxable individuals only, and business
income only from 2010-11; those are the figures for those years.
From 2010-11, all individuals: the table without a taxable-status split
where there is one (Table 6B, or the "postcode only" table), otherwise
the taxable and non-taxable rows added together. 2010-11 has no taxable
status split, so it has no avg_income_taxable.

The ATO changed the table layout most years, so columns are found by
their headings ("Salary or wages", "Total business income - primary
production", ...), with "no." and "$" taken from the heading or the row
below it.

Checked against the hand-made 2026 workbook for Roma (4455):
avg_income_taxable 23 of 23 years; business income 12-14 of 14; wages
20 of 24. The other wage years in the workbook (2011-12 to 2015-16) hold
taxable individuals only, where this fetcher gives all individuals.

Not part of the annual update: run it when a town is added,
    python run_update.py --only income_history
The downloads (about 125 MB) are cached.

Output: cache/ato/<town slug>_income_history.json
    {"town", "postcode", "years": {"2001": {"avg_income_taxable": ..., ...}, ...}}
Years are the calendar year in which the income year ends (2000-01 -> 2001).
Read by transform/xlsx_update/update_income.py, which fills a new row's
earlier years from it.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher

PACKAGE_URL = "https://data.gov.au/data/api/3/action/package_show?id=taxation-statistics-postcode-data"
FIRST_INCOME_YEAR = 2000          # 2000-01, the first year on the Income sheet

ALL_STATUS = "All"


def _measure(label: str) -> str | None:
    text = label.lower()
    if text.startswith("postcode"):
        return "postcode"
    if text.startswith("taxable status"):
        return "status"
    if text.startswith("taxable individuals"):
        return "taxable_individuals"
    if text.startswith(("mean taxable income", "average taxable income")):
        return "taxable_mean"
    if text.startswith("taxable income or loss"):
        return "taxable"
    if text.startswith("salary or wages"):
        return "wages"
    if re.match(r"total business income\s*[-–]?\s*primary production", text):
        return "bus_pp"
    if re.match(r"total business income\s*[-–]?\s*non[- ]primary production", text):
        return "bus_npp"
    if re.match(r"total business income\d*(\s|$)", text):
        return "bus_total"
    return None


def column_map(heading: list, units: list | None) -> dict[str, int]:
    """Measure -> column, from a table's heading row and the row below it.
    A heading spanning a "no." and a "$" column is carried into the
    unlabelled second column."""
    heading = list(heading)
    if units is not None:
        for j in range(1, len(heading)):
            if heading[j] in (None, "") and j < len(units) and units[j] not in (None, "") and heading[j - 1]:
                heading[j] = heading[j - 1]
    found: dict[str, int] = {}
    for j, cell in enumerate(heading):
        if cell in (None, ""):
            continue
        label = re.sub(r"\s+", " ", str(cell)).strip()
        measure = _measure(label)
        if measure is None:
            continue
        if measure in ("postcode", "status", "taxable_individuals", "taxable_mean"):
            found.setdefault(measure, j)
            continue
        unit_text = label.lower()
        if units is not None and j < len(units) and units[j] not in (None, ""):
            unit_text += " " + str(units[j]).lower()
        unit = "$" if "$" in unit_text else ("no" if re.search(r"\bno\b", unit_text) else None)
        if unit:
            found.setdefault(f"{measure}_{unit}", j)
    return found


def _number(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def postcode_rows(path: Path, postcodes: set[str]) -> list[dict]:
    """Every row for the given postcodes, as {"sheet", "postcode", "status", measures...}."""
    import openpyxl
    book = openpyxl.load_workbook(path, read_only=True)
    found = []
    try:
        for name in book.sheetnames:
            rows = [list(r) for r in book[name].iter_rows(values_only=True)]
            header = next((i for i, r in enumerate(rows[:10])
                           if r and any(isinstance(c, str) and c.strip().lower().startswith("postcode") for c in r)), None)
            if header is None:
                continue
            cols = column_map(rows[header], rows[header + 1] if header + 1 < len(rows) else None)
            if "postcode" not in cols:
                continue
            for row in rows[header + 1:]:
                if len(row) <= cols["postcode"]:
                    continue
                postcode = str(row[cols["postcode"]]).split(".")[0].strip()
                if postcode not in postcodes:
                    continue
                entry = {"sheet": name, "postcode": postcode,
                         "status": str(row[cols["status"]]).strip() if "status" in cols else ALL_STATUS}
                for key, j in cols.items():
                    if key not in ("postcode", "status") and j < len(row):
                        entry[key] = _number(row[j])
                found.append(entry)
    finally:
        book.close()
    return found


def year_figures(rows: list[dict]) -> dict:
    """The Income-sheet figures for one postcode and one income year."""
    out: dict = {}
    taxable = next((r for r in rows if r["status"].lower() == "taxable"), None)
    if taxable and taxable.get("taxable_no"):
        out["avg_income_taxable"] = taxable["taxable_$"] / taxable["taxable_no"]

    whole = [r for r in rows if r["status"] == ALL_STATUS and r.get("wages_no") is not None]
    split = [r for r in rows if r["status"] != ALL_STATUS and r.get("wages_no") is not None]
    if whole:
        base = max(whole, key=lambda r: sum(v is not None for v in r.values()))
    elif split:
        base = {}
        for r in split:
            for key, value in r.items():
                if isinstance(value, float):
                    base[key] = base.get(key, 0.0) + value
    else:
        return out
    if "avg_income_taxable" not in out:
        # Before 2009-10: taxable individuals only, with a mean column or a count.
        if base.get("taxable_mean"):
            out["avg_income_taxable"] = base["taxable_mean"]
        elif base.get("taxable_individuals") and base.get("taxable_$"):
            out["avg_income_taxable"] = base["taxable_$"] / base["taxable_individuals"]
    for name, key in (("earners_no", "wages_no"), ("wages_total", "wages_$"),
                      ("abn_npp_income_no", "bus_npp_no"), ("abn_npp_income_total", "bus_npp_$"),
                      ("abn_pp_income_no", "bus_pp_no"), ("abn_pp_income_total", "bus_pp_$"),
                      ("abn_total_income_no", "bus_total_no"), ("abn_total_income_total", "bus_total_$")):
        if base.get(key) is not None:
            out[name] = base[key]
    return {k: round(v, 2) if k == "avg_income_taxable" else round(v) for k, v in out.items()}


class ATOIncomeHistoryFetcher(BaseFetcher):

    SOURCE_NAME      = "income_history"
    SUPPORTED_STATES = []

    def fetch_all(self):
        import requests
        towns = [t for t in self.config.study_towns() if t.postcode]
        postcodes = {str(t.postcode) for t in towns}

        try:
            resources = requests.get(PACKAGE_URL, timeout=60).json()["result"]["resources"]
        except Exception as exc:
            self.result.add_error("ALL", f"could not read the data.gov.au package: {exc}")
            return

        by_year: dict[int, list[dict]] = {}
        for resource in resources:
            found = re.match(r"(\d{4})-(\d{2}) income year", resource.get("name", "").strip())
            if not found or int(found.group(1)) < FIRST_INCOME_YEAR:
                continue
            end_year = int(found.group(1)) + 1
            url = resource["url"]
            path = self.download(url, f"ato_postcode_table6_{end_year}", suffix=Path(url).suffix or ".xlsx")
            if path is None:
                self.result.add_warning(str(end_year), f"could not download {url}")
                continue
            try:
                by_year[end_year] = postcode_rows(path, postcodes)
            except Exception as exc:
                self.result.add_warning(str(end_year), f"could not read {path.name}: {exc}")

        out_dir = Path(__file__).parent.parent / "cache" / "ato"
        out_dir.mkdir(parents=True, exist_ok=True)
        for town in towns:
            years = {}
            for year in sorted(by_year):
                figures = year_figures([r for r in by_year[year] if r["postcode"] == str(town.postcode)])
                if figures:
                    years[str(year)] = figures
            if not years:
                self.result.add_warning(town.name, f"postcode {town.postcode} not found")
                continue
            with open(out_dir / f"{town.slug}_income_history.json", "w", encoding="utf-8") as f:
                json.dump({"town": town.name, "postcode": str(town.postcode), "years": years}, f, indent=2)
            self.log.info(f"  {town.name} ({town.postcode}): {min(years)}-{max(years)}")
            self.result.towns_ok.append(town.name)
