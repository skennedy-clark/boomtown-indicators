"""
regional-indicators/fetchers/fetch_fuel_history.py

Fetches the history of RACQ's annual average regular unleaded petrol
(RULP) prices from every Annual Fuel Price Report RACQ still publishes
(2010 onwards), so that a town added to the workbook starts with its
earlier years. The annual update itself uses fetch_fuel.py, which reads
only the latest report (about eleven years of history).

Why the older reports: RACQ's list of locations changes. Charleville,
for example, was monitored until August 2017; it is in the reports up
to 2021 (annual figures 2009-2016) and absent from the 2022 report on,
so the latest report cannot fill it. Mitchell is in no report.

Source: the "Annual Fuel Price Report <year>" PDFs linked from RACQ's
reports page (RACQ_PAGES in fetch_fuel.py). File names are not
predictable ("racq-annual-fuel-price-report---2015.pdf",
"annual-fuel-price-report-2025-v3.pdf"), so they are discovered from
the page. Each report has a regional table, titled over the years
"Average ULP Prices for Queensland Regional Centres", "Average ULP
Prices in Queensland" or "Average RULP Prices in Queensland": one row
per location, twelve monthly columns for the report year, then one
annual-average column per year, newest first. Values are cents per
litre; "nd" means no data.

Which report a year is taken from: the newest report that holds it.
Checked on the 2011-2025 reports: every report agrees with every newer
one, to the decimal place, wherever both hold a figure, and the Roma
figures equal the workbook's in every year 2012-2018.

Safeguard: a report whose figures disagree with the newer reports for
more than a tenth of the town-years they share is set aside and named
in the run report, since that means its layout was misread (the 2010
report has a thirteen-month layout and is set aside this way).

Downloads are kept in cache/fuel/ under the same names fetch_fuel.py
uses (racq-annual-fuel-price-report-<year>.pdf), so they are copied
into the "1 Data/Fuel" folder with the year's other files, and later
runs download only reports not already there. If RACQ's site cannot be
reached, the reports already in cache/fuel/ are used; copying the PDFs
from an earlier year's "1 Data/Fuel" folder into cache/fuel/ is enough.

Not part of the annual update: run it when a town is added,
    python run_update.py --only fuel_history

Output: cache/fuel/racq_rulp_history.json
    {"source", "reports": [2011, ...], "set_aside": {"2010": "reason"},
     "locations": {"Charleville": {"2009": 131.3, ...}, ...},
     "from_report": {"Charleville": {"2009": 2021, ...}, ...}}
Read by transform/xlsx_update/update_fuel.py, which fills a new row's
earlier years from it.
"""

from __future__ import annotations

import json
import re
import sys
from collections import Counter
from pathlib import Path
from urllib.parse import urljoin

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from fetchers.fetch_fuel import FUEL_CACHE_DIR, HEADERS, RACQ_BASE, RACQ_PAGES, TIMEOUT_S

REPORT_LINK_RE = re.compile(
    r'href="([^"]*annual-fuel-price-report-+(\d{4})[^"]*?\.pdf)[^"]*"', re.IGNORECASE
)
LOCAL_NAME_RE = re.compile(r"annual-fuel-price-report-+(\d{4})", re.IGNORECASE)

TABLE_TITLE_RE = re.compile(
    r"Average R?ULP Prices (?:for (?:all )?Queensland Regional Centres|in Queensland)", re.IGNORECASE
)
NEXT_TABLE_RE = re.compile(r"Average Diesel Prices", re.IGNORECASE)
MONTH_COLUMNS = 12
TABLE_PAGES = 3                 # pages a table may run over
AGREE_WITHIN = 0.15             # cents per litre
MAX_DISAGREEMENT = 0.1          # share of shared town-years

OUTPUT_NAME = "racq_rulp_history.json"

_VALUE = r"(?:\d+(?:\.\d+)?|nd)"
_ROW_RE = re.compile(rf"^([A-Za-z][A-Za-z .'\-]*?)\s+((?:{_VALUE}\s+)+{_VALUE})\s*$")


def report_links(html: str) -> dict[int, str]:
    """{report year: absolute PDF url} for every annual report linked from a page."""
    found: dict[int, str] = {}
    for href, year in REPORT_LINK_RE.findall(html):
        found.setdefault(int(year), urljoin(RACQ_BASE, href.replace("&amp;", "&")))
    return found


def local_reports(folder: Path) -> dict[int, Path]:
    """{report year: path} for the annual report PDFs already in `folder`."""
    found: dict[int, Path] = {}
    if folder.exists():
        for path in sorted(folder.glob("*.pdf")):
            m = LOCAL_NAME_RE.search(path.name)
            if m:
                found.setdefault(int(m.group(1)), path)
    return found


def table_text(pages: list[str]) -> str | None:
    """The text of the regional RULP table, given each page's text: from
    its title (a title followed by the month header, not a mention of the
    table elsewhere) to the diesel table's title, across up to
    TABLE_PAGES pages. None if the report has no such table."""
    for index, page in enumerate(pages):
        for title in TABLE_TITLE_RE.finditer(page):
            if "Jan" not in page[title.end(): title.end() + 300]:
                continue
            text = page[title.start():]
            last = index
            while not NEXT_TABLE_RE.search(text) and last + 1 < min(len(pages), index + TABLE_PAGES):
                last += 1
                text += "\n" + pages[last]
            end = NEXT_TABLE_RE.search(text)
            return text[: end.start()] if end else text
    return None


def parse_table(text: str, report_year: int) -> tuple[dict[str, dict[int, float]], int]:
    """({location: {year: cents per litre}}, rows set aside).

    A row is "<name> <12 monthly values> <annual values, newest first>";
    the first annual value is the report year's. The number of values
    most rows have is taken as the table's width; a row of another width
    is set aside. 'nd' cells are left out."""
    rows = []
    for line in text.splitlines():
        match = _ROW_RE.match(line.strip())
        if match:
            tokens = match.group(2).split()
            if len(tokens) > MONTH_COLUMNS:
                rows.append((match.group(1).strip(), tokens))
    if not rows:
        return {}, 0
    width = Counter(len(tokens) for _, tokens in rows).most_common(1)[0][0]
    years = [report_year - k for k in range(width - MONTH_COLUMNS)]
    locations, set_aside = {}, 0
    for name, tokens in rows:
        if len(tokens) != width:
            set_aside += 1
            continue
        locations[name] = {y: float(v) for y, v in zip(years, tokens[MONTH_COLUMNS:]) if v != "nd"}
    return locations, set_aside


def canonical_names(reports: dict[int, dict[str, dict[int, float]]]) -> None:
    """Give each location one name across reports, in place: "Mt " is
    written "Mount ", and a name that is the tail of exactly one longer
    name in any report (the text of some reports splits "Charters Towers"
    over two lines, leaving "Towers") takes the longer name."""
    def tidy(name: str) -> str:
        return re.sub(r"^Mt\.? ", "Mount ", name)

    names = {tidy(n) for table in reports.values() for n in table}
    for year, table in reports.items():
        renamed = {}
        for name, series in table.items():
            name = tidy(name)
            longer = [n for n in names if n != name and n.endswith(" " + name)]
            renamed[longer[0] if len(longer) == 1 else name] = series
        reports[year] = renamed


def merge_reports(reports: dict[int, dict[str, dict[int, float]]]):
    """Combine reports, newest first: each town-year comes from the newest
    report that holds it. A report that disagrees with the newer ones on
    more than MAX_DISAGREEMENT of the town-years they share is set aside.

    Returns (locations, from_report, set_aside), where set_aside is
    {report year: reason}."""
    locations: dict[str, dict[int, float]] = {}
    from_report: dict[str, dict[int, int]] = {}
    set_aside: dict[int, str] = {}
    for report_year in sorted(reports, reverse=True):
        table = reports[report_year]
        shared = differing = 0
        for name, series in table.items():
            for year, value in series.items():
                if year in locations.get(name, {}):
                    shared += 1
                    differing += abs(locations[name][year] - value) > AGREE_WITHIN
        if shared and differing / shared > MAX_DISAGREEMENT:
            set_aside[report_year] = (f"{differing} of {shared} figures differ from newer reports "
                                      f"(layout not read correctly)")
            continue
        for name, series in table.items():
            for year, value in series.items():
                if year not in locations.get(name, {}):
                    locations.setdefault(name, {})[year] = value
                    from_report.setdefault(name, {})[year] = report_year
    return locations, from_report, set_aside


def _page_texts(pdf_path: Path) -> list[str]:
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        return [pdf[i].get_textpage().get_text_range() for i in range(len(pdf))]
    finally:
        pdf.close()


class RACQFuelHistoryFetcher(BaseFetcher):

    SOURCE_NAME      = "fuel_history"
    SUPPORTED_STATES = ["QLD"]

    def fetch_all(self):
        import requests

        FUEL_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        links: dict[int, str] = {}
        for page in RACQ_PAGES:
            try:
                response = requests.get(page, headers=HEADERS, timeout=TIMEOUT_S)
                response.raise_for_status()
                links.update({y: u for y, u in report_links(response.text).items() if y not in links})
            except requests.RequestException as exc:
                self.log.warning(f"  Could not read {page}: {exc}")

        have = local_reports(FUEL_CACHE_DIR)
        for year, url in sorted(links.items()):
            if year in have and not self.force:
                continue
            path = FUEL_CACHE_DIR / f"racq-annual-fuel-price-report-{year}.pdf"
            try:
                response = requests.get(url, headers=HEADERS, timeout=TIMEOUT_S)
                response.raise_for_status()
                if not response.content.startswith(b"%PDF"):
                    raise ValueError("response is not a PDF")
                path.write_bytes(response.content)
                self.cache.register(f"fuel_history_{year}", path, url=url)
                self.result.new_files.append(path)
                have[year] = path
                self.log.info(f"  Downloaded the {year} report ({len(response.content) // 1024} KB)")
            except (requests.RequestException, ValueError) as exc:
                self.result.add_warning(str(year), f"could not download {url}: {exc}")
        if not links:
            self.result.add_warning("ALL", "RACQ's reports page could not be read; using the reports "
                                           f"already in {FUEL_CACHE_DIR}")
        if not have:
            self.result.add_error("ALL", f"no Annual Fuel Price Reports downloaded or in {FUEL_CACHE_DIR}")
            return

        reports = {}
        for year, path in sorted(have.items()):
            try:
                text = table_text(_page_texts(path))
            except Exception as exc:
                self.result.add_warning(str(year), f"could not read {path.name}: {exc}")
                continue
            if text is None:
                self.result.add_warning(str(year), f"no regional RULP table found in {path.name}")
                continue
            table, rows_set_aside = parse_table(text, year)
            if not table:
                self.result.add_warning(str(year), f"no rows read from {path.name}")
                continue
            reports[year] = table
            note = f", {rows_set_aside} rows of another width set aside" if rows_set_aside else ""
            self.log.info(f"  {year} report: {len(table)} locations{note}")

        canonical_names(reports)
        locations, from_report, set_aside = merge_reports(reports)
        for year, reason in sorted(set_aside.items()):
            self.result.add_warning(f"{year} report", f"set aside: {reason}")
        if not locations:
            self.result.add_error("ALL", "no figures read from any report")
            return

        used = sorted(y for y in reports if y not in set_aside)
        out = {
            "source": f"RACQ Annual Fuel Price Reports {used[0]}-{used[-1]}, regional RULP tables",
            "units": "cents per litre, annual average, regular unleaded (RULP)",
            "reports": used,
            "set_aside": {str(y): r for y, r in sorted(set_aside.items())},
            "locations": {name: {str(y): v for y, v in sorted(series.items())}
                          for name, series in sorted(locations.items())},
            "from_report": {name: {str(y): r for y, r in sorted(series.items())}
                            for name, series in sorted(from_report.items())},
        }
        with open(FUEL_CACHE_DIR / OUTPUT_NAME, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)
        first = min(min(s) for s in locations.values())
        last = max(max(s) for s in locations.values())
        self.log.info(f"  {len(locations)} locations, {first}-{last}, from {len(used)} reports")
        self.result.towns_ok.extend(sorted(locations))


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = RACQFuelHistoryFetcher().run()
    sys.exit(0 if result.success else 1)
