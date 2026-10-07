"""
regional-indicators/fetchers/fetch_fuel.py
---------------------------------------------
Fetches annual average regular unleaded petrol (RULP) prices for
Queensland locations from RACQ's Annual Fuel Price Report -- the source
for the Exogenous sheet's Fuel section ("Average RULP Price (cents)"
under Bowen, Brisbane, Dalby, Goondiwindi, Miles, Moranbah, Roma and
Toowoomba).

Built 2026-10-06. Fuel had no fetcher or writer before this.

SOURCE
RACQ publishes one "Annual Fuel Price Report <year>" PDF each January.
Its Appendix 1 has a table headed "Average RULP Prices in Queensland":
one row per location (41 in the 2025 report), twelve monthly columns
for the report year, then one annual-average column per year going
back about eleven years (2025 ... 2015). Values are cents per litre;
"nd" means no data.

One report therefore carries the new year AND the published history,
which is what lets the writer cross-check the sheet's existing figures
against the source.

CONFIRMED LIVE 2026-10-06 against the hand-built 2026 answer key: all
eight 2025 figures match exactly --
  Bowen 178.4 / Brisbane 185.2 / Dalby 171.2 / Goondiwindi 170.8 /
  Miles 174.5 / Moranbah 189.5 / Roma 169.1 / Toowoomba 174.8
-- and the plain mean of all 41 locations' 2025 averages is 178.98,
matching the answer key's "Fuel for benchmark: Qld" figure.

FINDING THE REPORT (nothing year-specific is hardcoded)
RACQ's fuel pages are scanned for links whose file name contains
"annual-fuel-price-report-<year>"; the newest year wins. Confirmed
live: both pages below return the 2025 report's link to an ordinary
request. The file name itself is not predictable (the 2025 one ends
"-v3.pdf"), which is why it is discovered, not constructed.

IF THE DOWNLOAD FAILS (site change, network block)
Download the PDF by hand from the page named in the error message and
save it, under any name containing "annual-fuel-price-report-<year>",
into  regional-indicators/cache/fuel/  then re-run. The newest such
file already in that folder is used whenever the live download is
unavailable.

PDF READING
pypdfium2 (already a project dependency). In its text output each
table row comes out as ONE line: the location name followed by its
numbers, single-space separated (confirmed against the 2025 report),
so a row is parsed as "<name> <12 monthly values> <N annual values>".
The annual years are taken from the "Dec-<yy>" header -- the first
annual column is always the report year, then descending.

WHICH LOCATIONS
All of them. This fetcher does not consult towns.toml: RACQ's list is
its own (Bowen is there; Chinchilla, Tara, Wandoan are not), and the
writer simply fills whichever town blocks exist in the sheet's Fuel
section. To track another RACQ location, add a block to the sheet.

OUTPUT: cache/fuel/racq_rulp_annual.json
  {"source": "...", "report_year": 2025, "report_url": "...",
   "pdf_file": "...", "years": [2025, 2024, ...],
   "locations": {"Dalby": {"2025": 171.2, "2024": 175.6, ...}, ...},
   "queensland_mean_of_locations": {"2025": 178.98, ...}}
Read by transform/xlsx_update/update_fuel.py.

NOT COVERED (yet): diesel (same appendix, second table -- the sheet has
no diesel rows), and the Australian Institute of Petroleum national
benchmark the 2026 answer key adds at the bottom of the Fuel section.

Usage:
    python run_update.py --only fuel
    python fetchers/fetch_fuel.py            # standalone
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from urllib.parse import urljoin

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR

try:
    import requests
    import pypdfium2 as pdfium
except ImportError:
    raise ImportError("pip install requests pypdfium2")


RACQ_PAGES = [
    "https://www.racq.com.au/car/queensland-fuel-prices",
    "https://www.racq.com.au/about-us/advocacy/reports-and-research",
]
RACQ_BASE = "https://www.racq.com.au"
REPORT_LINK_RE = re.compile(
    r'href="([^"]*annual-fuel-price-report-(\d{4})[^"]*?\.pdf)[^"]*"', re.IGNORECASE
)
LOCAL_NAME_RE = re.compile(r"annual-fuel-price-report-(\d{4})", re.IGNORECASE)

TABLE_TITLE      = "Average RULP Prices in Queensland"
NEXT_TABLE_TITLE = "Average Diesel Prices in Queensland"
MONTH_COLUMNS    = 12

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
}
TIMEOUT_S = 90

FUEL_CACHE_DIR = CACHE_DIR / "fuel"
OUTPUT_NAME    = "racq_rulp_annual.json"


def discover_report_url(log) -> tuple[str, int] | None:
    """Scan RACQ's fuel pages for the newest Annual Fuel Price Report
    link. Returns (absolute url, report year) or None."""
    found: dict[int, str] = {}
    for page in RACQ_PAGES:
        try:
            resp = requests.get(page, headers=HEADERS, timeout=TIMEOUT_S)
            resp.raise_for_status()
        except requests.RequestException as exc:
            log.warning(f"  Could not read {page}: {exc}")
            continue
        for href, year in REPORT_LINK_RE.findall(resp.text):
            # keep the first link seen for each year; drop the ?rev=&hash= query
            found.setdefault(int(year), urljoin(RACQ_BASE, href.replace("&amp;", "&")))
    if not found:
        return None
    year = max(found)
    log.info(f"  Found annual reports for {sorted(found)}; using {year}: {found[year]}")
    return found[year], year


def newest_local_report() -> tuple[Path, int] | None:
    """The newest annual report PDF already sitting in cache/fuel/."""
    best: tuple[Path, int] | None = None
    if FUEL_CACHE_DIR.exists():
        for path in FUEL_CACHE_DIR.glob("*.pdf"):
            m = LOCAL_NAME_RE.search(path.name)
            if m and (best is None or int(m.group(1)) > best[1]):
                best = (path, int(m.group(1)))
    return best


def extract_table_text(pdf_path: Path) -> str:
    """Text of the RULP table: everything from its title up to the
    diesel table's title (or the end of that page)."""
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        for index in range(len(pdf)):
            text = pdf[index].get_textpage().get_text_range()
            start = text.find(TABLE_TITLE)
            if start == -1:
                continue
            end = text.find(NEXT_TABLE_TITLE, start)
            return text[start: end if end != -1 else len(text)]
    finally:
        pdf.close()
    raise ValueError(f"No table titled '{TABLE_TITLE}' found in {pdf_path.name}")


def parse_rulp_table(text: str) -> tuple[list[int], dict[str, dict[int, float]]]:
    """Parse the table text into (annual years newest-first,
    {location: {year: cents per litre}}). 'nd' cells are left out.

    The report year comes from the "Dec-<yy>" month header. Every data
    row must have the same number of values (12 months + one per annual
    year); a row with a different count is skipped, and if NO rows parse
    or the counts disagree the whole thing raises rather than guess.
    """
    m = re.search(r"Dec-(\d{2})\b", text)
    if not m:
        raise ValueError("Could not find the 'Dec-<yy>' month header in the RULP table")
    report_year = 2000 + int(m.group(1))

    value = r"(?:\d+(?:\.\d+)?|nd)"
    row_re = re.compile(rf"^([A-Za-z][A-Za-z .'\-]*?)\s+((?:{value}\s+)+{value})\s*$")

    rows: list[tuple[str, list[str]]] = []
    for line in text.splitlines():
        match = row_re.match(line.strip())
        if not match:
            continue
        tokens = match.group(2).split()
        if len(tokens) <= MONTH_COLUMNS:
            continue
        rows.append((match.group(1).strip(), tokens))

    if not rows:
        raise ValueError("No data rows could be read from the RULP table")

    counts = sorted({len(tokens) for _, tokens in rows})
    if len(counts) != 1:
        raise ValueError(
            f"RULP table rows have inconsistent numbers of values ({counts}) -- "
            f"the report's layout may have changed."
        )
    n_years = counts[0] - MONTH_COLUMNS
    years = [report_year - offset for offset in range(n_years)]

    locations: dict[str, dict[int, float]] = {}
    for name, tokens in rows:
        annual = tokens[MONTH_COLUMNS:]
        locations[name] = {
            year: float(tok) for year, tok in zip(years, annual) if tok != "nd"
        }
    return years, locations


class RACQFuelFetcher(BaseFetcher):

    SOURCE_NAME      = "fuel"
    SUPPORTED_STATES = ["QLD"]

    def fetch_all(self):
        FUEL_CACHE_DIR.mkdir(parents=True, exist_ok=True)

        pdf_path, report_url, report_year = None, "", None
        discovered = discover_report_url(self.log)
        if discovered:
            report_url, report_year = discovered
            pdf_path = FUEL_CACHE_DIR / f"racq-annual-fuel-price-report-{report_year}.pdf"
            if self.force or not pdf_path.exists():
                try:
                    resp = requests.get(report_url, headers=HEADERS, timeout=TIMEOUT_S)
                    resp.raise_for_status()
                    if not resp.content.startswith(b"%PDF"):
                        raise ValueError("response is not a PDF")
                    pdf_path.write_bytes(resp.content)
                    self.log.info(f"  Downloaded {pdf_path.name} ({len(resp.content) // 1024} KB)")
                except (requests.RequestException, ValueError) as exc:
                    self.log.warning(f"  Could not download {report_url}: {exc}")
                    pdf_path = None
            else:
                self.log.info(f"  Using cached {pdf_path.name}")

        if pdf_path is None:
            local = newest_local_report()
            if local is None:
                self.result.add_error(
                    "ALL",
                    "Could not download RACQ's Annual Fuel Price Report and no copy is "
                    "saved locally. Open https://www.racq.com.au/car/queensland-fuel-prices "
                    "in a browser, download the latest 'Annual Fuel Price Report', save it "
                    f"into {FUEL_CACHE_DIR} keeping 'annual-fuel-price-report-<year>' in "
                    "the file name, then re-run."
                )
                return
            pdf_path, report_year = local
            self.log.warning(
                f"  Live download unavailable -- using the local copy {pdf_path.name} "
                f"({report_year}). Check a newer report hasn't been published."
            )

        try:
            years, locations = parse_rulp_table(extract_table_text(pdf_path))
        except ValueError as exc:
            self.result.add_error("ALL", f"Could not read {pdf_path.name}: {exc}")
            return

        queensland_mean = {}
        for year in years:
            values = [series[year] for series in locations.values() if year in series]
            if values:
                queensland_mean[str(year)] = round(sum(values) / len(values), 2)

        out = {
            "source":      f"RACQ Annual Fuel Price Report {years[0]}, Appendix 1: {TABLE_TITLE}",
            "report_year": years[0],
            "report_url":  report_url,
            "pdf_file":    pdf_path.name,
            "units":       "cents per litre, annual average, regular unleaded (RULP)",
            "years":       years,
            "locations":   {
                name: {str(y): v for y, v in sorted(series.items())}
                for name, series in sorted(locations.items())
            },
            "queensland_mean_of_locations": queensland_mean,
        }
        out_path = FUEL_CACHE_DIR / OUTPUT_NAME
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        self.log.info(
            f"  Parsed {len(locations)} locations, annual averages {years[-1]}-{years[0]}; "
            f"Queensland mean of locations {years[0]} = {queensland_mean.get(str(years[0]))}"
        )
        self.result.towns_ok.extend(sorted(locations))


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = RACQFuelFetcher().run()
    sys.exit(0 if result.success else 1)
