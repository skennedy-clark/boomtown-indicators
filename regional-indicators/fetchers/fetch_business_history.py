"""
regional-indicators/fetchers/fetch_business_history.py

Fetches the history of ABS business counts by SA2 and turnover size, so
that a town added to the workbook starts with its earlier years. The
annual update itself uses fetch_business.py, which reads only the latest
year.

Source: "Counts of Australian Businesses, including Entries and Exits"
(ABS 8165.0), the data cube "Businesses by Industry Division by
Statistical Area Level 2 by Turnover Size Ranges". Each release holds
three years (one table per year); RELEASES lists every release since
June 2009, newest first. Releases before 2019 are on the ABS archive
site; their addresses do not change.

Which release a year is taken from: the newest release that holds it,
since each release revises the two years before its latest. Checked
against the hand-made 2026 workbook (Roma, Chinchilla, Moranbah): 12 of
17 years match exactly; in the others the workbook holds an earlier
release's figure and the difference is about 1% of the total.

Turnover bands: older releases split the bands more finely ($50k-$100k,
$100k-$200k, $200k-$500k, $500k-$2m, ...). Bands are read from the
column headings and combined into the workbook's four (0k-50k,
50k-200k, 200k-2m, 2m+). Industry A (Agriculture, Forestry and Fishing)
is primary production; every other industry is non-primary production,
as in fetch_business.py.

SA2 codes: the regions are those of fetch_business.py (its SA2_REGIONS
plus each study town's SA2). The tables up to June 2016 use ASGS 2011
codes and later ones ASGS 2016 or 2021 codes; most regional Queensland
SA2s kept their codes, but a year whose table lacks a region's code is
left out of that region's history.

Not part of the annual update: run it when a town is added,
    python run_update.py --only business_history
The downloads (about 150 MB) are cached.

Output: cache/business/history/<region slug>_business_history.json
    {"region", "sa2_codes", "years": {"2009": {"npp": {band: n}, "pp": {...}}, ...},
     "release_by_year": {"2009": "jun2007-jun2011", ...}}
Read by transform/xlsx_update/update_business.py, which fills a new
row's earlier years from it.
"""

from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from fetchers.fetch_business import (
    ABS_DC9_BASE, PRIMARY_PRODUCTION_CODE, SA2_REGIONS, _discover_current_release, business_regions,
)

_ARCHIVE = "https://www.abs.gov.au/ausstats/subscriber.nsf/log?openagent"

# (release, url), newest first. The current release is added at run time.
RELEASES = [
    ("jul2020-jun2024", f"{ABS_DC9_BASE}/jul2020-jun2024/8165DC09.xlsx"),
    ("jul2019-jun2023", f"{ABS_DC9_BASE}/jul2019-jun2023/8165DC09_revised.xlsx"),
    ("jul2018-jun2022", f"{ABS_DC9_BASE}/jul2018-jun2022/816509.xlsx"),
    ("jul2017-jun2021", f"{ABS_DC9_BASE}/jul2017-jun2021/816509.xlsx"),
    ("jul2016-jun2020", f"{ABS_DC9_BASE}/jul2016-jun2020/816509.xls"),
    ("jul2015-jun2019", f"{ABS_DC9_BASE}/jul2015-jun2019/816509.xls"),
    ("jun2013-jun2017", f"{_ARCHIVE}&816509.xls&8165.0&Data%20Cubes&0B14A92FE1CD3585CA25823A000AEA3B"
                        f"&0&Jun%202013%20to%20Jun%202017&20.02.2018&Latest"),
    ("jun2012-jun2016", f"{_ARCHIVE}&816509.xls&8165.0&Data%20Cubes&1C44398D27622CB5CA2580CD0022B084"
                        f"&0&Jun%202012%20to%20Jun%202016&21.02.2017&Latest"),
    ("jun2011-jun2015", f"{_ARCHIVE}&816509.xls&8165.0&Data%20Cubes&B90162D7BD6B0A78CA257FD4001E5C43"
                        f"&0&Jun%202011%20to%20Jun%202015&17.06.2016&Latest"),
    ("jun2010-jun2014", f"{_ARCHIVE}&816509.xls&8165.0&Data%20Cubes&1EE27342B2EF44DBCA257E3E00127254"
                        f"&0&Jun%202010%20to%20Jun%202014&08.05.2015&Latest"),
    ("jun2009-jun2013", f"{_ARCHIVE}&816508.xls&8165.0&Data%20Cubes&1D01CA5E9E3B2C15CA257CE700169B1E"
                        f"&0&Jun%202009%20to%20Jun%202013&30.05.2014&Previous"),
    ("jun2008-jun2012", f"{_ARCHIVE}&816508.xls&8165.0&Data%20Cubes&28A1562C44DDD865CA257BA50012F4ED"
                        f"&0&Jun%202008%20to%20Jun%202012&12.07.2013&Latest"),
    ("jun2007-jun2011", f"{_ARCHIVE}&sa2_turn.xls&8165.0&Data%20Cubes&1A9F9FED8D87C07BCA2579E90017FDC1"
                        f"&0&Jun%202007%20to%20Jun%202011&24.04.2012&Previous"),
]

BANDS = ("0k-50k", "50k-200k", "200k-2m", "2m+")


def band_of(heading) -> str | None:
    """The workbook band a column heading belongs to ("$50k to less than
    $100k" -> "50k-200k"), or None for other columns (Total, labels)."""
    text = str(heading or "").strip().lower()
    if not text or "total" in text:
        return None
    if text.startswith("zero"):
        return "0k-50k"
    match = re.match(r"\$(\d+)\s*(k|m)", text)
    if not match:
        return None
    value = int(match.group(1)) * (1_000 if match.group(2) == "k" else 1_000_000)
    if value < 50_000:
        return "0k-50k"
    if value < 200_000:
        return "50k-200k"
    if value < 2_000_000:
        return "200k-2m"
    return "2m+"


def _sheets(path: Path):
    """Yield (sheet name, rows) for an .xls or .xlsx file."""
    if path.suffix.lower() == ".xls" or path.read_bytes()[:4] == b"\xd0\xcf\x11\xe0":
        import xlrd
        book = xlrd.open_workbook(str(path), on_demand=True)
        for index, name in enumerate(book.sheet_names()):
            sheet = book.sheet_by_index(index)
            yield name, [sheet.row_values(r) for r in range(sheet.nrows)]
    else:
        import openpyxl
        book = openpyxl.load_workbook(path, read_only=True)
        for name in book.sheetnames:
            yield name, [list(row) for row in book[name].iter_rows(values_only=True)]


def parse_release(path: Path, wanted: set[str]) -> dict[int, dict[str, dict]]:
    """{year: {sa2 code: {"npp": {band: n}, "pp": {band: n}}}} for the SA2
    codes in `wanted`, from every year table in one release."""
    sheets = list(_sheets(path))

    # Table number -> year, from the Contents sheet (newer releases name
    # their sheets "Table 1", "Table 2", ...).
    table_years = {}
    for name, rows in sheets:
        if name.strip().lower() != "contents":
            continue
        for row in rows:
            cells = [c for c in row if c not in (None, "")]
            if len(cells) >= 2 and isinstance(cells[0], (int, float)):
                found = re.search(r"June (\d{4})", str(cells[1]))
                if found:
                    table_years[int(cells[0])] = int(found.group(1))

    result: dict[int, dict] = {}
    for name, rows in sheets:
        header = next((i for i, row in enumerate(rows[:12])
                       if any("SA2" in str(c) for c in row) and any("50k" in str(c) for c in row)), None)
        if header is None:
            continue
        named = re.search(r"(\d{4})", name)
        table = re.search(r"Table (\d+)", name)
        year = int(named.group(1)) if named else table_years.get(int(table.group(1))) if table else None
        if year is None:
            continue
        heading = rows[header]
        sa2_col = next(j for j, c in enumerate(heading) if "SA2" in str(c))
        bands = {j: band_of(c) for j, c in enumerate(heading) if band_of(c)}
        by_sa2: dict = defaultdict(lambda: {"npp": defaultdict(float), "pp": defaultdict(float)})
        for row in rows[header + 1:]:
            if len(row) <= sa2_col:
                continue
            code = str(row[sa2_col]).strip().split(".")[0]
            if code not in wanted:
                continue
            category = "pp" if str(row[0]).strip()[:1] == PRIMARY_PRODUCTION_CODE else "npp"
            for j, band in bands.items():
                try:
                    by_sa2[code][category][band] += float(row[j] or 0)
                except (TypeError, ValueError, IndexError):
                    pass
        result[year] = {code: {cat: {b: int(round(v)) for b, v in counts.items()} for cat, counts in cats.items()}
                        for code, cats in by_sa2.items()}
    return result


class ABSBusinessHistoryFetcher(BaseFetcher):

    SOURCE_NAME      = "business_history"
    SUPPORTED_STATES = []

    def fetch_all(self):
        regions = business_regions(self.config)
        wanted = {str(code) for codes in regions.values() for code in codes}

        current = _discover_current_release(self.log)
        releases = [(current, f"{ABS_DC9_BASE}/{current}/8165DC09.xlsx")]
        releases += [(name, url) for name, url in RELEASES if name != current]

        # year -> (release, {sa2: counts}), from the newest release holding the year.
        by_year: dict[int, tuple[str, dict]] = {}
        for name, url in releases:
            path = self.download(url, f"abs_business_sa2_{name}", suffix=Path(url.split("&")[0]).suffix or ".xls")
            if path is None:
                self.result.add_warning(name, f"could not download {url}")
                continue
            try:
                tables = parse_release(path, wanted)
            except Exception as exc:
                self.result.add_warning(name, f"could not read {path.name}: {exc}")
                continue
            self.log.info(f"  {name}: years {sorted(tables)}")
            for year, table in tables.items():
                by_year.setdefault(year, (name, table))

        if not by_year:
            self.result.add_error("ALL", "no release could be read")
            return

        out_dir = Path(__file__).parent.parent / "cache" / "business" / "history"
        out_dir.mkdir(parents=True, exist_ok=True)
        for region, codes in regions.items():
            years, sources = {}, {}
            for year in sorted(by_year):
                release, table = by_year[year]
                parts = [table[str(c)] for c in codes if str(c) in table]
                if len(parts) != len(codes):
                    continue
                years[str(year)] = {
                    cat: {band: sum(p[cat].get(band, 0) for p in parts) for band in BANDS}
                    for cat in ("npp", "pp")
                }
                sources[str(year)] = release
            if not years:
                if region in SA2_REGIONS:
                    self.result.add_warning(region, "no years found")
                continue
            slug = region.lower().replace(" ", "_").replace("-", "_")
            with open(out_dir / f"{slug}_business_history.json", "w", encoding="utf-8") as f:
                json.dump({"region": region, "sa2_codes": codes, "years": years,
                           "release_by_year": sources}, f, indent=2)
            self.log.info(f"  {region}: {min(years)}-{max(years)} ({len(years)} years)")
            self.result.towns_ok.append(region)
