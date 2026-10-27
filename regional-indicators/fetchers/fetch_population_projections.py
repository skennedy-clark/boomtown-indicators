"""
regional-indicators/fetchers/fetch_population_projections.py

Fetches the population projections shown on the Population sheet, for
every area the sources cover (not only the study towns), so that a town
added later needs no new download.

Queensland: QGSO projections, 2018 edition (the edition the workbook
holds; Isaac and Moranbah match it exactly). Only QGSO's archive site
still carries it:
    LGA, low, medium and high series, 2016 to 2041 in five-year steps:
      https://archive.qgso.qld.gov.au/products/tables/proj-pop-lga-qld/proj-pop-lga-qld-2018-edn.xlsx
      (sheets proj-pop-lga-medium-qld, -low-qld, -high-qld; LGA names
      with their type, "Isaac (R)")
    SA2, medium series, same years:
      https://archive.qgso.qld.gov.au/products/tables/proj-pop-medium-series-sa2-sa3-sa4-qld/proj-pop-medium-series-sa2-sa3-sa4-qld-2018-ed.xlsx
      (sheet proj-pop-sa2-qld; SA2 code and name in columns E and F)
  The 2016 column is the base estimate, equal in all three series.

New South Wales: NSW Department of Planning "Common Planning Assumption"
projections, 2024 edition (the edition of Narrabri's hand-entered
figures: 2031, 2036 and 2041 match it exactly, LGA and SA2), annual to
2041, LGA and SA2. The files are linked from
    https://www.planning.nsw.gov.au/data-and-resources/population-projections/explore-the-data
as .../<edition>-nsw-population-projections-local-government-areas.xlsx
and ...-sa2s.xlsx; sheet "Total population", a header row "Local
Government Area" (or "SA2") followed by years, one area per row. The
title row says which years are projected ("Historic (2001-2021) and
Projected (2022-2041)"); only projected years are kept.

Queensland resource regions: QGSO non-resident population projections,
2022 to 2026 (published June 2022), one table per LGA ("Table 3
Projected non-resident population, Central Highlands"): the June 2021
estimate and series A to D for 2022-2026, in non-resident workers
on-shift. A series QGSO does not publish for an LGA is absent from its
table (the table's note says so). QGSO's page says these projections
"are no longer current and should not be used"; they are what the
sheet's "Population projection: NRW 2022e" rows hold.
      https://www.qgso.qld.gov.au/issues/3376/bowen-galilee-basins-non-resident-population-projections-2022-2026.pdf
      https://www.qgso.qld.gov.au/issues/3191/surat-basin-non-resident-population-projections-2022-2026.pdf

The source files are kept in cache/population/projections/ and are
downloaded only when missing (or with force). A projection edition is
fixed, so the files never need replacing; a new edition is taken only
by deleting the cached file, and a new edition would differ from the
figures already in the workbook.

Output: cache/population/population_projections.json
    {"qld_2018": {"source", "years",
                  "lga": {name: {"low"|"medium"|"high": {year: value}}},
                  "sa2": {name: {"code", "medium": {year: value}}}},
     "nsw": {"source", "edition", "first_projected_year",
             "lga": {name: {year: value}}, "sa2": {name: {year: value}}},
     "nrw_2022": {"source", "warning", "lga": {name: {"estimate_year",
             "series": {"Series A": {year: value}}, "not_published",
             "table", "page", "url"}}}}
Names are as the sources write them; update_population_projections.py
compares them loosely ("Isaac (R)" = "Isaac", "Broadsound - Nebo" =
"Broadsound-Nebo").
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR

QLD_LGA_URL = ("https://archive.qgso.qld.gov.au/products/tables/proj-pop-lga-qld/"
               "proj-pop-lga-qld-2018-edn.xlsx")
QLD_SA2_URL = ("https://archive.qgso.qld.gov.au/products/tables/proj-pop-medium-series-sa2-sa3-sa4-qld/"
               "proj-pop-medium-series-sa2-sa3-sa4-qld-2018-ed.xlsx")
NSW_PAGE = "https://www.planning.nsw.gov.au/data-and-resources/population-projections/explore-the-data"
NSW_SITE = "https://www.planning.nsw.gov.au"
NSW_FILE = re.compile(r'href="([^"]*?(\d{4})-nsw-population-projections-(local-government-areas|sa2s)\.xlsx)"')

NRW_REPORTS = [
    "https://www.qgso.qld.gov.au/issues/3376/bowen-galilee-basins-non-resident-population-projections-2022-2026.pdf",
    "https://www.qgso.qld.gov.au/issues/3191/surat-basin-non-resident-population-projections-2022-2026.pdf",
]
NRW_SOURCE = "QGSO non-resident population projections, 2022 to 2026 (June 2022)"
NRW_WARNING = ("QGSO: the 2022 to 2026 projections are representative of information available in 2021; "
               "'They are no longer current and should not be used.'")

SOURCE_DIR = CACHE_DIR / "population" / "projections"
OUTPUT = CACHE_DIR / "population" / "population_projections.json"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


# ── Parsing ───────────────────────────────────────────────────────────────────

def _year(value) -> int | None:
    """2021, or "2016 (b)" -> 2016."""
    if isinstance(value, (int, float)) and not isinstance(value, bool) and float(value).is_integer():
        return int(value) if 1990 <= value <= 2100 else None
    match = re.match(r"\s*(\d{4})\b", str(value or ""))
    return int(match.group(1)) if match and 1990 <= int(match.group(1)) <= 2100 else None


def _number(value) -> float | None:
    return float(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else None


def _year_header(rows) -> tuple[int, dict[int, int]] | None:
    """The first row holding a run of at least three years: (row index,
    {column: year})."""
    for i, row in enumerate(rows):
        years = {c: _year(v) for c, v in enumerate(row) if _year(v)}
        values = list(years.values())
        if len(values) >= 3 and all(b > a for a, b in zip(values, values[1:])):
            return i, years
    return None


def parse_qld_lga(rows) -> dict[str, dict[str, float]]:
    """{LGA name: {year: persons}} from one series sheet of the QGSO LGA
    table. The name is in the first column; the years are in a header
    row below the title."""
    header = _year_header(rows)
    if not header:
        return {}
    start, years = header
    out = {}
    for row in rows[start + 1:]:
        name = row[0] if row else None
        if not isinstance(name, str) or not name.strip():
            continue
        series = {str(y): _number(row[c]) for c, y in years.items() if c < len(row) and _number(row[c]) is not None}
        if len(series) == len(years):
            out[name.strip()] = series
    return out


def parse_qld_sa2(rows) -> dict[str, dict]:
    """{SA2 name: {"code", "medium": {year: persons}}} from the QGSO SA2
    table. The SA2 code is the nine-digit number in the row; the name
    follows it."""
    header = _year_header(rows)
    if not header:
        return {}
    start, years = header
    out = {}
    for row in rows[start + 1:]:
        for c, value in enumerate(row):
            if isinstance(value, (int, float)) and len(str(int(value))) == 9 and c + 1 < len(row) \
                    and isinstance(row[c + 1], str):
                series = {str(y): _number(row[k]) for k, y in years.items()
                          if k < len(row) and _number(row[k]) is not None}
                if len(series) == len(years):
                    out[row[c + 1].strip()] = {"code": str(int(value)), "medium": series}
                break
    return out


def parse_nsw(rows) -> tuple[int | None, dict[str, dict[str, float]]]:
    """(first projected year, {area name: {year: persons}}) from the
    "Total population" sheet of an NSW projections file. Only projected
    years are kept."""
    first = None
    for row in rows[:10]:
        for value in row:
            match = re.search(r"Projected\s*\((\d{4})\s*-\s*\d{4}\)", str(value or ""))
            if match:
                first = int(match.group(1))
    header = _year_header(rows)
    if not header:
        return first, {}
    start, years = header
    out = {}
    for row in rows[start + 1:]:
        name = row[0] if row else None
        if not isinstance(name, str) or not name.strip():
            continue
        series = {str(y): _number(row[c]) for c, y in years.items()
                  if (first is None or y >= first) and c < len(row) and _number(row[c]) is not None}
        if series:
            out[name.strip()] = series
    return first, out


NRW_TABLE = re.compile(r"^Table \d+ Projected non-resident population, (.+?)\s*$")


def parse_nrw_tables(pages: list[list[str]]) -> dict[str, dict]:
    """{LGA: {"estimate_year", "series": {"Series A": {year: value}},
    "not_published", "table", "page"}} from the text of a QGSO
    non-resident population projections report (fetchers/pdf_text.py).
    The first year of each table is the estimate the projections start
    from. Region totals ("Bowen Basin", "Surat Basin") are kept as well."""
    out: dict[str, dict] = {}
    for number, lines in enumerate(pages, start=1):
        for i, line in enumerate(lines):
            match = NRW_TABLE.match(line)
            if not match:
                continue
            years, series, note = [], {}, ""
            for following in lines[i + 1:i + 14]:
                if NRW_TABLE.match(following) or following.startswith("Source:"):
                    break
                if not years and re.fullmatch(r"(20\d\d\s+){2,}20\d\d", following):
                    years = [int(y) for y in following.split()]
                row = re.match(r"^(Series [A-D])\s+([\d,\s]+)$", following)
                if row and years:
                    values = [int(v.replace(",", "")) for v in row.group(2).split()]
                    if len(values) == len(years):
                        series[row.group(1)] = {str(y): v for y, v in zip(years, values)}
                if following.startswith("(a)"):
                    note = following
            if years and series:
                out[match.group(1).strip()] = {"estimate_year": years[0], "series": series,
                                               "not_published": note, "table": line, "page": number}
    return out


def nsw_links(html: str) -> dict[str, tuple[str, int]]:
    """{"local-government-areas"|"sa2s": (url, edition)} from the NSW
    page, the newest edition of each."""
    found: dict[str, tuple[str, int]] = {}
    for href, edition, kind in NSW_FILE.findall(html):
        url = href if href.startswith("http") else NSW_SITE + href
        if kind not in found or int(edition) > found[kind][1]:
            found[kind] = (url, int(edition))
    return found


# ── Fetcher ───────────────────────────────────────────────────────────────────

class PopulationProjectionsFetcher(BaseFetcher):

    SOURCE_NAME      = "population_projections"
    SUPPORTED_STATES = ["QLD", "NSW"]

    def fetch_all(self):
        import openpyxl
        import requests

        SOURCE_DIR.mkdir(parents=True, exist_ok=True)

        def get(url: str, path: Path) -> Path | None:
            if path.exists() and not self.force:
                return path
            try:
                response = requests.get(url, headers=HEADERS, timeout=180)
                response.raise_for_status()
                if not response.content.startswith(b"PK"):
                    raise ValueError("not an .xlsx file")
                path.write_bytes(response.content)
                self.log.info(f"  downloaded {path.name} ({len(response.content):,} bytes)")
                return path
            except Exception as exc:
                self.result.add_error("ALL", f"{url}: {exc}")
                return None

        def sheet_rows(path: Path, name: str) -> list[tuple]:
            book = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                return [tuple(r) for r in book[name].iter_rows(values_only=True)]
            finally:
                book.close()

        out: dict = {}
        lga_file = get(QLD_LGA_URL, SOURCE_DIR / "qgso-proj-pop-lga-2018.xlsx")
        sa2_file = get(QLD_SA2_URL, SOURCE_DIR / "qgso-proj-pop-sa2-2018.xlsx")
        if lga_file and sa2_file:
            lga: dict[str, dict] = {}
            for series in ("low", "medium", "high"):
                for name, values in parse_qld_lga(sheet_rows(lga_file, f"proj-pop-lga-{series}-qld")).items():
                    lga.setdefault(name, {})[series] = values
            sa2 = parse_qld_sa2(sheet_rows(sa2_file, "proj-pop-sa2-qld"))
            out["qld_2018"] = {"source": "QGSO population projections, 2018 edition (QGSO archive site)",
                               "lga": lga, "sa2": sa2}
            self.log.info(f"  Queensland 2018 edition: {len(lga)} LGAs, {len(sa2)} SA2s")

        try:
            response = requests.get(NSW_PAGE, headers=HEADERS, timeout=120)
            response.raise_for_status()
            links = nsw_links(response.text)
        except Exception as exc:
            links = {}
            self.log.warning(f"  NSW projections page not read ({exc}); using cached files")
        nsw: dict = {}
        for kind, key in (("local-government-areas", "lga"), ("sa2s", "sa2")):
            cached = sorted(SOURCE_DIR.glob(f"nsw-*-{kind}.xlsx"))
            if cached and not self.force:
                path = cached[-1]
            elif kind in links:
                url, edition = links[kind]
                path = get(url, SOURCE_DIR / f"nsw-{edition}-{kind}.xlsx")
            else:
                self.result.add_error("NSW", f"no NSW projections file found for {kind}")
                continue
            if not path:
                continue
            edition = int(re.search(r"nsw-(\d{4})-", path.name).group(1))
            first, areas = parse_nsw(sheet_rows(path, "Total population"))
            nsw.update({key: areas, "edition": edition, "first_projected_year": first})
            self.log.info(f"  NSW {edition} edition, {key}: {len(areas)} areas, projected from {first}")
        if nsw:
            nsw["source"] = f"NSW Common Planning Assumption population projections, {nsw['edition']} edition"
            out["nsw"] = nsw

        nrw: dict[str, dict] = {}
        for url in NRW_REPORTS:
            path = SOURCE_DIR / url.rsplit("/", 1)[-1]
            if not path.exists() or self.force:
                try:
                    response = requests.get(url, headers=HEADERS, timeout=180)
                    response.raise_for_status()
                    if not response.content.startswith(b"%PDF"):
                        raise ValueError("not a PDF")
                    path.write_bytes(response.content)
                except Exception as exc:
                    self.result.add_error("QLD", f"{url}: {exc}")
                    continue
            from fetchers.pdf_text import page_lines
            for lga, table in parse_nrw_tables(page_lines(path)).items():
                nrw[lga] = dict(table, url=url)
        if nrw:
            out["nrw_2022"] = {"source": NRW_SOURCE, "warning": NRW_WARNING, "lga": nrw}
            self.log.info(f"  Non-resident projections 2022 to 2026: {', '.join(sorted(nrw))}")

        if not out:
            return
        OUTPUT.parent.mkdir(parents=True, exist_ok=True)
        with open(OUTPUT, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=1)

        # Report which study towns' areas the sources cover.
        for town in self.applicable_towns():
            key = "qld_2018" if town.state == "QLD" else "nsw"
            data = out.get(key, {})
            lga = loose_name(re.sub(r"\s*\([^)]*\)\s*$", "", town.lga))
            have_lga = any(loose_name(name) == lga for name in data.get("lga", {}))
            code = str(getattr(town, "qgso_sa2", "") or town.sa2_code or "")
            have_sa2 = any(loose_name(name) == loose_name(town.sa2_name) or area.get("code") == code
                           for name, area in data.get("sa2", {}).items() if isinstance(area, dict))
            if have_lga or have_sa2:
                self.result.towns_ok.append(town.name)
            else:
                self.result.add_warning(town.name, "neither its LGA nor its SA2 is in the projections")
            if not (have_lga and have_sa2):
                missing = "LGA" if not have_lga else "SA2"
                self.log.info(f"  {town.name}: no {missing} projection ({town.lga if missing == 'LGA' else town.sa2_name})")


def loose_name(name) -> str:
    """Area names compared loosely: case, trailing suffixes such as the
    LGA type and state ("Central Highlands (R) (Qld)") and the spacing
    around hyphens ("Broadsound - Nebo") are ignored."""
    text = re.sub(r"(\s*\((?:[A-Za-z]{1,3}|qld|nsw)\))+\s*$", "", str(name or ""), flags=re.IGNORECASE)
    text = re.sub(r"\s*-\s*", "-", text.casefold())
    return re.sub(r"\s+", " ", text).strip()


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = PopulationProjectionsFetcher().run()
    sys.exit(0 if result.success else 1)
