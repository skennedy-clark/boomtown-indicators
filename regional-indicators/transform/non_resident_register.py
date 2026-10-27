"""
regional-indicators/transform/non_resident_register.py

Builds the source register for the non-resident population series of the
towns added in 2026 (Emerald, Central Highlands; Mitchell, Charleville,
Murweh): every value with the publication, table and page it was read
from, and every gap with what was searched. The register is rebuilt
from the QGSO publications themselves, so each figure can be checked
again by running this script.

Output: docs/non-resident-population-register.csv, one row per
geography, year, measure and projection series, with the columns
    geography, geography_level, reference_year, measure, value, unit,
    data_type, projection_series, projection_vintage, source_url,
    source_table_page, verification_status, notes

verification_status is one of
    verified              read from the publication by this script
    not located           no public publication found (see notes)
    not published         the publication says the figure is not published
    not applicable        the geography is outside every QGSO study area

Sources read
  Bowen Basin population reports 2011-2018 (QGSO archive site, as at
  September 2019): table "FTE population estimates ... selected UC/Ls",
  the UCL's row: resident population, non-resident workers on-shift,
  FTE population. 2011 is at July 2011; the other years at 30 June.
  Bowen Basin population reports 2024 and 2025 (current QGSO site), and
  the 2024-2025 FTE spreadsheet.
  Bowen and Galilee Basins non-resident population projections, 2022 to
  2026 (QGSO, June 2022), Table 3 (Central Highlands). QGSO's page says
  of this report: "They are no longer current and should not be used."

  Surat Basin population reports 2008 and 2010-2018 (archive site) and
  2024-2025 (current site), for Mitchell (Maranoa LGA): the same table.
  In years where a town is not a row of the table, the table's note says
  why (towns with one accommodation provider or fewer than 20
  non-resident workers are grouped as "other towns").

What was searched for the years not found (Emerald 2010, both towns
2019-2023), and for Murweh and Charleville, is recorded in the notes.

Usage (from the repository root; needs network access the first time,
the publications are then kept in cache/population/nrw_publications/):
    python regional-indicators/transform/non_resident_register.py
"""

from __future__ import annotations

import csv
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent
sys.path.insert(0, str(HERE.parent))

CACHE = HERE.parent / "cache" / "population" / "nrw_publications"
OUTPUT = REPO / "docs" / "non-resident-population-register.csv"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

ARCHIVE = "https://archive.qgso.qld.gov.au/products/reports/bowen-basin-pop-report/"
CURRENT = "https://www.qgso.qld.gov.au/issues/"

# (reference year, publication, url). The 2011 report is the combined
# Bowen and Galilee Basins report; its estimates are at July 2011.
SURAT_ARCHIVE = "https://archive.qgso.qld.gov.au/products/reports/surat-basin-pop-report/"

BOWEN_REPORTS = [
    (2011, "Bowen and Galilee Basins population report, 2011", ARCHIVE + "bowen-galilee-basins-pop-report-2011.pdf"),
    *[(year, f"Bowen Basin population report, {year}", ARCHIVE + f"bowen-basin-pop-report-{year}.pdf")
      for year in range(2012, 2019)],
    (2024, "Bowen Basin population report, 2024", CURRENT + "3366/bowen-basin-population-report-2024.pdf"),
    (2025, "Bowen Basin population report, 2025", CURRENT + "3366/bowen-basin-population-report-2025.pdf"),
]
# Surat Basin: reports for 2008 and 2010-2018 on the archive site (none for
# 2009), 2024 and 2025 on the current site. 2011 is at July 2011.
SURAT_REPORTS = [
    *[(year, f"Surat Basin population report, {year}", SURAT_ARCHIVE + f"surat-basin-pop-report-{year}.pdf")
      for year in (2008, *range(2010, 2019))],
    (2024, "Surat Basin population report, 2024", CURRENT + "3186/surat-basin-population-report-2024.pdf"),
    (2025, "Surat Basin population report, 2025", CURRENT + "3186/surat-basin-population-report-2025.pdf"),
]
FTE_TABLE_2024_2025 = (CURRENT + "3341/bowen-basin-population-report-tables-full-time-equivalent-fte-population-"
                       "estimates-local-government-area-lga-selected-urban-centres-localities-ucls-2024-2025.xlsx")
PROJECTIONS = (CURRENT + "3376/bowen-galilee-basins-non-resident-population-projections-2022-2026.pdf")
PROJECTION_VINTAGE = "Bowen and Galilee Basins non-resident population projections, 2022 to 2026 (QGSO, June 2022)"
PROJECTION_WARNING = ("QGSO: 'Data in the Bowen and Galilee Basins non-resident population projections, 2022 to "
                      "2026 report are representative of information available at the time of collection in 2021. "
                      "They are no longer current and should not be used.'")
QGSO_PAGE = ("https://www.qgso.qld.gov.au/statistics/theme/population/"
             "non-resident-population-queensland-resource-regions/bowen-galilee-basins")

MEASURE_NRW = "non-resident workers on-shift"

# The LGAs of each QGSO study area (current QGSO pages and reports).
BASIN_LGAS = {"Bowen": {"Banana", "Central Highlands", "Isaac", "Whitsunday"},
              "Surat": {"Maranoa", "Western Downs", "Toowoomba"}}
COLUMNS = ["geography", "geography_level", "reference_year", "measure", "value", "unit", "data_type",
           "projection_series", "projection_vintage", "source_url", "source_table_page",
           "verification_status", "notes"]

IN_TOWN_NOTE = ("'In town' counts the UCL plus non-resident workers housed within 5 km of town "
                "(definition stated in the 2012, 2018, 2024 and 2025 reports)")


def download(url: str) -> Path:
    import requests
    CACHE.mkdir(parents=True, exist_ok=True)
    path = CACHE / url.rsplit("/", 1)[-1]
    if not path.exists():
        response = requests.get(url, headers=HEADERS, timeout=180)
        response.raise_for_status()
        path.write_bytes(response.content)
    return path


def _number(text: str) -> int:
    return int(text.replace(",", ""))


FTE_TITLE = re.compile(r"Table \d+:?\s+(?:Full.time equivalent(?: \(FTE\))?|FTE) population (?:estimates )?(?:for )?")


def ucl_row(pdf_path: Path, ucl: str) -> dict | None:
    """The UCL's row of the report's FTE table: {"page", "title",
    "resident", "nrw", "fte"}; or, when the town is not in the table,
    {"page", "title", "aggregation"} with the table's note on towns
    grouped into "other towns" (None if the report has no FTE table)."""
    from fetchers.pdf_text import page_lines
    row = re.compile(rf"(?:^|\s){re.escape(ucl)}(?:\s\(L\))?\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)(?:\s+[\d.]+)?\s*$")
    first_table = None
    for number, lines in enumerate(page_lines(pdf_path), start=1):
        title = next((line for line in lines if FTE_TITLE.match(line) and ".." not in line), None)  # not the contents page
        if not title:
            continue
        for line in lines:
            match = row.search(line)
            if match:
                resident, nrw, fte = (_number(g) for g in match.groups())
                return {"page": number, "title": title, "resident": resident, "nrw": nrw, "fte": fte}
        if first_table is None:
            note = next((line for line in lines if "aggregated" in line or "included in the LGA remainder" in line
                         or "fewer than" in line), "")
            first_table = {"page": number, "title": title, "aggregation": note}
    return first_table


def projection_table(pdf_path: Path, lga: str) -> dict | None:
    """{"page", "title", "years", "series": {name: [values]}, "note"} for
    the LGA's table in the projections report (parsed as the projections
    fetcher parses it)."""
    from fetchers.fetch_population_projections import parse_nrw_tables
    from fetchers.pdf_text import page_lines
    table = parse_nrw_tables(page_lines(pdf_path)).get(lga)
    if not table:
        return None
    years = sorted(int(y) for y in next(iter(table["series"].values())))
    return {"page": table["page"], "title": table["table"], "years": years,
            "series": {name: [values[str(y)] for y in years] for name, values in table["series"].items()},
            "note": table["not_published"]}


def build() -> list[dict]:
    rows: list[dict] = []

    def add(**fields):
        rows.append({column: fields.get(column, "") for column in COLUMNS})

    def town_series(ucl: str, reports: list, region: str) -> dict:
        found = {}
        for year, publication, url in reports:
            try:
                result = ucl_row(download(url), ucl)
            except Exception as exc:
                result, error = None, str(exc)
            else:
                error = ""
            when = "July 2011" if year == 2011 else f"30 June {year}"
            if result and "nrw" in result:
                found[year] = result["nrw"]
                add(geography=ucl, geography_level="UCL", reference_year=year, measure=MEASURE_NRW,
                    value=result["nrw"], unit="persons", data_type="observed (QGSO estimate)", source_url=url,
                    source_table_page=f"{publication}; {result['title']}; page {result['page']}",
                    verification_status="verified",
                    notes=f"At {when}. Same row: resident population {result['resident']:,}, FTE population "
                          f"{result['fte']:,}. " + IN_TOWN_NOTE
                          + (" The 2011 Bowen table is by LGA, SLA and UC/L." if year == 2011 and region.startswith("Bowen") else ""))
            elif result:
                add(geography=ucl, geography_level="UCL", reference_year=year, measure=MEASURE_NRW,
                    data_type="observed (QGSO estimate)", source_url=url,
                    source_table_page=f"{publication}; {result['title']}; page {result['page']}",
                    verification_status="not published",
                    notes=f"{ucl} is not a row of the table: it is grouped with other towns. Table note: "
                          f"{result['aggregation'] or '(none found)'}")
            else:
                add(geography=ucl, geography_level="UCL", reference_year=year, measure=MEASURE_NRW,
                    data_type="observed (QGSO estimate)", source_url=url, verification_status="not located",
                    notes=f"No FTE table by town found in this report. {error}".strip())
        for year in range(2019, 2024):
            add(geography=ucl, geography_level="UCL", reference_year=year, measure=MEASURE_NRW,
                verification_status="not located", data_type="observed (QGSO estimate)",
                notes="Workbook rows of towns that were in the study before 2024 hold hand-entered figures for "
                      "these years, from the reports while they were published. Searched: QGSO {region} page (current releases are the 2024 and 2025 reports and the "
                      "2024-2025 FTE table only); the current report URL pattern for that year (404); QGSO "
                      "archive site (reports to 2018); data.qld.gov.au (links back to QGSO); web search. "
                      "Internet Archive not reachable from the checking environment. No QGSO statement found "
                      "that the reports are available on request.")
        return found

    # Every Queensland study town in a basin LGA, by its UCL name.
    from config import get_config
    regions = {"Bowen": (BOWEN_REPORTS, "Bowen and Galilee Basins"), "Surat": (SURAT_REPORTS, "Surat Basin")}
    found = {}
    for town in get_config().study_towns():
        lga = re.sub(r"\s*\([^)]*\)\s*$", "", town.lga or "").strip()
        region = next((r for r, lgas in BASIN_LGAS.items() if lga in lgas), None)
        if town.state != "QLD" or not region or "(" in town.name:
            continue
        reports, page = regions[region]
        found[town.name] = town_series(town.name, reports, page)
        if town.name == "Emerald":
            add(geography="Emerald", geography_level="UCL", reference_year=2010, measure=MEASURE_NRW,
                verification_status="not located", data_type="observed (QGSO estimate)",
                source_url=ARCHIVE + "bowen-galilee-basins-pop-report-2011.pdf",
                notes="The 2011 report (Table 2, page 12) gives 2010 for the Emerald SLA only (910), a larger "
                      "area than the UCL; not used. The Bowen Basin UCL series starts with the 2011 report.")
        if region == "Surat":
            add(geography=town.name, geography_level="UCL", reference_year=2009, measure=MEASURE_NRW,
                verification_status="not located", data_type="observed (QGSO estimate)",
                notes="No 2009 Surat Basin population report on the QGSO archive site (reports for 2008 and "
                      "2010-2018).")
    found = found.get("Emerald", {})

    # The 2024-2025 FTE spreadsheet, as a second source for 2024 and 2025.
    try:
        import openpyxl
        book = openpyxl.load_workbook(download(FTE_TABLE_2024_2025), read_only=True, data_only=True)
        for sheet in book.worksheets:
            for values in sheet.iter_rows(values_only=True):
                cells = [v for v in values if v is not None]
                if cells and cells[0] == "Emerald" and len(cells) >= 7:
                    for year, nrw in ((2024, cells[2]), (2025, cells[5])):
                        agrees = "agrees with" if found.get(year) == nrw else "DIFFERS FROM"
                        add(geography="Emerald", geography_level="UCL", reference_year=year, measure=MEASURE_NRW,
                            value=nrw, unit="persons", data_type="observed (QGSO estimate)",
                            source_url=FTE_TABLE_2024_2025,
                            source_table_page=f"Bowen Basin FTE population estimates 2024 and 2025 (xlsx), sheet {sheet.title}",
                            verification_status="verified",
                            notes=f"Second source; {agrees} the report. The written workbook value comes from this table.")
    except Exception as exc:                       # the register is still useful without it
        print(f"FTE spreadsheet not read: {exc}")

    # Central Highlands LGA: projections (historical vintage).
    table = projection_table(download(PROJECTIONS), "Central Highlands")
    if table:
        for name, values in table["series"].items():
            for year, value in zip(table["years"], values):
                estimated = year == table["years"][0]
                add(geography="Central Highlands", geography_level="LGA", reference_year=year,
                    measure=MEASURE_NRW, value=value, unit="persons",
                    data_type="estimated base (QGSO estimate)" if estimated else "projection (historical vintage)",
                    projection_series=name, projection_vintage=PROJECTION_VINTAGE, source_url=PROJECTIONS,
                    source_table_page=f"{table['title']}; page {table['page']}", verification_status="verified",
                    notes=("At 30 June. " + PROJECTION_WARNING) if not estimated else
                          "At 30 June 2021; the estimate the projections start from, equal in every series.")
        if "Series B" not in table["series"]:
            for year in table["years"]:
                add(geography="Central Highlands", geography_level="LGA", reference_year=year, measure=MEASURE_NRW,
                    data_type="projection (historical vintage)", projection_series="Series B",
                    projection_vintage=PROJECTION_VINTAGE, source_url=PROJECTIONS,
                    source_table_page=f"{table['title']}; page {table['page']}", verification_status="not published",
                    notes=table["note"] or "Series B not published at LGA level.")

    # Geographies outside the QGSO study areas.
    areas = ("QGSO non-resident population study areas (checked on the current QGSO pages): Bowen Basin = Banana, "
             "Central Highlands, Isaac, Whitsunday (Bowen only); Galilee Basin = Barcaldine; Surat Basin = Maranoa, "
             "Western Downs, Toowoomba (Surat Basin population report 2025); North West region resource "
             "industry workforce = Burke, Cloncurry, McKinlay, Mount Isa (workforce counts, not non-resident "
             "population).")
    for geography, level in (("Murweh", "LGA"), ("Charleville", "UCL")):
        add(geography=geography, geography_level=level, reference_year="all", measure=MEASURE_NRW,
            verification_status="not applicable",
            notes=f"Outside every QGSO non-resident study area. {areas} No other official source found. The "
                  "workbook rows (structure copied from Maranoa/Roma by add_town) hold no figures; nothing was copied.")
    return rows


def main() -> int:
    rows = build()
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(rows)
    verified = sum(1 for r in rows if r["verification_status"] == "verified")
    print(f"{OUTPUT}: {len(rows)} rows, {verified} verified")
    return 0


if __name__ == "__main__":
    sys.exit(main())
