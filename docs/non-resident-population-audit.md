# Non-resident population: source audit (27 October 2026)

This audit re-checked the non-resident population series for the towns added in 2026 against the QGSO publications. It covers:

- Emerald (Central Highlands LGA)
- Mitchell (Maranoa LGA)
- Charleville (Murweh LGA)

It also compared every existing town's town-level figures with the same publications. It was prompted by an independent review of earlier statements in `historic-data-for-new-towns.md`.

Every figure was read from the publication by `regional-indicators/transform/non_resident_register.py`. That script downloads the reports, reads each table's text with pypdfium2 (`fetchers/pdf_text.py`, a project dependency), and writes:

- `docs/non-resident-population-register.csv`: one row per geography, year, measure and projection series, with the publication, table, page and status.

Re-run the script to repeat every check. The publications it reads are kept in `cache/population/nrw_publications/`.

Status values in the register:

- `verified`: read from the publication.
- `not published`: the publication says, or shows, that the figure is not published separately.
- `not located`: no public publication found. The sources searched are in the notes.
- `not applicable`: the geography is outside every QGSO study area.

## 1. Verdicts on the earlier statements

| Earlier statement | Verdict | Finding |
|---|---|---|
| Emerald non-resident workers, 2011-2018, are in the archived Bowen Basin reports: 2011 850, 2012 920, 2013 425, 2014 215, 2015 185, 2016 95, 2017 150, 2018 405 | **Confirmed** | Each value read from the FTE table of that year's report (register rows give table and page). 2011 is at July 2011; the other years are at 30 June. |
| Emerald 2019-2023 are "only available from QGSO on request" ("QGSO can supply them") | **Not established** | No QGSO statement says so. The reports are on neither QGSO site. Status is now `not located`; see the gap register for what was searched. |
| Emerald 2010 | (not claimed before) | **Not located** at town level. The 2011 report gives 2010 for the Emerald SLA only (910), which is a larger area; not used. |
| Central Highlands non-resident projections are in a QGSO PDF; series B is not published by LGA | **Confirmed** | Table 3, page 6, of the 2022 to 2026 projections report. Series A, C and D match the independent review's table exactly. The table's note (a) states that series B is not published at LGA level. |
| Mitchell and Charleville have no non-resident figures (towns outside the two basins) | **Partly incorrect** | **Mitchell is in the Surat Basin** (Maranoa LGA). It is a row of the Surat Basin FTE table in 2010 (50), 2011 (30), 2012 (20), 2013 (20) and 2017 (25). In 2008, 2014-2016, 2018, 2024 and 2025 it is grouped into "other towns" (the table's note: towns with one accommodation provider or fewer than 20 non-resident workers). **Charleville and Murweh** are outside every QGSO study area: confirmed. |
| (Review) Maranoa and Roma figures were copied into Murweh and Charleville | **Incorrect** | add_town copies a block's rows and labels and clears the figures. In the towns workbook (`towns.xlsx`, 27 October 2026), the Murweh and Charleville non-resident rows hold no figures. Checked cell by cell: Population rows 17, 73 and 150 are empty. Nothing was copied, so there is nothing to remove from the data. The empty rows remain (see change log). |

## 2. Verified values

### Emerald UCL: non-resident workers on-shift ("in town")

| Year | Value | Publication; table; page |
|---|---:|---|
| 2011 (July) | 850 | Bowen and Galilee Basins population report 2011; Table 4; p. 14 |
| 2012 | 920 | Bowen Basin population report 2012; Table 3; p. 4 |
| 2013 | 425 | 2013 report; Table 2; p. 4 |
| 2014 | 215 | 2014 report; Table 2; p. 5 |
| 2015 | 185 | 2015 report; Table 2; p. 5 |
| 2016 | 95 | 2016 report; Table 2; p. 5 |
| 2017 | 150 | 2017 report; Table 2; p. 5 |
| 2018 | 405 | 2018 report; Table 2; p. 5 |
| 2024 | 635 | 2024 report; Table 2; p. 5 (resident 15,195, FTE 15,830, as the review's anchor). Also in the 2024-2025 FTE table (resident since revised to 15,180) |
| 2025 | 555 | 2025 report; Table 3; p. 5. Also in the 2024-2025 FTE table |

Definitions:

- **"In town"** counts the UCL plus non-resident workers housed within 5 km of town. This definition is stated in the 2012, 2018, 2024 and 2025 reports.
- **The 2011 report** is by LGA, SLA and UC/L, at July. The tables say "UC/L" up to 2014 and "UCL" from 2015. Whether the town boundaries changed between them was not checked.
- **The measure** is non-resident workers on-shift. It is not the FIFO/DIDO workforce, the FTE population or the resident population. The register keeps the resident and FTE figures from the same row in its notes, and does not use them.

### Mitchell UCL

| Year | Value | Status |
|---|---:|---|
| 2008 | | Not published: grouped into "other towns" |
| 2009 | | Not located: no 2009 Surat report on the archive site |
| 2010 | 50 | Verified: "Mitchell (L)"; Surat Basin population report 2010, Table 1, p. 7 |
| 2011 (July) | 30 | Verified: Table 5, p. 16 |
| 2012 | 20 | Verified |
| 2013 | 20 | Verified |
| 2014-2016, 2018 | | Not published: grouped into "other towns" |
| 2017 | 25 | Verified: Table 2, p. 5 |
| 2019-2023 | | Not located |
| 2024, 2025 | | Not published: grouped into "other towns" |

### Central Highlands LGA: non-resident population projections, 2022 to 2026

Source: *Bowen and Galilee Basins non-resident population projections, 2022 to 2026* (QGSO, June 2022). Table 3, page 6. Measure: non-resident workers on-shift at 30 June.

| Series | 2021 (estimate) | 2022 | 2023 | 2024 | 2025 | 2026 |
|---|---:|---:|---:|---:|---:|---:|
| A | 4,100 | 4,200 | 4,230 | 4,230 | 4,120 | 4,120 |
| B | not published at LGA level | | | | | |
| C | 4,100 | 4,200 | 4,430 | 4,460 | 4,260 | 4,400 |
| D | 4,100 | 4,200 | 4,430 | 4,460 | 4,360 | 4,820 |

- **2021** is the estimate the projections start from. **2022-2026** are projections of a historical vintage, not observations.
- **QGSO's page** (checked 27 October 2026) says: "Data in the Bowen and Galilee Basins non-resident population projections, 2022 to 2026 report are representative of information available at the time of collection in 2021. They are no longer current and should not be used."
- **No replacement projection release** is listed on the Bowen and Galilee Basins page or the Surat Basin page. The observed LGA series, 2006-2025, is published, and the pipeline already writes it.

## 3. Gap register

| Geography | Years | Category | Cause and what was searched |
|---|---|---|---|
| Emerald, Mitchell and every other basin town (UCL) | 2019-2023 | Not located | **Searched:** the QGSO Bowen and Galilee Basins and Surat Basin pages (current releases: the 2024 and 2025 reports and the 2024-2025 FTE tables only); the current report URL pattern for each year (404); the QGSO archive site (reports to 2018); data.qld.gov.au (its entries link back to QGSO); web search. **Not checked:** the Internet Archive, which could not be reached from the checking environment. The 2026 workbook holds hand-entered 2019-2023 figures for Moranbah, Dysart, Roma and other towns, so the reports existed and the research team may hold copies. |
| Emerald UCL | 2010 | Not located | Only the SLA figure (910) is published for 2010. |
| Mitchell UCL | 2008, 2014-2016, 2018, 2024, 2025 | Not published | Grouped into "other towns" in the table (the table's note gives the rule). |
| Mitchell UCL | 2009 | Not located | No 2009 Surat Basin report on the archive site. The 2026 workbook has 2009 figures for Chinchilla, Dalby and Miles, so a 2009 publication existed. |
| Central Highlands LGA projections | Series B | Explicitly not published | Table 3, note (a). |
| Murweh LGA, Charleville UCL | all | Not applicable | Outside every QGSO non-resident study area. Bowen Basin = Banana, Central Highlands, Isaac, Whitsunday (Bowen only). Galilee = Barcaldine. Surat = Maranoa, Western Downs, Toowoomba (Surat Basin population report 2025). North West region = Burke, Cloncurry, McKinlay, Mount Isa (a resource-industry workforce report, not a non-resident population estimate). No other official source found. |

Nothing in this register is "available on request": no QGSO statement to that effect was found.

## 4. What the comparison with the 2026 workbook found

The register's town figures were compared with the hand-entered rows of the 2026 workbook (`Indicators_Data-Charts_2026.xlsx`). Of the 97 figures both hold, 96 agree. The others need checking by the research team; nothing was changed:

1. **Toowoomba 2018.** The workbook has 205; the 2018 Surat Basin report (Table 2, p. 5) has 165.
2. **Tara and Wallumbilla, years the report does not publish.**
   - Tara: 2011, 2015-2018 (5, 10, 5, 0, 0).
   - Wallumbilla: 2013-2018 (18, 10, 5, 0, 0, 30).

   In those years the report groups these towns into "other towns". The workbook's figures may come from another table or another estimate; their source is not recorded.
3. **NRW projection rows sat under the wrong years.** On the Population sheet, the existing towns' "Population projection: NRW 2022e" rows (Isaac, Maranoa, Western Downs, Toowoomba) held the 2022, 2023, 2024, 2025 and 2026 projections in the columns headed 2019, 2021, 2023, 2024 and 2026. The 2021 estimate is not shown. Every value matches the report; only the columns were wrong. **Fixed 28 October 2026** (see the change log).
4. **2018-edition population projections sat under the wrong years.** This is a similar layout issue: the 2016, 2021 and 2026 projections were in the columns headed 2011, 2017 and 2025. **Fixed 28 October 2026.**

## 5. Change log

| Item | Before | After | Reason | Files and outputs |
|---|---|---|---|---|
| Emerald UCL non-resident workers 2019-2023 | Documented as "QGSO can supply them" | `not located`, with the search recorded | No evidence of request-only availability | `docs/historic-data-for-new-towns.md` |
| Mitchell UCL non-resident workers | Documented as "none" | 2010, 2011, 2012, 2013, 2017 verified; other years not published or not located | Mitchell is in the Surat Basin study area | `historic-data-for-new-towns.md`; the writer now fills the row |
| Emerald and Mitchell rows in the workbook | Empty after add_town; the writer wrote only the latest year (Emerald 2025), with no history | `update_population_nrw.py` fills a new town's empty row with the register's verified years only. Years not published or not located stay blank (null, not zero) | Automate the verified history | `transform/xlsx_update/update_population_nrw.py`; Population sheet; website CSV "Non-resident workers in town" |
| Existing towns' town figures | Checked by the shape-based outlier test, which flagged Dysart, Moranbah, Roma and Toowoomba for 2012-2023 points | Compared with the publications instead (`audit_historical_series`) | The flags were on figures that match the reports | Same writer; fewer false flags |
| Murweh and Charleville non-resident rows | Empty rows with labels copied by add_town | Unchanged: no figures were ever written. The research team can delete the empty rows. Their website lines (`Charleville`, "Non-resident workers in local govt area" and "... in town") will produce empty files | No data exists for these areas; no proxy was used | `towns.xlsx`; `web_export_map.csv` lines 344-345 |
| Projection rows on the Population sheet: QGSO 2018 edition (LGA high/medium/low, SA2 medium) and NRW 2022e series A-D | Figures under other years' columns (section 4, items 3 and 4) | Each figure moved to its own year's column by `update_population_projections.py`. On the 2025 starting workbook this is 129 cells in 40 rows; each move is listed in the run output | Agreed with the research team, 28 October 2026. A row is moved only when its figures match the source year by year in column order; nothing else is moved | Population sheet. The charts reading these rows (Benchmarks NRW charts; town SA2 projection charts) span the years concerned and need no change. No website file reads them |
| Central Highlands NRW projections | Not in the workbook | Series A, C and D written for 2022-2026 under their own years. Series B left empty (not published); the 2021 estimate is not shown, as for the other LGAs | Layout fixed | Population sheet |

## 6. Questions for QGSO

1. Are the Bowen Basin and Surat Basin population reports and FTE tables for 2019 to 2023 still available? Their town-level figures, for Emerald in particular, are not on the current or archived QGSO sites.
2. Was a Surat Basin population report published for 2009? The archive site has 2008 and 2010-2018.
3. Have the town figures in the 2024-2025 FTE tables (resident population revised from the 2024 report) replaced the 2024 report's figures for all towns?
4. Is a newer non-resident population projection release planned to replace the withdrawn 2022 to 2026 report?
5. Does QGSO publish, or hold, non-resident population estimates for Murweh LGA or Charleville?

## Sources

- QGSO, Bowen and Galilee Basins page: https://www.qgso.qld.gov.au/statistics/theme/population/non-resident-population-queensland-resource-regions/bowen-galilee-basins
- QGSO, Surat Basin page: https://www.qgso.qld.gov.au/statistics/theme/population/non-resident-population-queensland-resource-regions/surat-basin
- QGSO, North West region page: https://www.qgso.qld.gov.au/statistics/theme/population/non-resident-population-queensland-resource-regions/north-west-region
- QGSO archive, Bowen Basin reports 2011-2018: https://archive.qgso.qld.gov.au/products/reports/bowen-basin-pop-report/index.html
- QGSO archive, Surat Basin reports 2008, 2010-2018: https://archive.qgso.qld.gov.au/products/reports/surat-basin-pop-report/index.html
- Bowen Basin population report 2024: https://www.qgso.qld.gov.au/issues/3366/bowen-basin-population-report-2024.pdf (and `-2025.pdf`)
- Bowen Basin FTE table 2024-2025: https://www.qgso.qld.gov.au/issues/3341/bowen-basin-population-report-tables-full-time-equivalent-fte-population-estimates-local-government-area-lga-selected-urban-centres-localities-ucls-2024-2025.xlsx
- Surat Basin population reports 2024 and 2025: https://www.qgso.qld.gov.au/issues/3186/surat-basin-population-report-2025.pdf
- Projections 2022 to 2026: https://www.qgso.qld.gov.au/issues/3376/bowen-galilee-basins-non-resident-population-projections-2022-2026.pdf
- QGSO resource communities research: https://www.qgso.qld.gov.au/about-qgso/our-services/resource-communities-research
