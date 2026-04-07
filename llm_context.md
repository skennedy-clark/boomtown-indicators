# boomtown-indicators: Project Context Document
# Generated: 2026-04-07 | Paste this at the start of a new conversation

## PROJECT OVERVIEW
Automating the data pipeline for UQ Centre for Natural Gas `boomtown-indicators` project.
Two outputs: CSVs feeding https://boomtown-indicators.org/compare (charting tool) and PDF booklets
at https://gas-energy.centre.uq.edu.au/indicators/download-years-town-booklets
Related site: https://boomtown-toolkit.org

**Repo**: https://github.com/skennedy-clark/boomtown-indicators.git
**Local**: `C:\Users\uqsken12\OneDrive - The University of Queensland\Desktop\boomtown-indicators\regional-indicators\`
**2025 data dir**: `C:\Users\uqsken12\OneDrive - The University of Queensland\Desktop\ccsg-social\Indicators 2025\1 Data\`
**Environment**: Windows + Docker Ubuntu (`docker run -it --rm -v ${PWD}:/workspace -w /workspace ubuntu bash`)

---

## ABSOLUTE RULES (scientific publication, not LLM convenience)
1. NEVER make up data or fill gaps with assumptions
2. Priority: automate existing QLD process — do NOT break QLD to fix NSW/VIC
3. All SA2/geography code changes require ground-truth justification from source documents, not just "it works"
4. Every data discrepancy vs. previous publications must be documented — auditable and reproducible
5. Towns with problematic data sources are PARKED with notes, not silently skipped or proxied without disclosure
6. The pipeline must work from scratch each year — no permanent manual input files assumed to exist
7. Where full automation is not possible, the pipeline must: attempt automation, report clearly what failed, and provide exact manual instructions so a researcher can complete the step

---

## PIPELINE STATUS: 10 of 25 CSVs AUTOMATED AND VALIDATED

### WORKING (do not break these)
| Fetcher | Source | Status |
|---|---|---|
| `income` | ATO Table 8, data.gov.au CKAN API | ✓ 17 towns |
| `income_table6` | ATO Table 6A/6B, data.gov.au CKAN API | ✓ 17 towns |
| `population_ucl` | QGSO UCL ERP CSV, page-scrape fallback | ✓ 11 QLD towns |
| `crime_qps` | QPS open data S3 bucket | ✓ 14 QLD towns |

### PRODUCING DATA BUT NEEDS VALIDATION
| Fetcher | Status | Issue |
|---|---|---|
| `salm_unemployment` | 14/17 towns ok | NSW/VIC towns absent from SALM dataset — PARKED |
| `bom_rainfall` | 0/14 ok | Wrong approach — see DATA SOURCES section |
| `qgso_housing` | 0 ok | Needs annual manual download — see DATA SOURCES section |

---

## DATA SOURCES: AUTOMATION STATUS FOR EACH INDICATOR

### GROUP A: FULLY AUTOMATED
**ATO Income (postcode)**
- data.gov.au CKAN API auto-discovers current Table 6 and Table 8 xlsx URLs
- `https://data.gov.au/data/api/3/action/package_show?id=taxation-statistics-2022-23`
- Lag: ~18 months

**QPS Crime (QLD only)**
- QPS open data S3 (monthly updated, no auth required)
- `https://open-crime-data.s3-ap-southeast-2.amazonaws.com/Crime%20Statistics/division_Reported_Offences_Rates.csv`
- towns.toml field: `qps_division`

**QGSO Population UCL (QLD only)**
- QGSO direct CSV, URL scraped from issue page
- Pattern: `https://www.qgso.qld.gov.au/issues/{N}/estimated-resident-population-urban-centre-locality-qld-2001-{year}p.csv`
- Equivalent xlsx in 2025 data dir: `QGSO erp ucl 2001-2024.xlsx`
- Sheet: `estimated-resident-population-u`, row 5+: UCL name | 2001 | 2002 | ...

**SALM Unemployment (QLD — national SA2 coverage)**
- DEWR direct CSV download, URL scraped from resource page
- Resource page: `https://www.dewr.gov.au/employment-research/resources/salm-smoothed-sa2-datafiles-asgs-2021`
- Fallback URL (Dec 2025): `https://www.dewr.gov.au/download/17068/salm-smoothed-sa2-datafiles-asgs-2021-december-quarter-2025/42403/salm-smoothed-sa2-datafiles-asgs-2021-december-quarter-2025/csv`
- ~2351KB CSV, 2336 SA2s, years 2010 to current, ASGS 2021 codes
- Equivalent xlsx in 2025 data dir: `SALM Smoothed SA2 Datafiles (ASGS 2021) - December quarter 2024.xlsx`
- PARKED: Narrabri, Shepparton, Yarram absent from dataset (too small for SA2 threshold)

---

### GROUP B: REQUIRES ANNUAL MANUAL DOWNLOAD — QGSO REGIONAL DATABASE
The QGSO Regional Database (`http://www.qgso.qld.gov.au/products/tables/qld-regional-database/index.php`)
requires browser interaction. The researcher manually selects datasets and downloads xlsx files.

**Pipeline approach for these sources**:
1. Attempt programmatic download (investigate network requests in browser — may have direct URLs)
2. If automation not possible: clearly instruct researcher what to download and where to save it
3. Read from `cache/qgso_and_bom_{YEAR}.xlsx` if present
4. Run summary must state: "X indicators require manual download — see instructions"

**The assembled file** the researcher creates is named `QGSO and BoM 2024.xlsx` (21KB for 2024).
It contains exactly these sheets with this format:

**Labour sheet** (SALM unemployment — NOTE: already superseded by Group A DEWR fetcher)
- Collection header: `Labour Force - Small Area`
- Series header: `Smoothed - Unemployment Rate (Per cent)`  
- Row 5 = header: `Period | SA2/307011172 - Chinchilla | SA2/307011175 - Miles - Wandoan | ...`
- Rows 6+ = quarterly data: `Qtr Ended 31 Mar 2024 | 2.8 | 3.1 | ...`
- Columns include SA2 and LGA codes in format `SA2/XXXXXXXXX - Name` and `LGA/XXXXX - Name (R)`

**Sales sheet** (house sales and median price)
- Collection header: `Residential land and dwelling sales`
- Series header: `Detached dwelling: number of sales (Number)`
- Row 5 = header: `Region | Year Ended 31 Mar 2024 | Year Ended 30 Jun 2024 | ...`
- Region format same as Labour: `SA2/307011176 - Roma`

**rent sheet** (median weekly rent 3-bedroom)
- No collection header — straight data
- Row 1 = header: `Region | Year Ended 31 Mar 2024 | ...`
- 20 rows of data covering all SA2s and LGAs for our towns

**Building approvals sheet**
- Collection: `Building Approvals`, Time period in header: `Year Ended 31 Dec 2024`
- Cols: `Region | Residential dwelling units (Private); New Houses (Number) | Non-residential...`

**Pop sheet** (SA2/LGA estimated resident population)
- Collection: `Population (ERP)(a) persons only`
- Row 5 = header: `Region | [town name] | [boundary type] | None | 2024`
- Single year snapshot only in this file

**NRW sheet** (non-resident workers)
- Format: `LGA | Location | UCL | 2024 ERP | NRW on-shift | FTE estimate`
- Covers Dysart, Moranbah (Isaac LGA) and Surat Basin towns

**Rainfall sheet** (monthly BOM data — manually entered by researcher from BOM CDO website)
- Row 1: year (e.g., 2024)
- Row 2: `Town | Station | Number | Jan | Feb | Mar | Apr | May | Jun | Jul | Aug | Sep | Oct | Nov | Dec | Annual`
- One row per town, monthly mm values, None where data missing, Annual total
- Source: BOM Climate Data Online `http://www.bom.gov.au/climate/data/`
- Station numbers: see CONFIRMED BOM STATIONS below

---

### GROUP C: QGSO NRW REPORTS (Surat Basin and Bowen Basin)
These are separate QGSO publication pages with downloadable xlsx files.

**Surat Basin** (Western Downs, Maranoa towns — Roma, Chinchilla, Dalby etc.):
- URL: `https://www.qgso.qld.gov.au/statistics/theme/population/non-resident-population-queensland-resource-regions/surat-basin`
- 2025 file: `QGSO surat basin FTE LGA UCL 2023-2024.xlsx` (13KB)

**Bowen Basin** (Isaac — Moranbah, Dysart):
- URL: `https://www.qgso.qld.gov.au/statistics/theme/population/non-resident-population-queensland-resource-regions/bowen-galilee-basins`
- 2025 compiled file: `QGSO Bowen Basin NRW 2012-2024.xlsx` (65KB) — researcher compiled across years
- 2025 latest year: `QGSO bowen basin FTE LGA UCL 2023-2024.xlsx` (14KB)

**NRW xlsx structure** (Bowen Basin 2012-2024 compiled):
- Sheet `Selected 2012-24`: cols = LGA | Location | UCL | year ERP | NRW | FTE | year ERP | NRW | FTE...
- Individual year sheets (2012-13 through 2023-24): row 1 = source URL (Wayback Machine for old years), then data
- Each sheet format: `LGA | Location(In town/Drive-in drive-out) | UCL | ERP | NRW | FTE`

**AUTOMATION**: These pages likely have direct xlsx download links — check page source. Should be automatable.

---

### GROUP D: BOM RAINFALL — REWRITE NEEDED
**CRITICAL**: The `fetch_bom_rainfall.py` FTP approach is ENTIRELY WRONG.
- The HQmonthlyR FTP dataset (`/anon/home/ncc/www/change/HQmonthlyR/`) returns 550 errors for all stations
- This is a different (historical quality-controlled) dataset, NOT what the booklets use
- The booklets use BOM Climate Data Online (CDO), current observations

**Correct approach**: BOM CDO HTTP endpoint
```
http://www.bom.gov.au/jsp/ncc/cdio/weatherData/av?p_display_type=dailyZippedDataFile&p_stn_num={STATION}&p_c=&p_nccObsCode=136&p_startYear=
```
- Requires browser session cookie (GET to bom.gov.au first, then use session for download)
- Returns a zip containing a CSV of monthly observations
- Many research pipelines use this successfully with proper headers + cookies
- The `fetch_bom_rainfall.py` must be rewritten to use CDO, not FTP

**Fallback**: If CDO download fails → read from `cache/qgso_and_bom_{YEAR}.xlsx` Rainfall sheet
(which has the single most recent year, manually entered by researcher)

**CONFIRMED BOM STATION NUMBERS** (from Indicator_database_2025.xlsx Notes sheet):
| Town | Station name | Number | Notes |
|---|---|---|---|
| Goondiwindi | Goondiwindi WTP | 41559 | Was New Kildonan (41507) for 2001-2019 |
| Chinchilla | Harewood | 42078 | 22km from town |
| Miles | Miles Post Office | 42023 | 0.6km |
| Wandoan | Gililgulgul TM | 35029 | 27.5km |
| Tara | Woodlea | 42104 | 16.8km |
| Dalby | Dalby Airport | 41240 | 2.2km |
| Wallumbilla | Yuleba Garden St | 43043 | 19.7km |
| Roma | Roma Airport | 43091 | 3.1km |
| Toowoomba | Toowoomba Airport | 41529 | 4.1km |
| Moranbah | Moranbah Airport | 34035 | 7.6km |
| Dysart | Booroondarra | 35109 | 29km — was Seloh Nolem (closed) pre-2017 |
| Narrabri | Narrabri Airport AWS | 54038 | 6km |

---

### GROUP E: ABS BUSINESS COUNTS (Cat. 8165)
- Source: ABS Counts of Australian Businesses
- URL: `https://www.abs.gov.au/statistics/economy/business-indicators/counts-australian-businesses-including-entries-and-exits`
- 2025 file: `ABS SA2 Turnover Businesses.xlsx` (8.4MB)
- Data by SA2, turnover range (0-50k, 50-100k, 100-499k, 500k-2m, 2m+), industry, year
- The Indicator_database_2025.xlsx Business sheet uses: Region, Town, Indicator, Count, Turnover Range, Industry, Year, Period, PP/NPP
- Likely automatable via ABS direct download

---

### GROUP F: NSW-SPECIFIC (Narrabri — do not build until QLD complete)
| Indicator | Source URL | 2025 file |
|---|---|---|
| NSW crime | `https://www.bocsar.nsw.gov.au/Pages/bocsar_datasets/Offence.aspx` | `BOCSAR Narrabri LGA.xlsx` |
| NSW housing | `https://dcj.nsw.gov.au/...` + ABS building approvals | `Narrabri housing 2024.xlsx`, `ABS Building Approvals NSW LGAs...xlsx` |
| NSW population | `https://www.abs.gov.au/statistics/people/population/regional-population` | `ABS SA2 ERP (Narrabri)...xlsx` |
| NSW unemployment | `https://www.abs.gov.au/.../labour-force-australia` | `ABS State NSW unemployment...xlsx` |

### GROUP G: NOT YET SOURCED
| Indicator | Notes |
|---|---|
| Student enrolments | ACARA. File: `ACARA School Profile 2008-2024.xlsx` (29MB). API available. |
| Fuel prices | Source unclear. File: `Annual Fuel Price Report 2024.pdf`. ACCC metro-focused. |
| VIC crime/housing | Vic Police + Vic housing. Not started. |

---

## CONFIRMED SA2 CODES (ASGS 2021 Edition 3)
Verified against `QGSO and BoM 2024.xlsx` sheet headers (ground truth for 2025 data).

| Town | sa2_code | ASGS 2021 name | Notes |
|---|---|---|---|
| Roma | 307011176 | Roma | Unchanged from 2016 |
| Chinchilla | 307011172 | Chinchilla | Was 307021343 in 2016 |
| Dalby | 307021183 | Wambo | Dalby absent from SALM; Wambo is correct proxy per 2022 booklet |
| Miles | 307011175 | Miles - Wandoan | Was 307021348 in 2016 |
| Tara | 307011178 | Tara | Was 307021352 in 2016 |
| Wandoan | 307011175 | Miles - Wandoan | Absorbed into Miles-Wandoan in ASGS 2021; identical SA2 data to Miles |
| Wallumbilla | 307011178 | ⚠️ VERIFY | towns.toml currently same code as Tara — may be wrong |
| Goondiwindi | 307011173 | Goondiwindi | Was 307031358 in 2016 |
| Moranbah | 312011341 | Moranbah | Was 318011429 in 2016; sa3_code = 31201 |
| Dysart | 312011338 | Broadsound - Nebo | Absorbed in ASGS 2021; confirmed in 2022 Dysart booklet p.3 |
| Toowoomba | 317011456 | Toowoomba - Central | Was 305031218 in 2016; sa3_code = 31701 |
| Toowoomba (Central) | 317011456 | Toowoomba - Central | Same SA2 as Toowoomba |
| Toowoomba (Harlaxton) | 317011454 | North Toowoomba - Harlaxton | Was 305031216 |
| Toowoomba (West) | 317011458 | Toowoomba - West | Was 305031220 |
| ⚠️ MISSING | 317011457 | Toowoomba - East | In 2025 QGSO data; not in towns.toml — needs decision |
| Narrabri | 102021116 | Narrabri | Correct but absent from SALM — PARKED |
| Shepparton | 213041536 | Shepparton | Correct but absent from SALM — PARKED |
| Yarram | 212031483 | Yarram | Correct but absent from SALM — PARKED |

---

## BOOKLET CHART INVENTORY (from 2022 Chinchilla/Dalby/Dysart PDFs)
17 charts per booklet in this order:
1. SA2 boundary map (source QGSO, ASGS code shown)
2. UCL boundary + postcode boundary maps
3. Population & Projections — UCL ERP line + SA2 ERP line + NRW UCL line + SA2 projection dashed
4. Resident vs non-resident — UCL stacked bar + LGA stacked bar
5. Non-resident projections — LGA line with Series A/B projections
6. Unemployment rate — SA2 smoothed line vs QLD benchmark red dashed
7. Average taxable income — postcode line vs QLD benchmark
8. Total wage and salary earnings — postcode dual-axis (bars=earner count, line=total $M)
9. Business count NPP by turnover — SA2 stacked bar (4 bands)
10. Business count PP by turnover — SA2 stacked bar
11. Median house sale price — SA2 bar (no. sales) + median price line vs Brisbane benchmark
12. Median weekly rent 3-bed — SA2 bar vs QLD benchmark dashed
13. New building approvals — SA2 stacked bar (residential + non-residential)
14. Traffic offences — QPS division bar vs QLD benchmark
15. Other relevant offences — QPS division lines (drugs, good order, other theft)
16. Total offences — QPS division stacked bar (person, property, other) vs QLD line
17. Rainfall — BOM station stacked bar (summer=Jan-Mar+Oct-Dec, winter=Apr-Sep) vs historic avg line

Colours: town=purple, QLD benchmark=red dashed, Brisbane benchmark=red dashed, projection=cyan dashed

---

## PARKED ISSUES
1. **Wallumbilla SA2**: towns.toml has 307011178 = same as Tara. Verify before next run.
2. **Toowoomba East** (317011457): In 2025 QGSO data but not in towns.toml. Add or not?
3. **Wandoan = Miles SA2** (307011175): Identical SA2-level data — must disclose in outputs.
4. **SALM for NSW/VIC**: Narrabri/Shepparton/Yarram absent. LGA fallback not yet researched.
5. **BOM CDO automation**: Session cookie approach not yet implemented. fetch_bom_rainfall.py needs full rewrite.
6. **QGSO housing automation**: Regional Database may require browser automation (Selenium). Needs investigation of direct download URLs first.
7. **No server access**: Cannot update boomtown-indicators.org (no FTP/CMS passwords). Step 4 blocked pending access.

---

## KEY COMMANDS
```powershell
python run_update.py                           # all fetchers
python run_update.py --only salm_unemployment  # specific fetcher
python run_update.py --force                   # ignore cache
```

## FILES TO PROVIDE IN A NEW CONVERSATION
1. Current `towns.toml`
2. Terminal run output log (paste)
3. The specific `.py` fetcher being worked on
4. `QGSO and BoM 2024.xlsx` (21KB) — if working on housing/rainfall/NRW/pop SA2
5. `Indicator_database_2025.xlsx` — if cross-validating values

## NEXT RECOMMENDED ACTIONS (in order)
1. Rewrite `fetch_bom_rainfall.py` to use BOM CDO HTTP with session cookie; fall back to xlsx; report what was automated vs manual
2. Verify Wallumbilla SA2 code in current towns.toml
3. Decide on Toowoomba East (317011457)
4. Investigate QGSO Regional Database network requests for direct download URLs (browser devtools)
5. Build housing fetchers with clear fallback/reporting when automation not possible
6. Build ABS business counts fetcher (Group D)
7. Cross-validate all automated outputs against Indicator_database_2025.xlsx