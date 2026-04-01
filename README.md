# boomtown-indicators

Data pipeline for the UQ Centre for Natural Gas regional indicators project.
Fetches, processes and outputs social and economic indicator data for Queensland,
NSW and Victorian towns affected by CSG/resource development.

---

## Setup

```bash
git clone https://github.com/skennedy-clark/boomtown-indicators.git
cd boomtown-indicators
python -m venv venv
venv\Scripts\activate        # Windows
pip install -r requirements.txt
```

---

## Commands

```bash
# Validate towns.toml config
python run_update.py --validate

# Run all fetchers
python run_update.py

# Run one fetcher
python run_update.py --only income

# Run one fetcher for specific towns
python run_update.py --only income --towns Roma Dalby

# Clear cache and re-download everything
python run_update.py --force

# Clear cache for one fetcher only
python run_update.py --only income --force

# Show what's in the cache
python run_update.py --list-cache

# Dry run — show what would run without fetching
python run_update.py --dry-run

# Skip a fetcher
python run_update.py --skip crime

# Linux command line via Docker
docker run -it --rm -v ${PWD}:/workspace -w /workspace ubuntu bash

# View this document in preview
# markdown preview: ctrl + shift + v
```

---

## Websites

- https://gas-energy.centre.uq.edu.au/resources/tools
- https://boomtown-indicators.org/compare

---

## Fetchers

| Name | Source | Method | States | Status |
|------|--------|--------|--------|--------|
| `income` | ATO Taxation Statistics Table 8 (data.gov.au) | CKAN API | All | ✓ Working |
| `income_table6` | ATO Taxation Statistics Table 6 (data.gov.au) | CKAN API | All | TODO |
| `population_ucl` | QGSO UCL Resident Population | Direct CSV download | QLD | ✓ Working |
| `population_erp` | ABS Estimated Resident Population (Narrabri) | ABS API | NSW | TODO |
| `population_nrw` | QGSO Surat/Bowen Basin Population Reports | QGSO web | QLD | TODO |
| `unemployment` | ABS Small Area Labour Markets (SALM) | data.gov.au | All | TODO |
| `housing` | QGSO Regional Database (QRIS) | QRIS web form | QLD | TODO |
| `housing_nsw` | NSW Valuer General / SCA rent tables | ABS/SCA web | NSW | TODO |
| `crime_qps` | QPS Reported Offences Rates | QPS web download | QLD | TODO |
| `crime_nsw` | BOCSAR LGA offences | BOCSAR web | NSW | TODO |
| `rainfall` | Bureau of Meteorology (station data) | BOM API | All | TODO |
| `business` | ABS Counts of Businesses by Turnover (8165) | ABS API | All | TODO |
| `fuel` | RACQ Annual Average Fuel Prices | RACQ web | QLD | TODO |
| `schools` | ACARA School Profile | ACARA data portal | All | TODO |

---

## Towns

Configured in `towns.toml`. Currently tracking:

**QLD:** Roma, Chinchilla, Dalby, Miles, Tara, Wandoan, Wallumbilla,
Goondiwindi, Moranbah, Dysart, Toowoomba (+ 3 sub-areas)

**NSW:** Narrabri

**VIC:** Shepparton, Yarram *(test towns)*

Add new towns by adding a `[[towns]]` block to `towns.toml`.

---

## Output

Processed data lands in `cache/{source}/` as JSON.
Final CSVs (one per indicator per town) will go in `output/{town}/`.

---

## TODO

### Fetchers
- [ ] `fetch_income_table6.py` — ATO Table 6 (taxable individuals series + wage/salary totals)
  - Same CKAN API pattern as Table 8, slug: `taxation-statistics-2022-23`
  - Cols needed: Taxable status split, taxable income total, salary/wages total
- [ ] `fetch_population_ucl.py` — QGSO UCL population (QLD towns)
  - QRIS URL: https://www.qgso.qld.gov.au/statistics/theme/population/population-estimates/state-regions
  - File: `estimated-resident-population-urban-centre-locality-qld-XXXX.xlsx`
  - One row per UCL, years as columns 2001–current
- [ ] `fetch_population_erp.py` — ABS ERP SA2/LGA (Narrabri NSW)
  - ABS API: https://api.data.abs.gov.au/ (ERP dataset, filter by SA2/LGA code)
  - Narrabri LGA code: 15750, SA2: 102021116
- [ ] `fetch_population_nrw.py` — QGSO Surat/Bowen Basin population reports
  - Surat: https://www.qgso.qld.gov.au/statistics/theme/population/population-estimates/surat-basin
  - Bowen: https://www.qgso.qld.gov.au/statistics/theme/population/population-estimates/bowen-basin
  - Annual PDF/Excel reports — NRW on-shift + FTE by UCL
- [ ] `fetch_unemployment.py` — ABS SALM smoothed SA2/LGA unemployment
  - data.gov.au: `small-area-labour-markets-australia-december-quarter`
  - Quarterly xlsx, filter by SA2 code or LGA
- [ ] `fetch_housing_qld.py` — QGSO QRIS housing (sales, rent, approvals)
  - QRIS API candidate: https://www.qgso.qld.gov.au/api/  (investigate)
  - Fallback: scrape download links from QRIS regional database pages
  - Indicators: detached dwelling median sale price, number of sales, 3-bed median rent, residential building approvals
- [ ] `fetch_housing_nsw.py` — Narrabri housing (NSW SCA rent + Valuer General sales)
  - SCA: https://www.facs.nsw.gov.au/resources/housing-data/rent-and-sales
  - ABS building approvals: data.gov.au `building-approvals-australia`
- [ ] `fetch_crime_qps.py` — QPS reported offences rates by division
  - Direct download: https://www.police.qld.gov.au/maps-and-statistics (investigate API/CSV)
  - Monthly rates per 100,000 by QPS division → annual average for 5 categories
  - Categories: All offences, Drug, Good order, Theft, Traffic
- [ ] `fetch_crime_nsw.py` — BOCSAR LGA offences (Narrabri)
  - BOCSAR: https://bocsar.nsw.gov.au/pages/bocsar/crimestats.aspx
  - LGA-level monthly data, filter Narrabri LGA
- [ ] `fetch_rainfall.py` — BOM monthly totals by station
  - BOM API: http://www.bom.gov.au/climate/data/ or ACORN-SAT
  - Station numbers already in towns.toml (Roma=43091, Chinchilla=42078, etc.)
- [ ] `fetch_business.py` — ABS 8165 counts of businesses by SA2 by turnover band
  - ABS: https://www.abs.gov.au/statistics/economy/business-indicators/counts-australian-businesses-including-entries-and-exits/latest-release
  - Annual release, SA2 level, PP/NPP split from industry codes
- [ ] `fetch_fuel.py` — RACQ average ULP/RULP prices by centre
  - RACQ annual fuel price report (PDF) — may need manual extraction
  - Fallback: https://www.racq.com.au/cars-and-driving/driving/fuel-prices
- [ ] `fetch_schools.py` — ACARA school enrolments by school → aggregate to town
  - ACARA data portal: https://www.acara.edu.au/reporting/national-report-on-schooling-in-australia/national-report-on-schooling-data-portal
  - File: `school-profile-YYYY.xlsx` — filter by suburb/postcode
- [ ] `transform/to_csv.py` — convert cache JSON → 2-row CSV format per indicator per town

### Geography
- [ ] Download ABS ASGS Edition 3 (2021) shapefiles for SA2, SA3, LGA boundaries
  - Source: https://www.abs.gov.au/statistics/standards/australian-statistical-geography-standard-asgs-edition-3/jul2021-jun2026/access-and-downloads/digital-boundary-files
  - Files needed: SA2_2021_AUST_GDA2020.zip, LGA_2021_AUST_GDA2020.zip
  - Needed for boundary maps in booklets and website
- [ ] Verify SA2 codes in towns.toml against downloaded shapefiles
- [ ] Generate boundary map images per town (SA2, UCL, postcode overlays)

### Booklets
- [ ] Create booklet template (replaces manual Word/Excel/Acrobat process)
  - Cover page + credits
  - CSG Development Story section (narrative — human input required)
  - Indicator sections: Population, Employment, Income, Housing, Safety & Wellbeing
  - Each section: summary box, data trends, community insights, chart page
  - Appendices: boundary maps, rainfall, project info
- [ ] Chart generation from CSVs (replace Excel charts with matplotlib/plotly)
- [ ] PDF assembly — combine indicator pages into final booklet
- [ ] Page numbering (Roman numerals for intro, Arabic from p.5)
- [ ] File naming and archiving conventions from original process doc

### Website
- [ ] Get sFTP credentials for boomtown-indicators.org server
- [ ] Understand current site structure and how CSVs are consumed
- [ ] Automate CSV upload to server after successful fetch run

---

## Data source automation notes

Most sources are automatable. Rough difficulty rating:

| Source | Automatable? | Notes |
|--------|-------------|-------|
| ATO Table 6 & 8 | ✓ Easy | CKAN API, stable dataset slugs |
| ABS ERP (Narrabri) | ✓ Easy | ABS API, filter by geography code |
| ABS SALM unemployment | ✓ Easy | data.gov.au CKAN, quarterly xlsx |
| ABS business counts 8165 | ✓ Easy | ABS API or data.gov.au |
| BOM rainfall | ✓ Easy | BOM climate data API, station numbers known |
| QGSO UCL population | ✓ Easy | Direct CSV from qgso.qld.gov.au/issues/{N}/ — issue number increments each release, fetcher scrapes page as fallback |
| QGSO NRW reports | ~ Medium | Annual xlsx on QGSO site, URL pattern predictable |
| QPS crime rates | ✓ Easy | S3 bucket: open-crime-data.s3-ap-southeast-2.amazonaws.com/Crime%20Statistics/division_Reported_Offences_Rates.csv — updated monthly, no auth |
| BOCSAR crime (NSW) | ~ Medium | Web form download, may be scriptable with requests |
| QGSO housing (QRIS) | ~ Hard | QRIS web form with region/date selections |
| NSW housing (SCA) | ~ Hard | Quarterly xlsx, multiple files per year |
| RACQ fuel prices | ✗ Manual | Annual PDF report — extract or maintain manually |
| ACARA schools | ~ Medium | Annual xlsx on data portal, stable URL pattern |

---

## Parked questions / decisions needed

### 1. ABS Statistical Area boundaries for booklet maps

The ABS ASGS Edition 3 (2021) provides SA1/SA2/SA3/SA4 boundaries as shapefiles.
SA2 is already used for population and unemployment data. SA1 is much more granular
(~400 people) and could enable more precise local maps in booklets.

Decision needed: do we need SA1 data, or is SA2 sufficient for boundary maps?

Search: https://www.abs.gov.au/statistics/standards/australian-statistical-geography-standard-asgs-edition-3/jul2021-jun2026/access-and-downloads/digital-boundary-files

### 2. Equivalent data sources for NSW and Victoria

Several indicators currently only fetch QLD data. Equivalent sources for NSW/VIC:

| Indicator | QLD source | NSW equivalent | VIC equivalent |
|---|---|---|---|
| Population - town | QGSO UCL ERP | ABS ERP by UCL (national file) | ABS ERP by UCL (national file) |
| Population - LGA | QGSO ERP | ABS ERP by LGA | ABS ERP by LGA |
| Crime - all | QPS division rates (S3 CSV) | BOCSAR LGA offences | Crime Statistics Agency (CSA) |
| Housing - sales | QGSO QRIS | NSW Valuer General | REIV / DSS |
| Housing - rent | QGSO QRIS | NSW FACS rent tables | DHHS rent reports |
| Building approvals | QGSO QRIS | ABS 8731 (national) | ABS 8731 (national) |
| NRW population | QGSO Surat/Bowen Basin reports | No direct equivalent (Narrabri only) | No equivalent |

Note: ABS ERP UCL data is national — the QGSO UCL fetcher could be extended to
cover NSW/VIC towns (Narrabri, Shepparton, Yarram) from the same national ABS file.
See: https://www.abs.gov.au/statistics/people/population/regional-population/latest-release


---

## Data sources — detailed notes

### QPS Crime Statistics (Queensland)

**Dataset:** Reported Offence Rates per 100,000 persons by QPS Division, monthly from July 2001

**Direct download URLs (no authentication required, updated monthly):**
```
Division rates:
https://open-crime-data.s3-ap-southeast-2.amazonaws.com/Crime%20Statistics/division_Reported_Offences_Rates.csv

QLD state benchmark:
https://open-crime-data.s3-ap-southeast-2.amazonaws.com/Crime%20Statistics/QLD_Reported_Offences_Rates.csv
```

**Source page:** https://www.data.qld.gov.au/dataset/offence-rates-police-divisions-monthly-from-july-2001
**Licence:** Creative Commons Attribution 3.0

**Column mapping to website CSVs** (values are rates per 100,000 — divide by 100 to get per 1,000):

| Website CSV | Source column | Notes |
|---|---|---|
| `Crime rate - all offences.csv` | Sum of all offence category columns | Annual = mean of 12 monthly rates |
| `Drug offences.csv` | `Drug Offences` | Annual = mean of 12 monthly rates |
| `Good order offences.csv` | `Good Order Offences` | Annual = mean of 12 monthly rates |
| `Theft.csv` | `Other Theft (excl. Unlawful Entry)` | Annual = mean of 12 monthly rates |
| `Traffic offences.csv` | `Traffic and Related Offences` | Annual = mean of 12 monthly rates |

**QPS Division → Town mapping** (from towns.toml `qps_division` field):

| Town | QPS Division |
|---|---|
| Roma | Roma |
| Chinchilla | Dalby |
| Dalby | Dalby |
| Miles | Miles |
| Tara | Tara |
| Wandoan | Wandoan |
| Wallumbilla | Wallumbilla |
| Goondiwindi | Goondiwindi |
| Moranbah | Moranbah |
| Dysart | Dysart |
| Toowoomba | Toowoomba |

Note: Chinchilla shares the Dalby division. Toowoomba sub-areas all share the Toowoomba division.

---

### QGSO UCL Population (Queensland)

**Dataset:** Estimated Resident Population by Urban Centre and Locality, Queensland, 2001–current

**Direct download URL:**
```
https://www.qgso.qld.gov.au/issues/{N}/estimated-resident-population-urban-centre-locality-qld-2001-{year}p.csv
```
The issue number `{N}` increments with each annual release. The fetcher scrapes
https://www.qgso.qld.gov.au/statistics/theme/population/population-estimates/state-regions
as a fallback to find the current URL automatically.

**Encoding:** latin-1 (file uses Windows cp1252 em-dash character)
**Year suffix:** `p` = provisional, `r` = revised — both stripped to plain year in output

---

### ATO Taxation Statistics (National)

**Table 8** — Median and average taxable income by postcode (all individuals)
**Table 6A** — Selected items by taxable status and postcode (split by Taxable/Non-Taxable)
**Table 6B** — Selected items combined totals by postcode

**Source:** data.gov.au CKAN API
```
https://data.gov.au/data/api/3/action/package_show?id=taxation-statistics-{year}
```
Dataset slug pattern: `taxation-statistics-2022-23` (update year each cycle)

ATO data lags ~18 months — 2022-23 is the most recent as of early 2026.
Financial year YYYY-YY maps to calendar year YY+1 in output CSVs (e.g. 2022-23 → 2023).


---

## Notes

- ATO data lags ~18 months — 2022-23 is the most recent available as of early 2026
- QGSO housing/unemployment data requires semi-manual download (web form navigation)
- Cache is excluded from git — re-run fetchers after a fresh clone
- Toowoomba sub-areas (Central, Harlaxton, West) share postcode 4350 — income data
  is identical across them; differences appear in SA2-based indicators (crime, housing)
- QPS crime data is rates per 100,000 population, monthly — annual value = mean of 12 months
- BOM station numbers per town: Roma=43091, Chinchilla=42078, Miles=42023, Wandoan=35029,
  Tara=42104, Dalby=41240, Wallumbilla=43043, Toowoomba=41529, Moranbah=34035,
  Dysart=35109, Narrabri=54038, Goondiwindi=41507 (to 2019), 41559 (2020 onwards)