# Data Source Reference List — Boomtown Indicators

Organised by XLSX updater. Each updater consumes one or more fetchers, which in turn pull from the data sources listed below.

---

## 1. update_business.py

**Fetcher:** `fetch_business.py`

| # | Data Source | Publisher | URL / API |
|---|---|---|---|
| 1 | Counts of Australian Businesses, including Entries and Exits — Data Cube 9 (SA2 by industry division and turnover) | Australian Bureau of Statistics (ABS) | `https://www.abs.gov.au/statistics/economy/business-indicators/counts-australian-businesses-including-entries-and-exits/{release}/8165DC09.xlsx` |

---

## 2. update_crime.py

**Fetchers:** `fetch_crime_bocsar.py`, `fetch_crime_qps.py`

| # | Data Source | Publisher | URL / API |
|---|---|---|---|
| 1 | Recorded Criminal Incidents — NSW-wide | NSW Bureau of Crime Statistics and Research (BOCSAR) | `https://bocsar.nsw.gov.au/content/dam/dcj/bocsar/documents/open-datasets/Incident_by_NSW.xlsx` |
| 2 | Recorded Criminal Incidents — by LGA | NSW Bureau of Crime Statistics and Research (BOCSAR) | `https://bocsarblob.blob.core.windows.net/bocsar-open-data/RCI_offencebymonth.xlsm` |
| 3 | Reported Offence Rates by police division (monthly, from Jul 2001) | Queensland Police Service (QPS) via Queensland Government open data | `https://open-crime-data.s3-ap-southeast-2.amazonaws.com/Crime%20Statistics/division_Reported_Offences_Rates.csv` |
| 4 | Regional Population (used for per-capita crime rates) | Australian Bureau of Statistics (ABS) | `https://www.abs.gov.au/statistics/people/population/regional-population/2024-25/32180DS0004_2001-25.xlsx` |

---

## 3. update_employment.py

**Fetchers:** `fetch_salm_unemployment.py` (SA2 + LGA), `fetch_qrsis_labour.py` (QLD LGA + state), `fetch_nsw_labour.py` (NSW state)

| # | Data Source | Publisher | URL / API |
|---|---|---|---|
| 1 | Small Area Labour Markets (SALM) — Smoothed SA2 Datafiles (ASGS 2021) | Department of Employment and Workplace Relations (DEWR) | `https://www.dewr.gov.au/employment-research/small-area-labour-markets` |
| 2 | Small Area Labour Markets (SALM) — Smoothed LGA Datafiles (ASGS 2025) | Department of Employment and Workplace Relations (DEWR) | `https://www.dewr.gov.au/employment-research/small-area-labour-markets` |
| 3 | Labour Force — Small Area collection (collgrp_id=12, coll_id=1953): QLD LGA + Queensland state unemployment rates | Queensland Government Statistician's Office (QGSO) — QRSIS Regional Database | `https://statistics.qgso.qld.gov.au/pls/qis_public/` (6-step wizard POST API) |
| 4 | Labour Force Status by Sex, NSW — Table 002 (62020002.xlsx): NSW state unemployment rate (mean of 12 monthly Original series) | Australian Bureau of Statistics (ABS) | `https://www.abs.gov.au/statistics/labour/employment-and-unemployment/` (latest release page scraped for `62020002.xlsx`) |

---

## 4. update_housing.py

**Fetchers:** `fetch_qgso_housing.py` (QLD), `fetch_narrabri_approvals.py` (Narrabri approvals), `fetch_narrabri_sales_rent.py` (Narrabri sales/rent)

| # | Data Source | Publisher | URL / API |
|---|---|---|---|
| 1 | QRSIS Regional Database — Collection 1925: Residential land and dwelling sales (SA2, quarterly, Sep 2000 – Sep 2025) | Queensland Government Statistician's Office (QGSO) | `https://statistics.qgso.qld.gov.au/pls/qis_public/` (6-step wizard POST API) |
| 2 | QRSIS Regional Database — Collection 1929: Median rent (SA2, quarterly, Dec 1989 – Mar 2026) | Queensland Government Statistician's Office (QGSO) | `https://statistics.qgso.qld.gov.au/pls/qis_public/` (6-step wizard POST API) |
| 3 | QRSIS Regional Database — Collection 2075: Building Approvals Historical (LGA, monthly, Jul 2001 – Dec 2018) | Queensland Government Statistician's Office (QGSO) | `https://statistics.qgso.qld.gov.au/pls/qis_public/` (6-step wizard POST API) |
| 4 | QRSIS Regional Database — Collection 2031: Building Approvals Current (LGA, monthly, Jan 2019 – present) | Queensland Government Statistician's Office (QGSO) | `https://statistics.qgso.qld.gov.au/pls/qis_public/` (6-step wizard POST API) |
| 5 | Building Approvals LGA data — SDMX REST API (dataflows BA_LGA2018 through BA_LGA2026, no API key needed) | Australian Bureau of Statistics (ABS) | `https://data.api.abs.gov.au/rest/` (SDMX 2.1 REST API) |
| 6 | Rent and Sales Report — quarterly Sales tables (LGA-level, all dwelling types) | NSW Department of Communities and Justice (DCJ) | `https://dcj.nsw.gov.au/about-us/families-and-communities-statistics/housing-rent-and-sales/rent-and-sales-report.html` |
| 7 | Rent and Sales Report — quarterly Rent tables (LGA-level, by dwelling type and bedrooms) | NSW Department of Communities and Justice (DCJ) | `https://dcj.nsw.gov.au/about-us/families-and-communities-statistics/housing-rent-and-sales/rent-and-sales-report.html` |

**QRSIS portal page:** `https://www.qgso.qld.gov.au/statistics/queensland-regions/regional-tools-statistics/queensland-regional-database`

---

## 5. update_income.py

**Fetchers:** `fetch_income.py` (Table 8), `fetch_income_table6.py` (Table 6), `ato_release.py` (release discovery)

| # | Data Source | Publisher | URL / API |
|---|---|---|---|
| 1 | Taxation Statistics — Table 8: Median and average taxable income by state/territory and postcode (multi-year historical series) | Australian Taxation Office (ATO) via data.gov.au | `https://data.gov.au/data/api/3/action/package_show?id={slug}` (CKAN API — resolves to downloadable xlsx) |
| 2 | Taxation Statistics — Table 6A: Taxable income by postcode (taxable status split) | Australian Taxation Office (ATO) via data.gov.au | `https://data.gov.au/data/api/3/action/package_show?id={slug}` (CKAN API) |
| 3 | Taxation Statistics — Table 6B: Combined taxable income by postcode (all individuals) | Australian Taxation Office (ATO) via data.gov.au | `https://data.gov.au/data/api/3/action/package_show?id={slug}` (CKAN API) |
| 4 | data.gov.au catalogue — discovers newest ATO Taxation Statistics release | data.gov.au (Australian Government) | `https://data.gov.au/data/api/3/action/package_show` (CKAN package_show API) |

**Current confirmed dataset slug:** `taxation-statistics-2022-23` (Table 8)  
**Table 6 is a single-year snapshot each cycle; Table 8 spans multiple non-contiguous years.**

---

## 6. update_population_erp.py

**Fetcher:** `fetch_population_erp.py`

| # | Data Source | Publisher | URL / API |
|---|---|---|---|
| 1 | QRSIS Regional Database — Population (ERP) collection (collgrp_id=1, coll_id=1961): Estimated Resident Population, persons only, 1991–2025, ASGS 2021 boundaries | Queensland Government Statistician's Office (QGSO) | `https://statistics.qgso.qld.gov.au/pls/qis_public/` (6-step wizard POST API) |

---

## 7. update_population_nrw.py

**Fetcher:** `fetch_population_nrw.py`

| # | Data Source | Publisher | URL / API |
|---|---|---|---|
| 1 | Bowen Basin Population Report — FTE population estimates by LGA and UCL (2024–2025) | Queensland Government Statistician's Office (QGSO) | `https://www.qgso.qld.gov.au/issues/3341/` |
| 2 | Bowen Basin Population Report — Non-resident workers on shift by LGA (2006–2025) | Queensland Government Statistician's Office (QGSO) | `https://www.qgso.qld.gov.au/issues/3341/bowen-basin-population-report-tables-non-resident-workers-on-shift-local-government-area-lga-2006-2025.xlsx` |
| 3 | Surat Basin Population Report — FTE population estimates by LGA and UCL (2024–2025) | Queensland Government Statistician's Office (QGSO) | `https://www.qgso.qld.gov.au/issues/6606/` |
| 4 | Surat Basin Population Report — Non-resident workers on shift by LGA (2008–2025) | Queensland Government Statistician's Office (QGSO) | `https://www.qgso.qld.gov.au/issues/6606/surat-basin-population-report-tables-non-resident-workers-on-shift-local-government-area-lga-2008-2025.xlsx` |

---

## 8. update_population_nrw_lga.py

**Fetcher:** `fetch_population_nrw.py` (same source data, LGA-level output)

| # | Data Source | Publisher | URL / API |
|---|---|---|---|
| 1–4 | Same as update_population_nrw.py above — Bowen Basin and Surat Basin Population Reports | QGSO | `https://www.qgso.qld.gov.au/issues/3341/` and `https://www.qgso.qld.gov.au/issues/6606/` |

---

## 9. update_population_ucl.py

**Fetcher:** `fetch_population_ucl.py`

| # | Data Source | Publisher | URL / API |
|---|---|---|---|
| 1 | Estimated Resident Population by Urban Centre and Locality (UCL), QLD (2001–latest, issue number increments each release) | Queensland Government Statistician's Office (QGSO) | `https://www.qgso.qld.gov.au/issues/5496/` (direct CSV; falls back to scraping `https://www.qgso.qld.gov.au/statistics/theme/population/` if issue number changes) |

---

## 10. update_rainfall.py

**Fetcher:** `fetch_bom_rainfall.py`

| # | Data Source | Publisher | URL / API |
|---|---|---|---|
| 1 | SILO Patched Point Dataset — annual and seasonal rainfall totals | Queensland Government (Long Paddock) / Bureau of Meteorology | `https://www.longpaddock.qld.gov.au/cgi-bin/silo/PatchedPointDataset.php` (API) |
| 2 | SILO documentation | Queensland Government | `https://www.longpaddock.qld.gov.au/silo/` |
| 3 | Climate Data Online — climate averages (fallback for stations not in SILO) | Bureau of Meteorology (BOM) | `https://www.bom.gov.au/climate/averages/tables/cw_{station}.shtml` |
| 4 | Climate Data Online — weather data request (fallback) | Bureau of Meteorology (BOM) | `https://www.bom.gov.au/jsp/ncc/cdio/weatherData/av` |

---

## Additional Fetchers (not directly tied to a specific xlsx updater)

These fetchers support the project's infrastructure (spatial data, geodatabase) rather than writing to a specific Excel sheet.

| Fetcher | Data Source | Publisher | URL / API |
|---|---|---|---|
| `fetch_sa2.py` | ASGS Edition 3 (2021) — SA2 digital boundary files | Australian Bureau of Statistics (ABS) | `https://www.abs.gov.au/statistics/standards/australian-statistical-geography-standard-asgs/edition-3-july-2021-june-2026/access-and-downloads/digital-boundary-files` |
| `fetch_qld_tenure.py` | Mines and Permits Current — ArcGIS REST service | Queensland Department of Resources | `https://spatial-gis.information.qld.gov.au/arcgis/rest/services/Economy/MinesPermitsCurrent/MapServer` |
