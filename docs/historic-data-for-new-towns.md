# Historic data for towns added to the workbook (October 2026)

Tested with Emerald (also Charleville, Mitchell). Code on branch `new-towns` (history-1.zip).

| Sheet / series | Source of history | Status |
|---|---|---|
| Population ERP (UCL, SA2, LGA) | QGSO UCL file, QRSIS, ABS (already in the fetchers) | Filled 2001-2025 by the writers' new-row back-fill |
| Non-resident workers, LGA | QGSO Bowen/Surat NRW by LGA (2006-2025) | Filled (Central Highlands needed the state-suffix name fix) |
| Non-resident workers, UCL | QGSO FTE table: latest two years only | No history online; earlier reports gone (ask QGSO if needed) |
| Projections | Workbook uses 2018 edition (no longer published); 2025 edition available (LGA low/med/high, SA2 medium) | Decision for the academic |
| Employment (SALM) | SALM SA2/LGA files hold 2011-2025 | Filled |
| Business counts | ABS 8165.0 SA2-by-turnover cube, 14 releases June 2009-June 2025 (pre-2019 on ABS archive site) | New on-request fetcher `business_history`; 12/17 years exact vs hand-made Roma/Chinchilla/Moranbah, rest ~1% (workbook used an earlier release) |
| Income (postcode) | ATO data.gov.au package `taxation-statistics-postcode-data`, one Table 6 per year 1994-95 to 2023-24 | New on-request fetcher `income_history`; Roma: avg taxable 23/23, ABN rows 14/14, wages 21/24 (workbook mixes taxable-only years 2012-13, 2013-14). 2010-11 avg taxable not derivable |
| Housing | QGSO QRSIS region series (26 years) | Fetcher now adds new towns' LGA/SA2; fills on writer run |
| Crime, rainfall, education | Existing fetchers carry history | Filled |
| Fuel | RACQ report (Emerald only); Qld fuel price open data (data.qld.gov.au, monthly site prices from 2019) could cover Mitchell/Charleville | data.qld.gov.au blocked from the cloud sandbox (WAF challenge); untested |

Writers fill a row that has no figures before the latest year from each source's history (base.backfill_new_row).
