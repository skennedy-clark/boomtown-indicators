# Historic data for added towns

Where the earlier years come from when a town is added to the workbook
(`transform/add_town.py`), checked with Emerald, Mitchell and Charleville in
October 2026. The writers fill a row that has no figures before the latest
year from each source's history (`base.backfill_new_row`); the rows below
without a fetcher have to be entered by hand.

| Series | Source of the earlier years | How it is filled |
|---|---|---|
| Population ERP (UCL, SA2, LGA) | QGSO UCL file, QRSIS, ABS Data API | Writers, 2001 onwards |
| Non-resident workers, LGA | QGSO Bowen and Surat Basin tables, by LGA, 2006 onwards | Writers |
| Non-resident workers, UCL | 2011-2018: Table 2 ("FTE population estimates ... UCLs") of each year's Bowen or Surat Basin population report on the QGSO archive site (see Links). 2019-2023: the reports are no longer online; QGSO can supply them. Latest two years: the current QGSO FTE table | By hand. Emerald: 2011 850, 2012 920, 2013 425, 2014 215, 2015 185, 2016 95, 2017 150, 2018 405. Towns outside the two basins (Mitchell, Charleville) have none |
| Population projections, 2018 edition | QGSO archive site: LGA low/medium/high and SA2 medium series (see Links). These are the workbook's existing source (every filled cell in the 2026 workbook matches). The current QGSO site has the 2025 edition only; the 2023 edition rows added in the 2026 workbook cannot be filled for a new town | `population_projections`, then `update_population_projections.py` |
| Population projections, New South Wales | NSW Planning, 2024 edition, LGA and SA2 (Narrabri's 2031, 2036 and 2041 match exactly) | `population_projections`, then `update_population_projections.py` |
| Non-resident population projections, 2022 edition | Current QGSO report (see Links), one table per LGA. Central Highlands has series A, C and D; series B is not published by LGA | By hand |
| Unemployment (SALM) | The current SA2 and LGA files hold 2011 onwards | Writers |
| Business counts | ABS 8165.0, SA2 by turnover, 14 releases back to June 2009 (before 2019 on the ABS archive site) | `run_update.py --only business_history`, then the writers. Matches the hand-made Roma, Chinchilla and Moranbah rows in 12 of 17 years; in the others the workbook holds an earlier release's figure (about 1% different) |
| Income (postcode) | ATO data.gov.au package `taxation-statistics-postcode-data`, one Table 6 per income year | `run_update.py --only income_history`, then the writers. 2010-11 has no average for taxable individuals |
| Housing | QGSO QRSIS series for the town's LGA and SA2, including non-residential dwelling units approved (SA2s from 2002, LGAs from 2007: full years only) | Writers |
| Crime, rainfall, education | The current fetchers' sources hold the history | Writers |
| Fuel | RACQ Annual Fuel Price Report for the towns it names. Earlier years: every RACQ annual report since 2011 (RACQ's reports page, see Links), back to 2009; RACQ's list of towns changes, so a town may be in older reports only (Charleville: 2009-2016, not monitored after August 2017). Other Queensland towns: the Queensland fuel price open data, site price changes from 2019 (fetcher `fuel_qld`, part of the annual run). Its annual averages agree with RACQ's to within 1.4 c/L for Goondiwindi and Roma in every year 2019-2025, and within 5 c/L for Dalby and Emerald | `run_update.py --only fuel_history`, then the writers. Emerald 2012-2025 (RACQ); Charleville 2012-2016 (RACQ), 2017-2018 blank, 2019 onwards (open data, 3 sites); Mitchell 2019 onwards (open data, 2 sites; in no RACQ report) |

## Links

- QGSO archive site (QGSO's website as at September 2019): https://archive.qgso.qld.gov.au/
- Bowen Basin population reports 2011-2018:
  https://archive.qgso.qld.gov.au/products/reports/bowen-basin-pop-report/index.html
- Surat Basin population reports 2008-2018:
  https://archive.qgso.qld.gov.au/products/reports/surat-basin-pop-report/index.html
- Projections 2018 edition, LGA:
  https://archive.qgso.qld.gov.au/products/tables/proj-pop-lga-qld/proj-pop-lga-qld-2018-edn.xlsx
- Projections 2018 edition, SA2:
  https://archive.qgso.qld.gov.au/products/tables/proj-pop-medium-series-sa2-sa3-sa4-qld/proj-pop-medium-series-sa2-sa3-sa4-qld-2018-ed.xlsx
- Non-resident population projections 2022 edition (Bowen and Galilee Basins):
  https://www.qgso.qld.gov.au/issues/3376/bowen-galilee-basins-non-resident-population-projections-2022-2026.pdf
- RACQ Annual Fuel Price Reports 2010 onwards: https://www.racq.com.au/about-us/advocacy/reports-and-research
- Queensland fuel price open data: https://www.data.qld.gov.au/dataset/fuel-price-reporting-2019-2022
  (and the later `fuel-price-reporting-*` datasets)
