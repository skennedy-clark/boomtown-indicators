# Version 2.0 (October 2026)

Version 1.0 extended the existing towns by one year. Version 2.0 adds new
towns and fills their history.

## New towns

- Queensland: Emerald (Central Highlands), Mitchell (Maranoa), Charleville (Murweh).
- New South Wales comparison towns: Muswellbrook and Singleton (no CSG).

Each was added with `transform/add_town.py` and filled by the annual run.
What remains empty, and why, is listed by `transform/town_coverage.py`
and in `docs/historic-data-for-new-towns.md`.

## New in the pipeline

- `add_town.py`: copies a town's page, rows and website lines from a
  similar town; New South Wales layouts; `--replace`.
- History for new rows: business counts, postcode income, RACQ fuel
  prices, non-resident workers by town (from the QGSO reports, via
  `docs/non-resident-population-register.csv`).
- Population projections fetched and written: QGSO 2018 edition, QGSO
  non-resident projections 2022-2026, NSW Planning 2024. Projections
  are now under their own year (they were under other years' columns).
- Non-residential building approvals (QGSO).
- Queensland fuel prices from the state's open data, for towns RACQ
  does not report.
- Rainfall station names from BOM's station list.
- Charts: axis limits copied from another town are freed; empty spacer
  categories hidden.
- `town_coverage.py`: lists the blocks and years still missing for a town.
- Website export: each New South Wales housing block read separately.
- New South Wales building approvals read from the ABS CSV response, so
  months with no approvals are counted (small LGAs lost whole years).

## Known gaps

- Non-resident workers by town, 2019-2023: the QGSO reports are not
  online (see `docs/non-resident-population-audit.md`).
- Murweh and Charleville: no non-resident estimates exist (outside every
  QGSO study area).
- QGSO 2023-edition projection rows added by hand in the 2026 workbook:
  that edition is no longer published.
- Muswellbrook rainfall historic average: BOM publishes no average table
  for station 061168; enter by hand if one is wanted.
- New South Wales building approvals before 2019 for Muswellbrook and
  Singleton: the ABS Data API starts in July 2018 (Narrabri's earlier
  years were entered by hand). Sales and rent for 2017: the NSW reports
  start part way through 2017.
- Goondiwindi and Moranbah rainfall stations are not in SILO (as in v1).
- Narrabri Surrounds series on the Muswellbrook and Singleton pages:
  delete or replace by hand (listed by add_town).
- The audit still flags many plausible values (see TODO.md).
