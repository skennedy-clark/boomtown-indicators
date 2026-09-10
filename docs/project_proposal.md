# Project proposal

As proposed to the supervising professor. This is the reference structure —
TODO.md tracks day-to-day tasks against it, this file records what was
actually proposed and why, so the two don't drift apart over time.

## Extension 1: Partial automation of report and web-data creation

Current manual workflow: each year, a new directory structure is created,
new data is downloaded into it, processing and data entry are done by hand,
`Indicators_Data-Charts.xlsx` is created/updated, flat CSVs for the website
are produced, and the docx booklets' charts are updated. This extension
automates as much of that as is sound to automate.

| # | Step | Status |
|---|------|--------|
| a | Automate the download of new data every year | Underway — most fetchers exist (income, crime, housing, population, BOM rainfall, ATO release detection in progress); geospatial fetchers (SA2, QLD tenure) drafted, not yet live-tested |
| b | Verify downloaded data — flag variance from prior year exceeding a threshold | Not started. Natural fit alongside the existing `CacheIndex`/fetcher pattern: compare newly fetched value against last known value per (town, indicator, year) before it's written anywhere |
| c | Store new data locally for reproducible research | Done — `cache/` + `CacheIndex` (checksums, register/has/invalidate) |
| d | Process data as required | Partial — `to_csv.py` transformers exist for some indicators (18 of 25 last checked) |
| e | Idempotently update the XLSX data and charts file | **In progress — this is the current MVP target.** Splits into two genuinely separate problems: (i) writing data values into the correct cell, idempotently, without touching anything else — prototyped and tested (`update_indicator.py`), chart-safe (227 charts survive a save unchanged); (ii) each chart's series range is a *fixed* cell reference (e.g. `Income!$M$8:$Z$8`) that doesn't automatically extend to a newly-added year column — ranges are inconsistent even across charts on the same sheet, so this needs per-chart handling, deliberately scoped as the next step after (i), not blocking it |
| f | Create the CSVs for upload to the website | Partial — `to_csv.py`, plus the current *upload* process is still manual and only partially documented (`docs/manual_processes.md`) pending real examples from Steve |
| g | Update the docx booklets with the recent charts | Code exists (`transform/booklet/`), not yet re-verified against the automated pipeline end to end |

**MVP definition (as agreed):** update `Indicators_Data-Charts.xlsx`'s data,
one indicator at a time, starting from last year's file — i.e. item (e)(i)
above, data only, charts untouched for now. Once that's solid across a
representative set of indicators, that's the thing to show as a working MVP.

## Extension 2: Spatial / GIS approach for all of QLD

DuckDB (spatial extension) + GeoPackage as the backend, chosen so the
database is a distributable file with no server required, and so the same
file opens directly in QGIS later with no conversion step. SA2 boundaries
(ABS ASGS 2021) and QLD petroleum tenure (EPP/PL, live ArcGIS REST service)
fetchers drafted; not yet run against live data or integrated with the town
config. See `TODO.md`'s Geospatial section for current state.

## Extension 3: Spatial / GIS approach for all of Australia

Extends Extension 2 nationally — same architecture, additional per-state
fetchers (NSW tenure already scoped; other states as needed). Deliberately
sequenced after QLD, since QLD is where the existing town data and domain
knowledge (tenure types, EPP/PL conversion rules) already concentrate.
