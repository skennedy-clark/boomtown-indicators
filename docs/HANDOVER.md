# Handover

State of the indicators pipeline at handover (October 2026), for whoever looks
after it next. The workbook's owner needs only
[`annual_update_runbook.md`](annual_update_runbook.md); this note is for the
person they call when something breaks.

## What it does

One run turns last year's finished workbook into this year's:

1. copies the starting workbook to a working copy (never edits the original);
2. downloads the latest figures from about 18 public sources into
   `regional-indicators/cache/`;
3. writes the new year into each sheet through Excel, checking every value
   before writing it;
4. brings the 230 charts up to date (shows the new year, extends ranges);
5. exports the website CSVs (`3 Web Content`);
6. gathers every source file into `1 Data`;
7. writes a plain-English `Run report.txt`.

Entry points:

| What | How |
|------|-----|
| Whole run, for the owner | double-click `Run annual update.bat` |
| Whole run, by hand | `python regional-indicators/run_end_to_end.py <last year's workbook> --output-root <folder>` |
| One step | each step's command line is printed in the log and can be run on its own |
| Tests | `python -m pytest` (no Excel or network needed) |

`README.md` describes every command, option and source.

## Status at handover

Tested by producing the 2026 workbook from the 2025 one and comparing it,
cell by cell, with the 2026 workbook made by hand:

- every step runs to the end, on Windows, from the launcher;
- 561 of the 809 new-year cells match the hand-made workbook; the rest are
  explained under "Known gaps";
- every chart shows the new year (one exception, below);
- starting from the hand-made 2026 workbook also runs cleanly, so next year's
  starting point works.

## What will need attention next year

**Source files move.** Most sources publish under a file name or address that
changes with each release. Where possible the fetchers find the current one
themselves; where they cannot, the fallback is a constant marked
`Update each cycle` in these files:

- `fetchers/fetch_population_ucl.py`, `fetch_population_nrw.py`
- `fetchers/fetch_qgso_housing.py`, `fetch_qrsis_labour.py` (QRSIS date ranges)
- `fetchers/fetch_salm_unemployment.py`
- `fetchers/fetch_business.py`, `fetch_narrabri_approvals.py`,
  `fetch_narrabri_sales_rent.py`, `fetch_nsw_labour.py`, `fetch_crime_bocsar.py`

A source that fails shows in the run report under "Downloading the latest
figures"; the full log names the address it tried. The rest of the run still
completes, so one broken source does not stop the update.

`fetch_crime_bocsar.py` also has the population years hard-coded in two places
(`range(2001, 2026)` and two column positions); these need moving on by a year.

**Manual items, every year:**

- the SALM smoothed LGA file, only if the run could not download it, saved as
  `1 Data/Employment/salm_smoothed_lga.csv`;
- monthly rainfall for Goondiwindi (041507) and Moranbah (034035), entered in
  `regional-indicators/manual_rainfall_data.toml`.

**Excel.** The writers drive Excel (xlwings). Keep the working copy out of
OneDrive: syncing while Excel opens the file makes Excel lose it part-way
through a step. The launcher already puts the working copy in the Windows temp
folder. A step stuck on an invisible Excel dialog is stopped after 15 minutes.

## Known gaps (fine for now, listed in TODO.md)

- **Population projections** are not fetched; projection rows are left as they
  are.
- **Earlier years are not revised.** Writers add the new year and report, but do
  not apply, revisions that sources have made to earlier years. The hand-made
  workbook applies some of these.
- **Chinchilla crime** uses the Dalby police division; the hand-made figures
  differ. Left as it is by decision.
- **Non-residential building approvals** have no fetcher.
- **Rainfall**: SILO figures differ slightly from BOM's raw station data
  (Dysart, Roma, Toowoomba); BOM's historic-average page returns 404 for some
  stations, so the previous average is kept and reported.
- **Toowoomba business composite** is about 5% above the hand-made figures.
- **Narrabri rainfall chart** (page Narrabri, cell C81) had broken axis labels
  in the 2025 workbook; it is already fixed in the 2026 workbook, so it will not
  come up again.
- **Booklets** (Word) are partly built (`transform/booklet/`) and not part of
  the run.

## Checking a run against a hand-made workbook

When someone has updated a workbook by hand and wants to compare:

- `python regional-indicators/transform/chart_audit.py <new workbook> --reference <hand-made workbook>`
  compares every chart series;
- the cell-by-cell comparison used at handover compared new-year cells sheet by
  sheet; the logic is simple (cells empty in the starting workbook and numeric
  in the hand-made one) and is worth adding as a script if this is done often.

## Adding a town

`transform/add_town.py` copies a similar town's page and rows (README,
"Add a town"). The writers then fill the new rows with each source's
history; `run_update.py --only business_history income_history fuel_history`
fetches the earlier years of business counts (ABS releases back to June
2009), postcode income (ATO tables back to 2000-01) and RACQ fuel prices
(every annual report since 2011, back to 2009), which the annual run
does not need. Non-resident workers by urban centre, the 2018 population
projections and fuel prices before 2019 for towns RACQ has never
reported have no history to fill.

## Where things are

| Path | What |
|------|------|
| `regional-indicators/towns.toml` | towns, geography codes, rainfall stations |
| `regional-indicators/fetchers/` | one module per source; each docstring says where and how |
| `regional-indicators/transform/xlsx_update/` | one writer per sheet; `base.py` and `audit.py` are shared |
| `regional-indicators/transform/chart_audit.py`, `xlsx_update/update_charts.py` | charts |
| `regional-indicators/transform/web_export.py`, `web_export_map.csv` | website CSVs |
| `regional-indicators/transform/data_folder.py` | `1 Data` |
| `regional-indicators/run_end_to_end.py`, `Run annual update.bat` | the whole run |
| `regional-indicators/verified_overrides.toml` | values checked by hand, to be written despite a flag |
| `TODO.md` | the working list of open items |
