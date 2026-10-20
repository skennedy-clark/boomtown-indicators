# boomtown-indicators

Data pipeline for the regional indicators project of the UQ Centre for
Natural Gas. The project tracks social and economic indicators for towns
affected by coal seam gas and other resource development in Queensland,
New South Wales and Victoria, and publishes them as a set of town booklets
and on the web:

- https://boomtown-indicators.org/compare
- https://gas-energy.centre.uq.edu.au/resources/tools

The indicators are held in one Excel workbook with a sheet per topic
(Population, Employment, Housing, Income, Business, Crime, Exogenous). Once
a year the workbook is extended by one year. This repository automates that
update:

1. **Fetch** the latest figures from each public source into a local cache.
2. **Write** the new year into the workbook, auditing every value before it
   is written.
3. **Export** the website's CSV files from the workbook.
4. **Build** the town booklets (in development).

**To use it:** download the latest release from
<https://github.com/skennedy-clark/boomtown-indicators/releases/latest>
(**Assets → Source code (zip)**), unzip it outside OneDrive, and follow
[`docs/annual_update_runbook.md`](docs/annual_update_runbook.md). The rest of
this page is technical detail.

## Contents

- [Requirements](#requirements)
- [Installation](#installation)
- [Annual update](#annual-update)
- [Running the steps individually](#running-the-steps-individually)
- [Configuration](#configuration)
- [Data sources](#data-sources)
- [Tests](#tests)
- [Repository layout](#repository-layout)
- [Known limitations](#known-limitations)

## Requirements

| Requirement | Needed for | Notes |
|---|---|---|
| Python 3.10 or later | everything | 3.11+ recommended (`tomllib` is built in; on 3.10 the `tomli` backport is installed automatically) |
| Microsoft Excel, on Windows or macOS | the workbook writers | The writers drive Excel through `xlwings`. They do not run on Linux or without Excel. |
| Internet access | the fetchers | All sources are public; no accounts or API keys are required. |
| Git | obtaining the code | |

The fetchers, the website export and the tests run on any platform,
including Linux. Only the step that writes into the workbook needs Excel.

### Python packages

Installed by `pip install -r requirements.txt`:

| Package | Used for |
|---|---|
| `requests` | all downloads |
| `beautifulsoup4`, `lxml` | parsing source web pages and the QGSO regional database |
| `pandas` | reading source spreadsheets and CSVs |
| `openpyxl` | reading workbooks (source files, and the indicators workbook for the website export) |
| `xlwings` | writing to the indicators workbook through Excel |
| `pypdfium2` | reading the RACQ fuel report (PDF); rendering map PDFs for booklets |
| `Pillow` | booklet images |
| `python-docx` | booklet documents |
| `tomli` | reading TOML on Python 3.10 only |
| `pytest` | tests |

Optional, for the stand-alone geography utilities only:

| Package | Used for |
|---|---|
| `duckdb` | `fetchers/fetch_sa2.py`, `fetchers/fetch_qld_tenure.py` |
| `geopandas`, `matplotlib`, `contextily` | `assemble_maps/postcode_plot.py` (see `assemble_maps/README.md`) |

## Installation

```bash
git clone https://github.com/skennedy-clark/boomtown-indicators.git
cd boomtown-indicators

python -m venv venv
venv\Scripts\activate          # Windows
source venv/bin/activate       # macOS / Linux

python -m pip install --upgrade pip
pip install -r requirements.txt
```

Check the installation:

```bash
python -m pytest                                      # all tests should pass
python regional-indicators/run_update.py --validate   # checks towns.toml
```

On macOS, the first run of a writer prompts for permission for Python to
control Excel; allow it.

Before the first rainfall fetch, set `SILO_EMAIL` in
`regional-indicators/config.py` to your own email address. The SILO service
uses it as the user name for its API.

Unless stated otherwise, commands in this document are run from the
repository root with the virtual environment active.

## Annual update

**The simple way (Windows):** double-click `Run annual update.bat` and choose
last year's finished workbook. Everything below happens automatically, and the
results land in a folder `Indicators update <year>` next to that workbook. The
one-page instructions for the workbook's owner are in
[`docs/annual_update_runbook.md`](docs/annual_update_runbook.md).

The rest of this section describes the same steps run by hand.

### 1. Prepare

- No code needs editing. The year to add is worked out from the starting
  workbook (the year after its latest year, but never a year that has not
  finished); `--year` overrides it.
- Have last year's delivered workbook to hand. It is the starting point and
  is never modified; the pipeline works on a copy.
- Close the workbook in Excel. If the files are in a synchronised folder
  (OneDrive, Dropbox), pause synchronisation or work in a local folder.

### 2. Run everything

```bash
python regional-indicators/run_end_to_end.py "Indicators Data-Charts 2025.xlsx" \
    --test-copy "Indicators Data-Charts 2026.xlsx" \
    --previous-web "path/to/last year/3 Web Content"
```

This copies the starting workbook to the working copy, runs every fetcher,
runs every writer against the working copy, brings the charts up to date,
builds the website folder, and gathers every source file into the data folder
`1 Data`
`3 Web Content`. Each step's command line is printed before it runs, so any
step can be repeated on its own.

| Option | Meaning |
|---|---|
| `--test-copy <file>` | working copy to create and write into (default `test-copy.xlsx`); replaced if it exists |
| `--output-folder <folder>` | put everything for the year in one folder: `1 Data`, `2 Workbook` (the finished workbook), `3 Web Content` and `Run report.txt` |
| `--output-root <folder>` | as `--output-folder`, with the folder created inside this one and named `Indicators update <year>` |
| `--year <year>` | the data year to add (default: worked out from the starting workbook) |
| `--previous-web <folder>` | an earlier website folder; differences from it are reported |
| `--web-out <folder>` | where to build the website folder (default `3 Web Content`) |
| `--data-out <folder>` | the data folder: source files by topic, and the place for manual downloads (default `1 Data`) |
| `--business-year <year>` | year for the Business sheet: the calendar year in which the financial year ends (default: last calendar year) |
| `--last-year <year>` | latest data year, used when updating the charts (default: last calendar year) |
| `--skip-fetch` | reuse the existing cache; do not download |
| `--visible` | show Excel while the writers run |

The run ends with one line per step, and the full output is saved to
`regional-indicators/logs/end_to_end_<date>_<time>.log`. A plain-English run
report (what was produced, and every value flagged for checking) is written
to the output folder, or beside the log.

### 3. Review

`OK` against a step means that it ran to completion. Its summary line gives
the number of values written and the number flagged.

A **flagged** value was not written. The writer reports why: the target
cell already holds a different value or a formula, the new value is out of
line with the existing series, or the row could not be identified
unambiguously. Check each flagged value against its source, then either
enter it by hand or, if the fetched value is correct, record it in
`regional-indicators/verified_overrides.toml` and rerun that writer.

Writers add the new year only. Where a source has revised earlier years,
the differences are reported and the earlier values in the workbook are
left as they are.

Some figures have no automated source. Before the run, put the manual
downloads in the data folder (`1 Data/<topic>/`, each with a "read me" that
says what to download and from where); the run copies them to where the
fetchers look. See also `docs/manual_processes.md`.

## Running the steps individually

### Fetch

```bash
python regional-indicators/run_update.py                     # all fetchers
python regional-indicators/run_update.py --only fuel schools # selected fetchers
python regional-indicators/run_update.py --skip crime_qps    # all but these
python regional-indicators/run_update.py --only income --towns Roma Dalby
python regional-indicators/run_update.py --force             # ignore the cache and download again
python regional-indicators/run_update.py --dry-run           # list what would run
python regional-indicators/run_update.py --list-cache        # show cached downloads
python regional-indicators/run_update.py --validate          # check towns.toml only
```

Results are written as JSON under `regional-indicators/cache/<area>/`.
Downloads are cached; `--force` downloads them again.

### Write to the workbook

Each writer takes the workbook and the cache folder it reads. All accept
`--visible`. Run them from `regional-indicators/transform/xlsx_update/` or
give the full path, as below (`W` and `C` are shown for brevity).

```bash
W=regional-indicators/transform/xlsx_update
C=regional-indicators/cache
BOOK="Indicators Data-Charts 2026.xlsx"

python $W/update_population_ucl.py      "$BOOK" $C/population
python $W/update_population_nrw.py      "$BOOK" $C/population
python $W/update_population_nrw_lga.py  "$BOOK" $C/population
python $W/update_population_erp.py      "$BOOK" $C/population
python $W/update_population_erp_lga.py  "$BOOK" $C/population
python $W/update_rainfall.py            "$BOOK" $C/rainfall
python $W/update_education.py           "$BOOK" $C/schools
python $W/update_fuel.py                "$BOOK" $C/fuel
python $W/update_crime.py               "$BOOK" $C/crime
python $W/update_employment.py          "$BOOK" $C/unemployment
python $W/update_housing.py             "$BOOK" $C/housing
python $W/update_income.py              "$BOOK" $C/ato
python $W/update_business.py            "$BOOK" $C/business 2025
```

On Windows `cmd`, substitute the paths in full; in PowerShell use
`$W = "..."` and `python "$W\update_fuel.py" ...`.

| Sheet | Writer | Fetchers that supply it |
|---|---|---|
| Population (town) | `update_population_ucl.py` | `population_ucl` |
| Population (non-resident workers) | `update_population_nrw.py`, `update_population_nrw_lga.py` | `population_nrw` |
| Population (SA2) | `update_population_erp.py` | `population_erp` |
| Population (LGA) | `update_population_erp_lga.py` | `population_erp_lga` |
| Exogenous: rainfall | `update_rainfall.py` | `bom_rainfall` |
| Exogenous: education | `update_education.py` | `schools` |
| Exogenous: fuel | `update_fuel.py` | `fuel`, `fuel_qld` |
| Crime | `update_crime.py` | `crime_qps`, `crime_bocsar` |
| Employment | `update_employment.py` | `salm_unemployment`, `qrsis_labour`, `nsw_labour` |
| Housing | `update_housing.py` | `qgso_housing`, `narrabri_approvals`, `narrabri_sales_rent` |
| Income | `update_income.py` | `income`, `income_table6` |
| Business | `update_business.py` | `business` |

Each module's docstring describes the sheet layout it expects and the rules
it applies.

### Export the website files

```bash
python regional-indicators/transform/web_export.py "Indicators Data-Charts 2026.xlsx" \
    --out "3 Web Content" \
    --compare-with "path/to/last year/3 Web Content"
```

Builds one folder per town containing one CSV per indicator. The workbook
is only read, so Excel is not needed. Which workbook row feeds which file
is defined in `regional-indicators/web_export_map.csv`.

| Option | Meaning |
|---|---|
| `--out <folder>` | output folder (default `3 Web Content`) |
| `--map <file>` | export map (default `web_export_map.csv`) |
| `--compare-with <folder>` | report differences from an earlier export |
| `--last-year <year>` | last year to export (default: the current year); later columns, such as projections, are excluded |
| `--dry-run` | report only; write nothing |

### Audit the chart ranges

```bash
python regional-indicators/transform/chart_audit.py "Indicators Data-Charts 2026.xlsx" \
    --out chart_audit.csv --last-year 2025
```

Lists every chart series in the workbook with the years it shows and the
years its row has data for, and writes one line per series to a CSV file.
Most charts use Excel chart filters, so a year can be missing from a chart
either because it is filtered out (`UNHIDE`) or because it lies beyond the
series' range (`EXTEND`). The workbook is only read, so Excel is not needed.
`--reference <workbook>` adds the columns each series shows in a second
workbook. The statuses are described in the module docstring.

### Bring the charts up to date

```bash
python regional-indicators/transform/xlsx_update/update_charts.py "Indicators Data-Charts 2026.xlsx" \
    --last-year 2025 --dry-run                      # list the changes
python regional-indicators/transform/xlsx_update/update_charts.py "Indicators Data-Charts 2026.xlsx" \
    --last-year 2025 --page Chinchilla              # one page
python regional-indicators/transform/xlsx_update/update_charts.py "Indicators Data-Charts 2026.xlsx" \
    --last-year 2025                                # every page
```

Runs the audit and applies it through Excel: years with data that are
filtered out of a chart are shown again, and series whose data runs past
their range get a longer range. Years without data stay hidden, and chart
formatting is not touched. The audit is run again after saving, and any
series still out of date is listed. This step is part of
`run_end_to_end.py`, after the writers. It uses the Excel object model
(chart filters need Excel 2013 or later) and has been written for Excel on
Windows.

### Add a town

1. Add the town to `regional-indicators/towns.toml` (codes and stations;
   the file's header explains each field) and run the fetchers for it,
   including the three that fetch earlier years of business counts,
   postcode income and RACQ fuel prices (they are not part of the annual
   run):

   ```bash
   python regional-indicators/run_update.py --towns Emerald
   python regional-indicators/run_update.py --only business_history income_history fuel_history
   ```

2. Add it to the workbook, copied from a similar town (its page, its rows
   on each data sheet, and its lines in `web_export_map.csv`):

   ```bash
   python regional-indicators/transform/add_town.py "Indicators Data-Charts 2026.xlsx" \
       --town Emerald --like Moranbah --dry-run     # list what would be added
   python regional-indicators/transform/add_town.py "Indicators Data-Charts 2026.xlsx" \
       --town Emerald --like Moranbah               # add it (through Excel)
   ```

   The original workbook is saved first as "... (before adding Emerald).xlsx".
   Where the new town shares an area with an existing one (an LGA or SA2),
   the existing rows are used. The new rows are empty.

   To copy a town from a different template later, run it again with the
   other town and `--replace`: the page is copied afresh, and the town's
   rows and figures are kept.

   ```bash
   python regional-indicators/transform/add_town.py "Indicators Data-Charts 2026.xlsx" \
       --town Emerald --like Dysart --replace
   ```

3. Run the writers (or the whole update). A row with no figures yet
   (or only the latest year's) receives the source's full history, not
   only the latest year. What has no history to fill:
   - non-resident workers by urban centre: QGSO publishes only the two
     latest years, and the earlier reports are no longer online;
   - population projections: the workbook holds the 2018 edition, which
     QGSO no longer publishes; the 2025 edition is available but would
     differ from the other towns' rows;
   - fuel prices for a town RACQ has never reported before 2019 (the
     Queensland open data used for it starts in 2019). A town RACQ
     reported in earlier years (Charleville, to 2016) gets those years
     from the earlier reports; years with no figure stay blank.

4. Run the chart step for the new page. Charts whose series were
   pointed at the new town's rows show empty slots after the last year
   until this is done, and any axis whose fixed limit cuts off the data
   is set back to automatic:

   ```bash
   python regional-indicators/transform/xlsx_update/update_charts.py "Indicators Data-Charts 2026.xlsx" \
       --page Emerald
   ```

5. Check the new page: text that names the template town and the CSG
   notice line, which is placed by hand. The add_town report lists what
   to look at. The rainfall station row on the Exogenous sheet is
   labelled by add_town with the town's own station (`bom_station` in
   towns.toml, named from BOM's stations list), for example "Mitchell
   Post Office 043020".

### Build a booklet

```bash
cd regional-indicators
python transform/booklet/make_booklet.py --town Chinchilla
python transform/booklet/make_booklet.py --town Roma --date "June 2026"
```

Writes `regional-indicators/booklets/<slug>/<Town>_Indicators_<year>.docx`.
Booklet generation is in development and does not yet cover every page of
the published booklets.

## Configuration

| File | Purpose |
|---|---|
| `regional-indicators/towns.toml` | Towns and benchmark regions with their geography codes (postcodes, SA2, LGA, police division, rainfall station). A new town starts here (see "Add a town"); its header documents every field. |
| `regional-indicators/config.py` | Year range (`YEAR_END` defaults to the last complete year; the end-to-end run sets it from the starting workbook), paths, `SILO_EMAIL`. |
| `regional-indicators/web_export_map.csv` | One line per website CSV: town folder, file name, and the sheet, section, block and row it is read from. |
| `regional-indicators/verified_overrides.toml` | Values checked by hand that are to be written despite an audit flag. |
| `regional-indicators/manual_rainfall_data.toml` | Monthly rainfall for stations that are not in the SILO dataset. |

Towns currently configured:

- **Queensland:** Roma, Chinchilla, Dalby, Miles, Tara, Wandoan,
  Wallumbilla, Goondiwindi, Moranbah, Dysart, Toowoomba (and its Central,
  Harlaxton and West sub-areas); Brisbane as a benchmark.
- **New South Wales:** Narrabri.
- **Victoria:** Shepparton, Yarram.

## Data sources

| Fetcher | Indicator | Source |
|---|---|---|
| `population_ucl` | Town (UCL) resident population | Queensland Government Statistician's Office (QGSO) |
| `population_erp` | SA2 resident population | QGSO Queensland Regional Database (QRSIS) |
| `population_erp_lga` | LGA resident population | ABS Data API (`ERP_LGA<year>`) |
| `population_nrw` | Non-resident workers | QGSO Surat Basin and Bowen Basin population reports |
| `salm_unemployment` | Unemployment, SA2 | Small Area Labour Markets (DEWR) |
| `qrsis_labour` | Unemployment, Queensland LGAs and state | QGSO Queensland Regional Database |
| `nsw_labour` | Unemployment, New South Wales | ABS Labour Force |
| `qgso_housing` | Sales, median price, rent, building approvals | QGSO Queensland Regional Database |
| `narrabri_approvals` | Building approvals, Narrabri | ABS Data API |
| `narrabri_sales_rent` | Sales and rent, Narrabri | NSW Department of Communities and Justice |
| `income`, `income_table6` | Taxable income by postcode | ATO Taxation Statistics (data.gov.au) |
| `business` | Business counts | ABS Counts of Australian Businesses |
| `crime_qps` | Offence rates, Queensland | Queensland Police Service open data |
| `crime_bocsar` | Offence rates, New South Wales | NSW Bureau of Crime Statistics and Research |
| `bom_rainfall` | Rainfall | SILO Patched Point Dataset (Bureau of Meteorology stations) |
| `fuel` | Average petrol price | RACQ Annual Fuel Price Report |
| `fuel_qld` | Average petrol price, for towns RACQ does not name (from 2019) | Queensland fuel price reporting open data (data.qld.gov.au) |
| `fuel_history` | Average petrol price, earlier years for a town just added (on request) | RACQ Annual Fuel Price Reports, 2011 onwards |
| `schools` | Enrolments and teaching staff | ACARA School Profile |

Source URLs, file layouts and the method used for each indicator are
documented in the docstring of the corresponding module in
`regional-indicators/fetchers/`. `Data Source Reference List.md` and
`Statistical_Areas_notes.md` give further background on sources and
geography.

Several sources publish under a URL that changes with each release. The
fetchers look for the current release where they can; where a URL has to
be updated by hand, the constant is marked `Update each cycle` in the
fetcher.

## Tests

```bash
python -m pytest            # whole suite
python -m pytest tests/test_web_export.py -v
```

The tests need neither Excel nor network access. Writer logic is exercised
against `tests/fake_xlwings_sheet.py`, a read-only stand-in for an Excel
sheet; the workbook files themselves are not part of the repository.

## Repository layout

```
regional-indicators/
    run_end_to_end.py          whole annual update in one command
    run_update.py              runs the fetchers
    config.py                  settings; loads towns.toml
    towns.toml                 towns, benchmarks and geography codes
    web_export_map.csv         workbook row -> website CSV
    verified_overrides.toml    hand-verified values
    manual_rainfall_data.toml  hand-entered rainfall
    fetchers/                  one module per source; base.py is the shared base class
    transform/
        xlsx_update/           one writer per sheet section
            base.py            locating rows and columns; writing through Excel
            audit.py           checks applied before each write
            update_charts.py   brings chart series up to the latest year
        web_export.py          website CSV folder from the workbook
        chart_audit.py         lists chart series that do not show the latest year
        add_town.py            adds a town to the workbook, copied from a similar town
        data_folder.py         builds the "1 Data" folder of source files
        to_csv.py              website CSVs from the cache (earlier approach; web_export.py is the current export)
        booklet/               booklet generation (Word)
    cache/                     downloads and fetched JSON (not in version control)
    logs/                      run logs (not in version control)
tests/                         pytest suite
assemble_maps/                 postcode boundary map utility
docs/                          runbook, handover notes, manual processes, project proposal
```

## Known limitations

- **Writers require Excel.** There is no supported way to run the workbook
  step on Linux or in CI.
- **Earlier years are not revised.** Writers add the new year and report,
  but do not apply, revisions that sources have made to earlier years.
- **Rainfall.** Stations that are not in the SILO dataset (currently
  Goondiwindi and Moranbah) need monthly figures entered in
  `manual_rainfall_data.toml`. SILO values are patched (gap-filled) and can
  differ slightly from the Bureau of Meteorology's raw station record.
- **Shared SA2s.** Miles and Wandoan share one SA2, so their SA2-level
  figures are identical. Dalby and Dysart use larger SA2s (Wambo and
  Broadsound-Nebo) that extend well beyond the town.
- **Coverage outside Queensland.** Several fetchers are Queensland-only.
  Narrabri is covered through separate New South Wales sources; the
  Victorian towns are configured but not covered for most indicators.
- **Derived rows.** Rows that the workbook calculates and benchmark rows
  without a fetcher are not yet updated automatically.
- **Booklets.** Generation is in development.
