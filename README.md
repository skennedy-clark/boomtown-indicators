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

### 1. Prepare

- Set `YEAR_END` in `regional-indicators/config.py` to the new data year
  (and the matching constant in `regional-indicators/transform/to_csv.py`).
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
runs every writer against the working copy, brings the charts up to date, and
builds the website folder
`3 Web Content`. Each step's command line is printed before it runs, so any
step can be repeated on its own.

| Option | Meaning |
|---|---|
| `--test-copy <file>` | working copy to create and write into (default `test-copy.xlsx`); replaced if it exists |
| `--previous-web <folder>` | an earlier website folder; differences from it are reported |
| `--web-out <folder>` | where to build the website folder (default `3 Web Content`) |
| `--business-year <year>` | year for the Business sheet: the calendar year in which the financial year ends (default: last calendar year) |
| `--last-year <year>` | latest data year, used when updating the charts (default: last calendar year) |
| `--skip-fetch` | reuse the existing cache; do not download |
| `--visible` | show Excel while the writers run |

The run ends with one line per step, and the full output is saved to
`regional-indicators/logs/end_to_end_<date>_<time>.log`.

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

Some figures have no automated source and are entered by hand; see
`docs/manual_processes.md`.

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
| Exogenous: fuel | `update_fuel.py` | `fuel` |
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
| `regional-indicators/towns.toml` | Towns and benchmark regions with their geography codes (postcodes, SA2, LGA, police division, rainfall station). Adding a town or a benchmark region is an edit to this file; its header documents every field. |
| `regional-indicators/config.py` | Year range (`YEAR_START`, `YEAR_END`), paths, `SILO_EMAIL`. |
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
        to_csv.py              website CSVs from the cache (earlier approach; web_export.py is the current export)
        booklet/               booklet generation (Word)
    cache/                     downloads and fetched JSON (not in version control)
    logs/                      run logs (not in version control)
tests/                         pytest suite
assemble_maps/                 postcode boundary map utility
docs/                          manual processes; project proposal
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
