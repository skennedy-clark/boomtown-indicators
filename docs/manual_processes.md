# Manual processes

Steps in the current workflow that this pipeline does not automate. Documented
here so the process is reproducible by someone unfamiliar with the project,
per the project's rule that anything not automated must have exact manual
instructions rather than being silently skipped.

---

## CSV upload to boomtown-indicators.org

**Status:** manual. A separate project will rebuild this properly. This
section exists to make the *current* process reproducible in the meantime —
not to justify automating it inside this pipeline. Whether automation belongs
here at all is a decision for after this section is filled in (see TODO.md,
Website section).

**Confirmed:** the uploaded CSVs are horizontal — a year header row followed
by a single data row per indicator (matching the long-format blocks already
seen in `Indicators_Data-Charts_2026.xlsx`'s data-page sheets: e.g. the
`Crime` sheet's `Total offences (person, property, other)` row against year
columns). This is consistent with `to_csv.py`'s stated "2-row quoted CSVs in
`output/{town}/`" — so the upload format likely already matches, or is close
to, what the pipeline emits. To confirm exactly, once at your desk:

- [ ] A real example CSV as currently uploaded (filename + contents) — to
      diff against what `to_csv.py` currently emits and confirm they match
- [ ] Which CSVs get uploaded — everything in `output/{town}/`, or a specific
      subset per town?
- [ ] Upload method — sFTP, a web form, manual copy into a CMS, something else?
- [ ] Destination path and file-naming convention the site expects (does it
      match the `output/{town}/{indicator}.csv` naming already used locally,
      or does it need renaming/restructuring first?)
- [ ] Who currently performs this step, and how often (per booklet cycle?
      on demand?)
- [ ] Where credentials/access currently live (README already flags sFTP
      credentials as outstanding — is there an existing manual process using
      different access, or is this genuinely blocked on getting credentials?)

Once these are answered, this section should read as a numbered, followable
procedure — see the QRSIS section below for the target level of detail.

---

## QGSO Regional Database (QRSIS) — fallback when the automated fetcher can't run

`fetch_qgso_housing.py` covers this automatically for the towns and series
listed in `docs/data_sources/qgso_housing.md`, via QRSIS's underlying API
(not just the browser wizard — see that fetcher's `COLLECTIONS` dict for
confirmed collection ids: 1925=sales, 1929=rent, 2075/2031=building
approvals). **Currently reported at 0 towns ok** per project status notes —
not confirmed working end-to-end, worth fixing before building further
automation on the same API.

`fetch_population_erp.py` (main SA2/LGA `Population (ERP)` row) does NOT
yet have a live-API path — its collection id hasn't been identified. It
only implements the manual-file fallback below. The manual fallback is:

1. Go to http://www.qgso.qld.gov.au/products/tables/qld-regional-database/index.php
2. [NEEDS INPUT — the exact wizard steps: which theme, which region type,
   which date format, which series names to select]
3. Assemble the export into a workbook containing (at minimum) a sheet
   named `Pop`, `Population`, or `Pop sheet`, with:
   - Collection header: `Population (ERP)(a) persons only`
   - A header row containing a `Region` column and a year column (a
     bare 4-digit number, e.g. `2026`)
   - One row per region: region name | ... | value for that year
   - **Single year snapshot only** — this file does not contain
     history, which is fine for this project's actual need (one more
     year at a time)
4. Save the resulting file as `cache/qgso_and_bom_{YEAR}.xlsx` (current
   year) — or `cache/qgso_and_bom_{YEAR-1}.xlsx` if last year's export
   is what's available
5. Re-run `python run_update.py --only population_erp` — **confirmed
   implemented**: the fetcher searches for either filename, then
   searches the sheet for its header row rather than assuming a fixed
   position (the row position has already been observed to shift year
   to year)

**Path to full automation** (see `fetch_population_erp.py`'s docstring
for the complete version): once `qgso_housing`'s QRSIS calls are
confirmed working again, the same collection-id-discovery method
documented there (browser devtools, watch the POST request while
manually selecting Population (ERP) data) should be used to find the
population collection id and add a live-API path the same way housing
already has one — rather than researching this from scratch a second
time.

### Non-resident worker (NRW) population — Surat/Bowen Basin

`fetch_population_nrw.py` covers this via direct theme-page URLs (more
tractable than the Regional Database wizard, since these two pages have
stable URLs), but is **not yet live-tested**. If it fails to find the
current file:

1. Surat Basin (Western Downs, Maranoa towns): go to
   https://www.qgso.qld.gov.au/statistics/theme/population/non-resident-population-queensland-resource-regions/surat-basin
2. Bowen Basin (Isaac — Moranbah, Dysart): go to
   https://www.qgso.qld.gov.au/statistics/theme/population/non-resident-population-queensland-resource-regions/bowen-galilee-basins
3. Download the latest "FTE LGA UCL" xlsx from whichever page is
   relevant
4. Save as `cache/surat_basin_nrw.xlsx` or `cache/bowen_basin_nrw.xlsx`
5. Re-run `python run_update.py --only population_nrw`

---

_Add further manual fallback processes below as they come up — same format:
numbered steps, exact URLs, exact filenames, and an explicit `[NEEDS INPUT]`
marker for anything not yet confirmed rather than a guess._