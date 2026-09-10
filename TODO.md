# TODO

Working task list for boomtown-indicators. No Jira, no external tracker —
this file *is* the tracker, versioned alongside the code it describes.

**How to use it:**
- Drop anything into **Inbox** the moment you think of it. Don't stop to
  categorise — that's a separate, later pass.
- Check items off with `- [x]` rather than deleting them. A visible "done"
  trail is worth more than a tidy file, and git history already gives you
  the tidy version if you want it (`git log -- TODO.md`).
- When an item needs a decision from someone other than whoever's coding
  (the research team, a data-sourcing call), it goes under **Research /
  academic input needed**, not Fetchers — the label matters, it tells you
  who's blocking whom.

See `docs/project_proposal.md` for the formal three-extension structure this
project was proposed under, and how the sections below map onto it.

---

## Inbox
_(untriaged — add here, sort later)_

-

---

## Fetchers
- [ ] `fetch_population_erp.py` — ABS ERP for NSW/VIC towns (Narrabri, Shepparton, Yarram)
      via https://api.data.abs.gov.au/ (ERP dataset, filter by SA2/LGA code)
- [ ] `fetch_population_nrw.py` — QGSO Surat/Bowen Basin non-resident worker population
      Surat: qgso.qld.gov.au/statistics/theme/population/population-estimates/surat-basin
      Bowen: qgso.qld.gov.au/statistics/theme/population/population-estimates/bowen-basin
- [ ] `fetch_housing_nsw.py` — Narrabri housing (NSW Valuer General + FACS rent + ABS 8731)
- [ ] `fetch_crime_nsw.py` — BOCSAR LGA offences (Narrabri)
- [ ] `fetch_business.py` — ABS 8165 business counts by SA2
- [ ] `fetch_fuel.py` — RACQ average ULP prices (likely manual — annual PDF)
- [ ] `fetch_schools.py` — ACARA school enrolments by suburb/postcode

## Geospatial
- [ ] `fetch_sa2.py` — QLD + NSW SA2 boundaries (drafted; not yet committed or run
      against the live ABS source)
- [ ] `fetch_qld_tenure.py` — EPP/PL tenure via QLD ArcGIS REST (drafted; not yet
      committed or run against the live service)
- [ ] `fetch_nsw_tenure.py` — PEL/PPL/PAL via spatial.industry.nsw.gov.au, mirrors
      fetch_qld_tenure.py's layer-resolution-by-name pattern
- [ ] `fetch_qld_production.py` — 6-monthly CSG production stats (data.qld.gov.au)
- [ ] `ST_Intersects` view joining `qld_tenure` to `sa2_boundaries`
- [ ] Decide: integrate geospatial fetchers into `FETCHER_REGISTRY`, or keep as a
      separate pre-pipeline stage
- [ ] Verify Wallumbilla SA2 code (307011178 vs 307011177) against ABS boundary files

## Resource-operation start-date detection (future scope)
Goal: for any SA2 nationally, determine when CSG (and later, other resource
extraction — mining, wind farms, other energy projects) operations began,
by spatially joining facility/well point data to SA2 polygons and taking
the earliest relevant date per SA2.

Source leads below came from a Perplexity research pass the user pasted in
(2026-09) — treat as a **lead sheet, not verified sources**. Several of the
citation links in that pasted output were visibly mismatched with their
descriptions (e.g. an ASGS SA2 page cited under a WA petroleum wells claim),
so every dataset name/URL here needs to be independently confirmed by
browsing the actual portal before any fetcher is written against it — same
verification standard as the QLD tenure ArcGIS service.

- [ ] **Method (reusable across categories):** download point dataset of
      wells/facilities with a status + key date field -> spatial join to
      SA2 polygons (`ST_Intersects` in DuckDB, matching the tenure-join
      pattern) -> filter to the relevant category -> `MIN(date)` per SA2.
      Decide up front which date is being reported (drilling/spud start vs
      first production vs commissioning) since these can differ by years
      and the three are not interchangeable in a research write-up.

- [ ] **CSG / petroleum wells, by jurisdiction** (verify each before use):
  - QLD: Queensland borehole series / petroleum well locations
    (data.qld.gov.au / GSQ Open Data Portal) — overlaps with the tenure
    fetcher work already planned; check whether well-level dates are on
    this dataset or need to come from a separate DNRM source
  - NSW: "Coal Seam Gas Boreholes" and "Petroleum drillholes" on the SEED
    portal (data.nsw.gov.au) — claimed to have a specific CSG-flagged
    layer, which would be more directly usable than QLD's if confirmed
  - NT / WA / SA / VIC / TAS: state petroleum/mineral well datasets exist
    per state (NTG Open Data, WAPIMS/DMIRS, SARIG/PEPS, GSV dbMap,
    MRT/THE LIST) — CSG activity is minor-to-negligible in most of these,
    so low priority unless a specific SA2 outside QLD/NSW is in scope

- [ ] **Mining more broadly:**
  - "Australian Operating Mines Map" (data.gov.au) — national, points,
    claims an operating-period/status-date field
  - Gavin Mudd's "Comprehensive Australian mine production dataset
    (1799-2021)" (RMIT Figshare) — tabular with an Operating Period field,
    not spatial by default, would need geocoding against another mine
    location dataset
  - State tenement/operating-mine layers (QLD QSpatial, WA MINEDEX/DMIRS,
    VIC discover.data.vic.gov.au, etc.) for finer-grained timing

- [ ] **Wind farms / renewable energy:**
  - Clean Energy Regulator large-scale renewable energy data — project
    stage (probable/committed/accredited); accredited ~= generating
  - AEMO/AREMI "Network Map Renewables" — spatial (SHP/GeoJSON/KML/CSV),
    claimed to include commissioning year directly, which would make it
    the most directly usable source in this category if confirmed

- [ ] Once at least the QLD CSG well-date source is confirmed, prototype
      the SA2 join against a handful of already-known towns (e.g.
      Chinchilla, Miles — well-documented CSG history) and sanity-check
      the derived start year against the public record before trusting it
      for towns with less documented history

- [ ] **Connects to `csg_notice_year`:** `Town` in `config.py` already has
      a `csg_notice_year` field, currently set by hand in `towns.toml` —
      **decision: leave it manual for now.** Revisit borehole-based
      automation only once Extension 2/3 (QLD-wide, then Australia-wide
      spatial work) is actually underway, not before — no point building
      the detection method against one town's worth of context when the
      whole point of doing it is to scale past hand-entry.

## XLSX data/chart update (Extension 1, item e — current MVP target)
**MVP milestone reached (2026-09-10):** first real, live, end-to-end run —
`update_population_ucl.py` against the real 2025→2026 workbook, all 11
towns with UCL cache data updated correctly (Chinchilla, Dalby, Dysart,
Goondiwindi, Miles, Moranbah, Roma, Tara, Toowoomba, Wallumbilla, Wandoan).
**UPDATE: the openpyxl-produced file was found to be corrupted (see
below) — pivoted to xlwings (Excel COM automation), re-confirmed working
on real data with no repair prompt on 2026-09-10. Milestone genuinely
reached.**

Goal per project_proposal.md: idempotently update Indicators_Data-Charts.xlsx
starting from last year's file, one indicator at a time. Two separate
sub-problems, deliberately sequenced:

- [x] **Data-writing, stage 1 (openpyxl version — SUPERSEDED, see critical
      finding below):** `update_indicator_value()` prototyped and
      tested against the real workbook structure — finds (town, indicator,
      year) cell, overwrites in place if the year column exists, appends a
      new year column (both header rows) if it doesn't. Confirmed idempotent
      (repeat calls don't duplicate columns) and confirmed all 227 charts
      across every sheet survive a save unchanged. Not yet wired to real
      fetcher output — currently tested with hand-supplied values only.
- [x] **CRITICAL FINDING — openpyxl corrupts this workbook, confirmed
      unrecoverable.** Real run against the real workbook produced a file
      Excel flagged as needing repair — and Excel's own repair failed to
      open it at all. Diffing raw OOXML parts (before/after) confirmed
      real, permanent data loss on every openpyxl save: all 231 charts'
      style/colour XML, external links, threaded comments + author
      metadata (downgraded to legacy comments), custom XML parts, an
      embedded image, printer settings. This is a documented openpyxl
      limitation, not fixable by patching around individual parts.
      **Decision: openpyxl is retired for this task entirely.**
      `base_openpyxl_DEPRECATED.py` kept for reference only — do not use
      it against the real workbook.
- [x] **Pivoted to xlwings (Excel COM automation)** — `base.py` and
      `update_population_ucl.py` rewritten to drive real Excel directly
      rather than reconstruct the file, which structurally eliminates
      this entire class of problem (Excel saves its own file the way it
      always does; nothing is reconstructed by a third party). Every
      xlwings API call used was confirmed to genuinely exist in the
      library.
- [x] **CONFIRMED WORKING on real data (2026-09-10).** Ran against a
      real workbook copy with `--visible`: all 11 towns updated
      correctly, file reopened cleanly afterward with **no repair
      prompt** — the thing that failed under openpyxl. This is the real
      MVP milestone, now genuinely reached (not just the logic — the
      actual file).
- [x] **Formatting bug found and fixed:** a newly-written value inherited
      a percentage format from an adjacent cell (Excel's own behaviour,
      not an xlwings bug) — at least one town's population figure showed
      as a percentage instead of a plain number. Fixed by explicitly
      setting `number_format = "General"` on every cell written,
      including the two header cells created for a brand-new year
      column — relying on whatever format Excel happens to carry over
      isn't safe.
- [ ] Once xlwings is confirmed working: revisit performance at full
      scale (17 towns × multiple indicators × multiple sheets) — COM
      calls are slower than openpyxl's in-memory model even with the
      bulk-range-read optimisation already applied; worth timing a real
      full run before assuming it's fast enough for comfortable everyday
      use, especially by a future non-technical user running this
      unattended.
- [ ] `docs/manual_processes.md` / the eventual non-technical runbook
      needs "Excel must be installed and closed before running this" as
      an explicit prerequisite — xlwings drives a real Excel process,
      which is a meaningfully different requirement than a pure-Python
      script, and worth stating plainly for someone who isn't
      technical.
- [x] **Critical finding — sheet has THREE parallel geography sections
      (LGA / SA2 / UCL), not one flat structure.** Confirmed by Steve
      reading column A directly: each section can contain a town of the
      SAME name with the SAME indicator name — `Population (ERP)` exists
      under both LGA and SA2 for Goondiwindi, with genuinely different
      values (LGA=10,219 vs SA2=5,873, a 42% difference). The old
      `_find_town_indicator_row` returned the first match found, which
      would have silently picked an arbitrary geography level with zero
      warning. **Fixed:** now scans the whole sheet and raises a clear
      "AMBIGUOUS" error naming every matching row when more than one
      match exists, instead of guessing. Confirmed via real test: raises
      correctly on the Goondiwindi collision, and the already-working UCL
      matches (unique indicator name, no collision) are unaffected.
- [ ] **Still needed — proper `section` parameter.** The current fix
      converts silent-wrong-answer into loud-error, which is the right
      immediate safety net, but it doesn't let you actually WRITE to
      Goondiwindi's LGA or SA2 row — every ambiguous case currently just
      stops. Real fix: `update_indicator_value(..., section="SA2")` (or
      LGA/UCL) as an explicit, required disambiguator whenever a sheet
      has multiple sections, scanning section-by-section rather than
      whole-sheet.
- [ ] **Name-mapping gap found in the same read-through:** the SA2
      section's row labels are actual ABS SA2 names, which don't map 1:1
      onto `Town.name` in towns.toml. Concretely: `Toowoomba (Central)` /
      `(Harlaxton)` / `(West)` in towns.toml need to match `"Toowoomba -
      Central"`, `"North Toowoomba - Harlaxton"`, `"Toowoomba - West"` in
      the sheet — different strings entirely. Also `"Miles-Wandoan"` is a
      single shared SA2 row covering both the Miles and Wandoan towns —
      a genuine one-SA2-to-two-towns case, not a bug, but something the
      eventual SA2 ERP fetcher/wiring needs to know about explicitly.
      Worth checking whether `Town.sa2_name` in config.py already holds
      the correct SA2 label for this mapping before building anything new
      — it may already solve this.
- [ ] `fetch_population_erp.py` scope needs revisiting given all of the
      above — it's not just NSW/VIC towns (as currently scoped in
      Fetchers section), every QLD study town needs it too for the main
      `Population (ERP)` row, and it needs both LGA and SA2 variants
      depending on town, with the section-parameter and name-mapping work
      above as prerequisites before this can be wired up safely
- [ ] **Handling genuinely-new rows — DEPRIORITISED.** Confirmed by
      comparing the real 2025 vs 2026 workbooks that row/column counts
      shift year to year, so this is real, but **explicitly not the
      current job.** Steve's priority: update existing towns/indicators
      with one more year of data, no new rows, no new towns — that's the
      whole of what "updating the workbook" needs to do right now.
      Revisit only once the update-existing-data path is solid and in
      regular use.
- [ ] **Non-technical runbook — ELEVATED to a first-class requirement,
      not an afterthought.** The real continuity risk: Steve may not be
      at the workplace next year, and someone non-technical or only
      moderately technical needs to be able to follow a written recipe to
      update the xlsx (manual process currently takes many 10s of hours).
      This means "the code works" is not sufficient on its own — a
      plain-language, step-by-step document (not a developer README) is
      an equally required deliverable alongside the update tooling itself,
      not something to write up afterward once the code is "done."
- [ ] **Validated against the real 2025 file, not just synthetic test
      data:** `update_population_ucl.py` (fixed `xlsx_update.base` import)
      ran successfully against the actual uploaded `Indicators_Data-
      Charts_2025.xlsx`, correctly wrote Chinchilla and Toowoomba's 2025
      UCL values, and all charts (including this file's known-messier
      ones — Toowoomba has one fewer chart than the 2026 version) survived
      unchanged.
- [ ] Repeat for a small set of indicators across sheets (Income, Crime,
      Housing) to confirm the row-finding logic generalises — the "town
      name row has empty column B" heuristic for locating blocks needs
      checking against sheets where that pattern might not hold exactly
- [ ] **Year must not be hardcoded anywhere in the real pipeline.**
      `update_indicator.py`'s self-test hard-codes 2025, which is fine for
      a test (it needs a fixed, checkable value) but the actual calling
      code must determine the target year dynamically (e.g. from the
      current date, or "one past the latest year column already present"),
      so the same code keeps working unattended in 2026, 2027, etc.
      without a code change each year.
- [ ] **Test plan once wired to real data:** run against an actual prior
      year's saved workbook, update it to the current year, confirm the
      new year's values are correct; then re-run the *same* update again
      for the same year and confirm it's a true no-op (identical file,
      not just "doesn't crash") — this is the real idempotency test, not
      just the synthetic one in the self-test.
- [ ] **Data-writing, stage 2 — chart ranges:** each chart's series
      reference is a fixed cell range (e.g. `Income!$M$8:$Z$8`), and ranges
      are inconsistent even within one town's chart sheet (some already
      extend well past current data, some stop dead at the last populated
      year). Needs per-chart inspection before deciding whether/how to
      extend — do NOT assume a uniform "add one column" fix works
      everywhere
- [ ] Once (b) verification-against-threshold exists: decide where it sits
      relative to this — before writing to the xlsx (block bad data) or
      after (write, but flag)?

## Output / transform
- [ ] `to_csv.py` — wire up `building_approvals` transformer (18 of 25 CSVs done)

## Infrastructure
- [x] `pyproject.toml` — dependency pinning + pytest config (still uses existing
      `regional-indicators/` folder name; not yet a proper installable package —
      see note in the file itself)
- [x] `tests/` skeleton — pytest, fixtures, config + cache-index + base-fetcher
      download tests (16 passing). Caught and fixed a real bug: `Config()`
      could not load the actual `towns.toml` at all (`_load_towns` iterated
      the raw dict instead of `.values()`) — see `config.py` `_load_towns`.
- [ ] More fetcher-parsing tests (per-source, against saved fixtures — the
      base download/retry/cache logic is covered, individual fetchers' own
      parsing logic mostly isn't yet)
- [ ] `orchestrator.py` — extract `run_pipeline()` out of `run_update.py`'s `main()` so
      a GUI can call it without going through argv
- [ ] `gui/` — internal front end over `orchestrator.run_pipeline()` (Streamlit is the
      lowest-effort option for an internal tool)

## Documentation
- [ ] `docs/indicators.md` — data dictionary: workbook row → fetcher → output CSV
      column → unit → source
- [ ] `docs/data_sources/*.md` — split out of README, one file per source
- [ ] `docs/manual_processes.md` — website CSV upload process (see Website section —
      need process details before this can be written accurately)
- [ ] `CONTRIBUTING.md` — how to add a fetcher / how to add a town

## Maintenance (recurring)
- [ ] Update SALM fallback URL each quarter (`fetch_salm_unemployment.py`)
- [ ] Verify Tara/Goondiwindi/Moranbah BOM station substitutions against previous
      booklets before next publication

## Research / academic input needed
- [ ] **A.** Building approvals: adopt calendar-year aggregation throughout, or
      replicate previous booklets' financial-year aggregation for historical data?
- [ ] **B.** Toowoomba sub-area booklets: LGA-level figure or SA2-level figure for
      approvals? (SA2 is very sparse for Central/Harlaxton)
- [ ] **C.** Wallumbilla housing: correct to reuse Roma's SA2 data, or should
      Wallumbilla housing be flagged as unavailable instead?
- [ ] **D.** Goondiwindi's SA3 reclassification (30703 → 30701 in ASGS 2021) —
      flag for any cross-year regional-level analysis
- [ ] **E.** Narrabri / Shepparton / Yarram: no SALM unemployment data — no chart,
      LGA-level fallback, or a different source (e.g. Census-based)?
- [ ] **F.** Benchmark geography for Toowoomba — still Brisbane, or QLD-wide?

## Website
- [ ] Get sFTP credentials for boomtown-indicators.org server
- [ ] Document current manual CSV upload process in `docs/manual_processes.md` —
      a **separate project** will rebuild this properly; document it here so the
      current process is reproducible in the meantime, not to justify automating
      it inside this pipeline
- [ ] Once documented, decide whether automating the upload is in scope for this
      pipeline at all, or stays manual until the separate project lands