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
      a `csg_notice_year` field (currently set by hand in `towns.toml`),
      which presumably drives the vertical "CSG activity started" line
      already present on many of the workbook's charts. Once the borehole
      spatial-join method above is confirmed for a town, the natural next
      step is treating its output as a *candidate* value for
      `csg_notice_year` — worth deciding whether it replaces the manual
      entry outright or is surfaced as a suggestion for someone to confirm
      (Steve: you cut off mid-thought here — "looking at borehole
      activities in ___" — worth finishing that when you're back at it,
      in case there's a specific dataset/scope in mind beyond what's
      already listed above)

## XLSX data/chart update (Extension 1, item e — current MVP target)
Goal per project_proposal.md: idempotently update Indicators_Data-Charts.xlsx
starting from last year's file, one indicator at a time. Two separate
sub-problems, deliberately sequenced:

- [x] **Data-writing, stage 1:** `update_indicator_value()` prototyped and
      tested against the real workbook structure — finds (town, indicator,
      year) cell, overwrites in place if the year column exists, appends a
      new year column (both header rows) if it doesn't. Confirmed idempotent
      (repeat calls don't duplicate columns) and confirmed all 227 charts
      across every sheet survive a save unchanged. Not yet wired to real
      fetcher output — currently tested with hand-supplied values only.
- [ ] Wire `update_indicator_value()` to actual fetcher/to_csv output for a
      first real indicator (Population ERP is the natural pilot — already
      has a working fetcher and the cell layout is mapped)
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