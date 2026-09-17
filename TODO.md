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

**North star, restated and staged out by Steve (2026-09-17), worth
re-reading whenever deep in one fetcher's details — "get it working" is a
TACTIC toward this, not the goal itself.** Data at LGA / SA2 / UCL scale are
all genuinely valid and wanted (benchmarking a town against its broader
context is intentional, not redundancy to eliminate). Design preference:
extend `towns.toml` to add a new town or a data mapping, not hardcode
per-town logic in Python — a new town should be addable by editing the toml
alone. The staged ambition, roughly in order:

1. **Automate this year's work** — the current focus, everything in
   Fetchers/XLSX data below.
2. **Generalize it to next year** — the "will it work next year" audit
   already underway (see Maintenance section): no hardcoded dates, no
   fragile assumptions that quietly go stale.
3. **Fork/extend to towns not currently in the booklet/xlsx.** `towns.toml`
   deliberately already carries towns beyond the current output set — the
   output side should eventually be extendable to include a new town by
   config alone, the same way fetching already aims to be.
4. **Extend to other states and specific new gas-affected regions** — NT
   explicitly named, **Beetaloo Basin towns** specifically flagged as a real
   candidate (a genuine NT gas basin, the same shape of thing as the
   Surat/Bowen Basin coverage already built).
5. **Go spatial** — DuckDB + GIS, connecting to the "Resource-operation
   start-date detection" future-scope section elsewhere in this file
   (spatially joining well/facility data to SA2 polygons).
6. **Ultimately: generalize to all of Australia**, not just the current
   CSG-affected Queensland/interstate town set.

Every stage above sits on top of the ones before it working reliably and
generically — a fix or a design choice made now is worth judging against
"does this help stage 3-6 too, or does it just patch stage 1" where that's
a real question, not just against whether it makes today's run pass.

---

## Inbox
_(untriaged — add here, sort later)_

- Checked whether the 5 flagged "isolated outlier" years (Isaac 2012,
  Toowoomba LGA 2013, Dysart 2015, Moranbah 2012 & 2016, Roma 2018,
  Toowoomba UCL 2023) match problems Steve already found and fixed
  manually going from 2025→2026: they don't — values are identical in
  both files, untouched by his manual pass. Isaac 2012 was subsequently
  confirmed correct against a real Bowen Basin download (see Fetchers —
  the historical ground-truth audit resolves this class of question
  wherever full source history is available; UCL-level entries still
  don't have that, see below).

---

## Fetchers
- [x] `fetch_population_nrw.py` — both regions now use real, confirmed
      URLs (2026-09-15): Surat Basin (issue 6606) and Bowen Basin
      (issue 3341) — different issue numbers per region and per report
      type, confirming issue numbers can't be guessed/incremented.
- [x] **Fixed and validated against real data (2026-09-15):** the LGA
      label and year headers turned out to be on TWO SEPARATE rows in
      the real file (`LGA(a) | Non-resident workers on-shift(b)` on
      one row, `2008 | 2009 | ...` on the next) — the original version
      searched for years in the same row as "LGA" and found nothing.
      Fixed and re-tested against a workbook built to exactly match the
      real structure Steve pasted — Maranoa/Toowoomba/Western Downs all
      parse correctly, `n.a.` values correctly excluded rather than
      crashing, units-label row correctly ignored. Also caught a second
      bug: the report's own TITLE row contains the substring "lga" too
      (via "...(LGA)..." in the title), which the original loose
      substring match grabbed instead of the real header — tightened to
      require the cell to *start with* "lga" and be short.
      **Confirmed from real data:** Toowoomba LGA has its own row in
      the Surat Basin file — added to that region's `lgas` list.
- [x] **Second real fix, FTE file (2026-09-15):** live run confirmed
      LGA-level parsing correct (13 towns, right year counts) but every
      single town showed "UCL: none". Root cause: the FTE file has
      THREE label columns — `LGA(a)` | `Location(b)` | `UCL(a)` — and
      the code was matching on `Location` (values like "In town",
      "Rural areas") instead of `UCL` (the actual place names); and
      each year spans three sub-columns (ERP | Non-resident workers
      on-shift | FTE estimate) needing the right one specifically, not
      just the first. Rewrote to find the UCL column explicitly and
      locate the correct sub-column per year group, renamed the output
      field `ucl_fte_latest` → `ucl_nrw_latest`. Tested against a
      workbook built to exactly match the real structure — Roma/
      Toowoomba/Millmerran/Oakey all correct, Injune correctly excluded
      (its 2025 figure was `n.p.`).
- [x] **Cross-validated against the real 2026 reference workbook
      (2026-09-15):** every town with a fetched UCL 2025 value matches
      the existing hand-entered figure exactly — Chinchilla 670, Dalby
      410, Dysart 2355, Miles 260, Moranbah 2625, Roma 185, Toowoomba
      120, Wandoan 145 (8/8 exact matches). The two `UCL: none` results
      (Tara, Wallumbilla) are correct too — the reference file's own
      2025 figures for those towns are also missing.
- [x] **First live run against real data (2026-09-15):** 4 clean
      writes, 8 flagged, breaking into three categories — real
      historical data-quality findings, step-changes on
      already-independently-confirmed values (Miles, Wandoan — see
      override question below), and one real bug (Toowoomba's
      hardcoded LGA sub-label was wrong for the real file — fixed,
      removed the incorrect override).
- [x] **Built the real fix, not a threshold tweak (2026-09-15):** added
      `audit_historical_series()` in `audit.py` — compares every
      existing workbook year against the FULL freshly-refetched source
      series (ground truth), instead of guessing from the existing
      series' shape alone. Tested against Steve's real Bowen Basin
      download: confirms zero discrepancy for Isaac's 2011/2012/2013
      figures, correctly clearing what had looked like a miscopy.
      Also tested: a genuine 10x typo correctly blocks
      (`MAJOR_DISCREPANCY`), a modest ~2% revision correctly notes
      without blocking (`MINOR_DISCREPANCY` — "note but don't assume
      wrong"), a 1-unit rounding difference is ignored entirely.
      Supersedes the old shape-based guess when available;
      `SCALE_MISMATCH` (wrong-row detection) still blocks independently
      either way. Wired into the LGA-level write path (full history
      available); **UCL-level writes still use the shape-based check
      only** — the FTE/UCL source only gives the latest year, no
      history to ground-truth against.
- [x] **`--deep-audit` flag built and tested (2026-09-15):** for any
      flagged UCL-level entry, cross-checks the corresponding LGA's own
      history in the same year(s) — a real regional workforce event
      should show up at both levels, a UCL-only blip is more likely a
      genuine error specific to that cell. Tested end-to-end with real
      Isaac/Moranbah data: correctly reports 2012 as corroborated
      (+26% LGA-level swing matching the flagged UCL point) and 2016 as
      weaker evidence (-9%, real but modest). Context for the human
      reviewing a flag, not a verdict.
- [x] **Confirmed working against the real file end-to-end
      (2026-09-15):** re-run after both fixes above — 6 written
      (Western Downs LGA, Maranoa LGA, Isaac LGA, Toowoomba LGA,
      Chinchilla UCL, Dalby UCL), 6 flagged (all UCL-level: Dysart
      2015, Miles, Moranbah 2012 & 2016, Roma 2018, Toowoomba UCL 2023,
      Wandoan), exactly as predicted once the ground-truth fix landed.
- [x] **Split into two files (2026-09-15), per the earlier Inbox note:**
      `update_population_nrw.py` now handles UCL-level only,
      `update_population_nrw_lga.py` (new) handles LGA-level only.
      Shared write logic (`write_one`, `deep_audit_context`) promoted
      into `base.py` so neither script duplicates it. Regression-tested
      against the same synthetic data used to validate the combined
      version — both splits produce identical results (LGA dedup still
      works, UCL/LGA sub-label handling unaffected).
- [ ] **Still open: override/force mechanism.** Miles and Wandoan are
      independently confirmed correct (matched the 2026 reference file
      exactly) but still sit flagged with no way to say "write it
      anyway" other than manually clearing the cell first. The 4
      remaining UCL flags (Dysart 2015, Moranbah 2016, Roma 2018,
      Toowoomba UCL 2023) are genuinely still unverified either way —
      `--deep-audit` gave weak/no corroboration for all four, unlike
      Moranbah's 2012 (strong corroboration, still unwritten).
- [x] `fetch_population_erp.py` — built (2026-09-10), **but QLD-only,
      not the NSW/VIC ABS fetcher originally planned under this name**
      (see naming-collision note below). Currently only works via a
      working fallback: reads a manually-assembled
      `cache/qgso_and_bom_{YEAR}.xlsx` if present (documented in
      `docs/manual_processes.md`). No live-API path yet —
      `fetch_qgso_housing.py` already has a working QRSIS integration
      pattern for a different collection (housing); the population
      collection id needs discovering the same way before a live path
      can be added here.
- [x] **Real progress toward automating this (2026-09-16):** QGSO
      publishes a machine-readable master index of every QRSIS
      collection — https://statistics.qgso.qld.gov.au/report-viewer/run?__report=sis-stats-available.rptdesign&systemName=QRSIS&__format=xls
      (an Excel-XML file, not modern xlsx — parses fine as plain text/
      XML). Confirmed the target collection genuinely exists: "Population
      (ERP)(a) persons only", group "Population Estimates", **SA2-level**,
      2016 ASGS geography (current — there's also a superseded 2011-ASGS
      version, don't use that one), data 1991-2025, updates every March.
      **No LGA-level version of this exact collection found in the
      index** — towns needing LGA-level ERP (e.g. Brisbane) may need a
      differently-named collection, still unconfirmed.
      **Still needed:** the index doesn't expose the internal numeric
      collection id the actual query API needs (same kind of id as
      housing's 1925/1929/2075/2031) — extending `fetch_qgso_housing.py`'s
      exact query mechanism to auto-discover this collection's id
      requires its real source code, not available in the sandbox this
      was investigated in after a reset. Get that file into the next
      session before attempting the live-API build.
- [x] **Got the real source, fully understood the query mechanism
      (2026-09-16, same session as above).** `fetch_qgso_housing.py`'s
      constants: `COLLGRP_ID = "22"` for housing specifically —
      confirms each subject area has its own group id, and Population's
      is genuinely unknown. Its `_select_collection()` POSTs a
      *known* `coll_id` (e.g. `1925`) straight to
      `QIS1110W$COLL.ProcessCollection` — it never demonstrates how to
      browse/discover collections within a group; that was presumably
      done once, manually, the same way originally for housing. No
      public documentation of QRSIS's internal `collgrp_id`/`coll_id`
      values exists anywhere searched. **Precise remaining gap:** the
      `collgrp_id` and `coll_id` for "Population Estimates" / "Population
      (ERP)(a) persons only". Everything else in the query mechanism
      (series selection, time period, region matching, submission,
      output parsing) is already proven working code, directly reusable
      once those two numbers are known. **Fastest way to get them:**
      one short devtools session — Network tab, select the collection
      in the wizard, find the POST to `QIS1110W$COLL.ProcessCollection`,
      read `collgrp_id` and `coll_id`/`sel_coll_name` from its form
      data. Full step-by-step given to Steve in chat.
- [x] **Steve found the real values directly from the live wizard HTML
      (2026-09-16):** `collgrp_id="1"` (Population group), `coll_id="1961"`
      ("Population (ERP)(a) persons only", 1991-2025, **ASGS 2021**
      boundaries — newer than the 2016 ASGS version the earlier catalog
      index suggested; older ASGS-2016 (id 1298) and ASGS-2011 versions
      also exist, deliberately not used). `fetch_population_erp.py`
      rewritten to use these — full live QRSIS query, reusing
      `fetch_qgso_housing.py`'s exact proven wire format (duplicated for
      now rather than shared, to avoid risking the working housing
      fetcher without live-testing capability — worth extracting into a
      shared module once this one's also confirmed working). Falls back
      to the manual-file path automatically if the live query fails —
      though that's now a pure safety net rather than the primary path.
- [x] **CONFIRMED WORKING LIVE END-TO-END (2026-09-16).** `population_erp`
      now fully automated, no manual file needed at all. First attempt
      with guessed `period="Annual"`/bare-year dates failed silently
      (empty region list, no visible error — a real gap, fixed by adding
      response logging to the time-period and region-type steps).
      Corrected hypothesis — ABS ERP is published as-at-30-June, QGSO's
      own catalog labels the frequency "Financial Year" — confirmed
      correct on the very next run: **`period="Financial Year"`,
      `from_date="Year Ended 30 Jun 2001"`, `to_date="Year Ended 30 Jun
      2025"`, `date_fmt="Y1"`.** Real series name confirmed:
      `'Persons (Persons)'`. 14/14 towns fetched real 2025 SA2-level ERP
      figures, including genuinely distinct values for all three
      Toowoomba sub-areas (14,434 / 6,596 / 17,905) — **resolves the
      oldest open SA2 name-mapping gap in the project**, flagged weeks
      ago as blocking this exact indicator. Also found and fixed a
      false-positive in the new error-detection logging: generic
      frameset-fallback HTML ("upgrade your browser") was matching a
      crude "error|invalid" pattern — tightened to look for actual
      QRSIS/Oracle error markers (`ORA-NNNNN`, `no data found`, etc.)
      instead of generic English words.
- [x] **The 2026-09-16 "confirmed" values were wrong, just not caught
      until they broke (2026-09-17) — corrected against Steve's real
      manual walkthrough of the whole QRSIS wizard, not another guess.**
      `date_fmt` is actually `"Y2"` for this collection, not `"Y1"`;
      `from_date`/`to_date` are plain year numbers ("2025", "2001"),
      never "Year Ended 30 Jun YYYY" — that compound string apparently
      happened to work on 09-16 but stopped once `to_date` became a
      dynamic year (2027) beyond what QRSIS actually has, which
      silently breaks the ENTIRE session (empty region list for every
      town, not just an out-of-range result) rather than gracefully
      clipping. **Real fix, more robust than either previous guess:**
      the fetcher now reads the real "To Date" dropdown directly off
      the live Time Periods page and uses whatever QRSIS itself says
      is newest — self-updating every year with no code change needed,
      rather than computing a guessed future date at all. Caught and
      fixed a bug in this discovery logic itself during testing: From
      Date and To Date share the same generic `<select>` element name,
      so naively taking "the first select with year options" grabbed
      From Date's default (1991, the oldest) instead — fixed by
      anchoring on the `to_date` hidden-field marker specifically.
      **Also rewrote `_parse_output_html` and `_submit_report`'s
      requested `display_style`** to match the REAL confirmed output
      table shape (Period as rows, region labels as columns, one table
      per series) — the original parser assumed a completely different,
      never-actually-confirmed region-grouped structure.
      **Tested against the real, verbatim server response Steve
      captured** (not a synthetic reconstruction): all four regions in
      his sample output (2 LGA, 2 SA2) parsed with exact correct
      values matching what his browser showed, full pipeline through
      `_aggregate` confirmed correct town/year/value/sa2_name mapping.
      Real, hard evidence this time, not another hypothesis — next live
      run is the actual confirmation.
- [ ] Verify the `_parse_output_html` change didn't matter for the
      earlier run that DID work end-to-end (2026-09-16, before this
      turn's fixes) — that run's actual response shape was never
      directly inspected, only its final written values were confirmed
      correct, so it's possible the old parser was accidentally correct
      for that specific case. Not expected to cause a regression (the
      new parser is more precisely matched to confirmed real structure),
      but worth keeping in mind if anything looks different this time.
- [ ] **Next: wire this into the workbook.** Targets the SA2-section
      `Population (ERP)` row — the exact same indicator name/section
      already confirmed to collide with the LGA-section version for
      Goondiwindi (42% value difference, LGA=10,219 vs SA2=5,873) —
      the still-outstanding `section` parameter for
      `_find_town_indicator_row` (see XLSX data/chart update section)
      is a real prerequisite here, not optional, since every write for
      this indicator needs to land in the SA2 section specifically,
      not whichever section happens to match first.
- [x] **`qgso_housing` status corrected (2026-09-16):** a full fetcher
      sweep showed this is NOT "0 towns ok" as previously recorded —
      sales and rent collections both work correctly end-to-end (11
      real SA2 regions matched, real prices parsed). Only the two
      approvals collections fail, and specifically: the exact same SA2
      codes that matched fine for sales/rent come back
      `"not in QRSIS list"` for approvals. Strong, narrow signal —
      building-approvals data on QRSIS likely only publishes at LGA
      level, not SA2, which lines up directly with Research Question B
      below (LGA vs SA2 for Toowoomba sub-area approvals). Worth
      checking what region list those two `udqctl_id`s actually expose
      before assuming a broader fix is needed.
- [x] **`bom_rainfall` status corrected (2026-09-16):** also not
      broadly broken as the old "wrong approach" note suggested — 16
      of 17 towns succeed cleanly (a few "synthetic data" warnings are
      SILO's own gap-filling, not a bug). The one real failure is
      narrow: Yarram's configured SILO station number (85151) is
      invalid, and the name-search fallback found nothing genuinely
      matching "Yarram" in Victoria. Worth a manual look at SILO's
      station list for the correct Yarram, VIC station.
- [x] **Full sweep confirms all three xlsx write scripts consistent and
      reproducible (2026-09-16):** ran `update_population_nrw.py`,
      `update_population_nrw_lga.py`, and `update_population_ucl.py`
      (the separate Population-ERP-UCL indicator) back to back against
      the same real file. All three matched previously-confirmed
      results exactly — NRW UCL 2 written/6 flagged, NRW LGA 4
      written/0 flagged, Population-ERP-UCL 10 written/1 flagged
      (Chinchilla's formula cell, same one found originally — still
      unresolved, see below).
- [ ] **Minor labeling nuance found in this sweep:** re-running an
      already-written value still logs as `WRITTEN`, not distinguished
      from a genuine new write — functionally correct (nothing's
      actually overwritten incorrectly, `CellState.MATCHES_NEW_VALUE`
      still triggers a write of the identical value) but worth a
      clearer log label eventually (e.g. "already correct, no-op" vs
      "WRITTEN") so a re-run's output is easier to read at a glance.
- [ ] **Chinchilla's formula cell — still an open decision, now
      confirmed on multiple separate runs.** `=(Y54-W54)/W54` sits in
      the Population-ERP-UCL row and has blocked every single write
      attempt so far. Worth actually resolving: confirm nothing else
      references it, then clear it — or decide it stays a permanent
      manual-entry cell going forward, rather than leaving it
      perpetually flagged.
- [ ] **NAMING COLLISION to resolve:** this TODO originally planned
      `fetch_population_erp.py` as the NSW/VIC ABS-based fetcher (line
      below). That name is now taken by the QLD/QGSO fetcher instead.
      The NSW/VIC one still needs building — needs a different
      filename (e.g. `fetch_population_erp_abs.py`) and a different
      FETCHER_REGISTRY key (`population_erp` is taken).
- [ ] ABS ERP for NSW/VIC towns (Narrabri, Shepparton, Yarram) via
      https://api.data.abs.gov.au/ (ERP dataset, filter by SA2/LGA
      code) — still not built, see naming note above
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
- [x] **`section` parameter built and tested (2026-09-16).**
      `_find_town_indicator_row`/`write_one` now accept `section=
      "LGA"/"SA2"/"UCL"`, tracking which section each row sits in as it
      scans (the sheet marks each with its own single-word header row).
      Tested against a mock reproducing the real Goondiwindi collision:
      no section still correctly raises AMBIGUOUS (unchanged, backward
      compatible); `section="LGA"` and `section="SA2"` now correctly
      resolve to the two different rows; an unmatched section correctly
      raises not-found rather than crashing; existing UCL matching
      (unique indicator name, no section needed) unaffected. Built
      `update_population_erp.py`, targeting `section="SA2"`, matching by
      `sa2_name` (the real ABS SA2 label) rather than `town.name`.
      **Caught a real bug while building this:** `fetch_population_erp.py`
      never actually captured the matched SA2 *name*, only the code —
      would have silently written under the wrong label for every town
      whose SA2 name differs from `town.name` (exactly the Toowoomba
      sub-areas this feature exists to fix). Fixed: `_match_regions` now
      parses and returns the real name (tested against all 7 real
      region strings from the live run, including the hyphenated ones —
      "Toowoomba - Central", "North Toowoomba - Harlaxton", "Broadsound
      - Nebo" — all extracted correctly), threaded through `_aggregate`
      and `_write_results` into the cache JSON's new `sa2_name` field.
      **NOT YET TESTED against a live Excel instance** — the section
      logic itself is proven via mock, but the actual write against the
      real workbook hasn't run yet. Also: re-run `fetch_population_erp.py`
      first before testing the wiring script, since its cache output
      format changed (added `sa2_name`) — existing cache files from the
      earlier live run won't have it.
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
- [x] **Validated against the real 2025 file, not just synthetic test
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
- [x] **Year must not be hardcoded anywhere in the real pipeline.**
      Confirmed satisfied by every script built this session (fetchers
      determine the latest year from the data itself; write scripts take
      whatever year the fetcher's cache gives them) — no literal year
      values in any control-flow logic.
- [x] **Test plan, ground-truth version:** ran against an actual prior
      year's real workbook (`Indicators Data-Charts 2025.xlsx`), updated
      it with real fetched data, confirmed values correct, re-ran the
      same update again and confirmed no duplicate writes (cells that
      already match the new value are correctly treated as a no-op via
      `CellState.MATCHES_NEW_VALUE`).
- [ ] **Data-writing, stage 2 — chart ranges:** each chart's series
      reference is a fixed cell range (e.g. `Income!$M$8:$Z$8`), and ranges
      are inconsistent even within one town's chart sheet (some already
      extend well past current data, some stop dead at the last populated
      year). Needs per-chart inspection before deciding whether/how to
      extend — do NOT assume a uniform "add one column" fix works
      everywhere

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
      parsing logic mostly isn't yet); worth adding real coverage for
      `audit.py`'s functions too now that they're load-bearing
- [ ] `orchestrator.py` — extract `run_pipeline()` out of `run_update.py`'s `main()` so
      a GUI can call it without going through argv
- [ ] `gui/` — internal front end over `orchestrator.run_pipeline()` (Streamlit is the
      lowest-effort option for an internal tool)
- [ ] Documentation convention: file docstring headers should name the
      file with its full path from repo root (e.g.
      `regional-indicators/transform/xlsx_update/base.py`), not just the
      bare filename — makes navigating name collisions much easier
      (there are two files named `base.py`). Applied consistently to
      every `xlsx_update/` file; not yet retrofitted to older files
      (e.g. `fetchers/base.py`, which currently just says
      `fetchers/base.py` without the `regional-indicators/` prefix).
- [ ] Refactor at some point: remove deprecated code such as the old
      openpyxl-based xlsx updater (`base_openpyxl_DEPRECATED.py` and any
      references to it) now that xlwings is the confirmed-working approach
- [ ] `requirements.txt` is incomplete compared with `pyproject.toml`: it
      omits `xlwings`, test dependencies, and `duckdb`

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
- [x] **"Will this work next year?" audit prompted by Steve (2026-09-16)
      — found and fixed one real bug, flagged one pre-existing one.**
      Fixed: `fetch_population_erp.py`'s `TO_DATE` was hardcoded to
      `"Year Ended 30 Jun 2025"` — would have silently capped the
      fetcher at 2025 forever, never erroring, just quietly going
      stale. Now computed as `datetime.now().year + 1`, confirmed to
      evaluate correctly (`Year Ended 30 Jun 2027` as of this date) —
      a RANGE query, so asking one year ahead is safe and self-
      maintaining, never needs a manual date bump again.
      **Not fixed, deliberately left alone:** `fetch_qgso_housing.py`
      has the exact same pattern, in currently-working, proven code —
      `sales.to_date = "Year Ended 30 Sep 2025"`, `rent.to_date =
      "Year Ended 31 Mar 2026"`, `approvals_curr.to_date = "Jan 2026"`
      are all hardcoded literals that will silently stop capturing new
      years once those dates pass. Not touched here since it can't be
      live-tested from this environment and the risk of breaking
      working code outweighs fixing it blind — needs the same
      `datetime.now()`-based fix, applied and tested on a machine that
      can actually run it against the live QRSIS endpoint.
- [ ] Same question worth asking of every other fetcher in the project,
      not just these two — `fetch_population_nrw.py`'s hardcoded issue
      numbers (Surat Basin 6606, Bowen Basin 3341) are a related but
      different risk: not a wrong date range, but a URL that will
      eventually 404 once QGSO issues next year's report under a new
      issue number. It already has a page-scrape fallback for exactly
      this (`_scrape_for_url`), but that fallback itself has never been
      proven working — only the hardcoded direct URLs have been tested
      live. Worth deliberately testing the fallback path (e.g.
      temporarily pointing at a wrong issue number) before trusting it
      to actually work when the real URLs do go stale.

## Research / academic input needed
- [ ] **A.** Building approvals: adopt calendar-year aggregation throughout, or
      replicate previous booklets' financial-year aggregation for historical data?
- [ ] **B.** Toowoomba sub-area booklets: LGA-level figure or SA2-level figure for
      approvals? (SA2 is very sparse for Central/Harlaxton)
- [x] **C — partially resolved, side effect flagged (2026-09-17).**
      Cross-checked against the truth workbook: "Roma" and "Roma Surrounds"
      are both real, distinct SA2 rows with genuinely different Population
      (ERP) figures (7,083 vs 6,287) — confirms this was a real data-
      correctness gap, not just a label choice. Fixed in `towns.toml`:
      Wallumbilla's `sa2_code`/`qgso_sa2` corrected to `307011177`
      (Roma Surrounds' actual ASGS 2021 code, confirmed via QGSO's own
      PDF and ABS Census QuickStats — the old `307011178` was genuinely
      Tara's code, `307011176` was Roma township's own). **Not yet known:
      whether this breaks Wallumbilla's housing fetch** — `qgso_sa2` is
      shared with `fetch_qgso_housing.py`, so this change also switches
      Wallumbilla from reusing Roma's housing data to trying "Roma
      Surrounds" directly. If that SA2 has no sales/rent data in QRSIS
      (plausible for a rural/dispersed area), housing may now fetch
      nothing for Wallumbilla instead of a working proxy. Test
      `qgso_housing` after this change, not just `population_erp`, before
      treating this as fully resolved — if housing does break, the real
      fix is probably a separate field (e.g. only override the SA2 used
      for population, keep housing on the Roma proxy deliberately) rather
      than one shared code serving every indicator identically.
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