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
- [x] **Override mechanism built and tested (2026-09-17).** New
      `regional-indicators/verified_overrides.toml` — a git-tracked,
      human-editable list of specific (town, indicator, sub_label,
      year, value) combinations independently confirmed correct, each
      with a `reason` recording what kind of verification it was. Wired
      into `write_one()` (in `base.py`, so it's automatic for every
      current and future indicator, not just NRW): a flagged write
      checks the override list before giving up, and only applies if
      the value matches EXACTLY — if the fetched value ever changes,
      the override silently stops applying and the flag returns, never
      a standing "trust this cell forever" bypass. `WriteAuditReport`
      extended with an `override_reason` field; its summary line shows
      BOTH the override and the original flag reason together, so
      nothing about why a cell was flagged gets hidden just because it
      was overridden. Tested end-to-end (5 checks: real file loads
      correctly, exact match works, mismatched value/town correctly
      don't match, a flagged report correctly becomes safe-to-write
      with a matching override applied). Seeded with the two
      confirmed-strong entries: Miles and Wandoan 2025 UCL NRW (both
      independently matched the 2026 reference workbook exactly).
      Moranbah 2012 left as a commented-out candidate in the file
      itself — corroborated by `--deep-audit` (Isaac LGA +26% same
      year) but that's regional-trend evidence, not an independent
      source check the way Miles/Wandoan's was; Steve's call whether
      that's enough to add for real.
- [ ] The 4 remaining UCL flags (Dysart 2015, Moranbah 2016, Roma 2018,
      Toowoomba UCL 2023) are genuinely still unverified either way —
      `--deep-audit` gave weak/no corroboration for all four. Needs an
      actual historical-source check, not something resolvable from
      here without live access to old QGSO report editions.
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
- [x] **CONFIRMED LIVE, FULLY CLEAN (2026-09-17), all fixes verified
      working together.** `"Discovered real max To Date from the page:
      2025"` — the discovery logic worked exactly as designed, no
      guessed/hardcoded year involved. All 12 regions matched, 14/14
      towns fetched, zero failures. **Wallumbilla = 6,287** — exactly
      "Roma Surrounds"' real figure from the truth document, confirming
      the towns.toml fix and this session's QRSIS fixes work correctly
      together (previously it silently returned Roma township's 7,083
      instead). Goondiwindi (6,251) and Roma (7,083) both match the
      truth document exactly too. This indicator is genuinely done on
      the fetch side — next is wiring `update_population_erp.py`
      (already built, using the `section="SA2"` parameter) against a
      real test copy, now that the underlying data is finally
      trustworthy enough to be worth testing against.
- [ ] **Wiring confirmed working, with one small real fix needed
      (2026-09-17).** First live run against the real 2025 file: 11/14
      written correctly (including Wallumbilla's confirmed-correct
      6,287), 3 flagged as row-not-found: Dysart, Miles, Wandoan. Root
      cause: the workbook's own row labels for these two SA2 areas are
      `"Broadsound-Nebo"` and `"Miles-Wandoan"` — **no spaces around
      the hyphen** — while QRSIS itself (and the toml, until now)
      calls them `"Broadsound - Nebo"` / `"Miles - Wandoan"` — **with**
      spaces. Every other hyphenated SA2 name that worked (Roma
      Surrounds, Toowoomba - Central, North Toowoomba - Harlaxton) uses
      consistent spacing in the workbook; this is a real inconsistency
      in the workbook's own labeling, not a data error, and not a
      pattern worth generalizing a fuzzy-match around (the row-finder's
      "no guessing" contract is a deliberate safety property, not
      something to weaken for two towns). Fixed via `towns.toml`
      directly instead: `sa2_name` for Dysart/Miles/Wandoan corrected
      to the exact no-space workbook spelling, with a note on each
      explaining the QRSIS-vs-workbook discrepancy for future
      reference. Re-run the wiring script to confirm all 14 write
      clean now.
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
- [x] **Yarram's station fixed (2026-09-17/21).** 85151 turned out to
      be real and currently active — confirmed via BOM's own climate
      archive ("Yarram Airport", live daily observations) — but SILO's
      Patched Point Dataset simply doesn't include it. Found the real
      fix via SILO's own station-name search directly (the fetcher's
      own name-search fallback oddly didn't surface it): **station
      85193, literally named "YARRAM"** (-38.560, 146.670, VIC,
      elevation 18m). Independently double-confirmed by Steve against
      BOM's own daily/monthly rainfall data pages for 85193 directly —
      two separate sources agreeing, not just a name match. `towns.toml`
      corrected, re-run confirmed 17/17.
- [x] **Major policy correction (2026-09-21), prompted by Steve
      checking rainfall against the actual published Exogenous tab in
      the real workbook — do NOT auto-substitute a different SILO
      station when the true one isn't found.** Continuity with years of
      already-published data matters more than automated availability.
      Confirmed real, concrete cases the old auto-substitution policy
      got wrong: Dalby (published station 41522, silently substituted
      with 41240), Moranbah (published 34035, substituted with 34038)
      — both fixed in `towns.toml` to the true station. **New towns
      with no prior published history are the one legitimate case for
      picking the nearest SILO-available station** — that's a one-time
      decision made by hand when adding the town (e.g. via the
      `nearest-bom-station` tool, confirmed working — its data source,
      `bom.gov.au/climate/data/lists_by_element/stations.txt`, is a
      plain static catalog, genuinely not subject to the same
      anti-scraping block as BOM's interactive data portal), not
      something the fetcher should do automatically at run time.
- [x] **Investigated BOM direct access as an alternative to SILO —
      confirmed not viable, for two separate, real reasons (2026-09-21).**
      (1) BOM's Climate Data Online web portal (`jsp/ncc/cdio/weatherData`)
      explicitly blocks automated access — a real, stated policy
      ("you should stop"), not a technical hurdle to route around; using
      browser automation (Selenium/Playwright) specifically to defeat a
      detected-and-blocked automation attempt would be circumventing an
      explicit access policy, not solving a bug — decided not to pursue
      this regardless of how legitimate the underlying use is. (2) BOM's
      anonymous FTP (`ftp2.bom.gov.au/anon/gen`) was tested directly and
      connects fine, but its own README confirms it only carries
      *current* forecasts/warnings/observations/charts (radar, satellite,
      NWP, aviation products) — historical station climate records are
      explicitly a separate, PAID "registered user" product, with only
      non-real samples free. Not a path to historical rainfall data.
      Don't revisit either path without a genuinely new reason to think
      something's changed.
- [x] **Built: manual-entry fallback + full fetcher redesign
      (2026-09-21).** New `regional-indicators/manual_rainfall_data.toml`
      — a human reads monthly figures directly off BOM's CDO page (real
      browser traffic, not blocked) and enters them; the fetcher runs
      manual entries through the EXACT SAME total/summer/winter
      aggregation SILO data gets (confirmed via a full round-trip test:
      12 real monthly figures → correct total/summer/winter split,
      correctly never flagged as "synthetic" since manual entries use a
      distinct source-code marker from SILO's interpolation flag).
      `fetch_bom_rainfall.py` rewritten: auto-substitution removed
      entirely; when the true station isn't in SILO, checks the manual
      file first, and if nothing there, fails with a direct clickable
      link to that station's BOM CDO page rather than a generic error.
      Also surfaced, genuinely good news: the fetcher ALREADY computed
      summer/winter splits all along (`_aggregate()`'s `summer`/`winter`
      keys) — the earlier "we need seasonal not just yearly" concern was
      about the logs only ever printing the annual total, not a real gap
      in the underlying data.
- [x] **Steve declined the browser-automation route (Selenium against
      BOM's climate/data/ form) — correctly, and worth recording why:
      not because it wouldn't technically work, but because it's still
      automated access to a service that's explicitly stated it doesn't
      want that, just via a page their bot-detection hasn't caught up
      to yet. Building around exploiting that gap would be circumventing
      an explicit policy, not solving a technical problem — declined
      regardless of how legitimate the underlying research use is.**
      Confirmed the actual right path instead: Steve reading BOM's site
      himself and entering figures into `manual_rainfall_data.toml`.
- [x] **All 12 stations tested directly against SILO (2026-09-21),
      using Steve's confirmed list of true published stations —
      genuinely good news: only 2 need manual entry, not 5.** Chinchilla
      (42078), Dalby (41522), Dysart (35109), Miles (42023), Narrabri
      (54038), Roma (43091), **Tara (42086)**, Toowoomba (41529),
      Wallumbilla (43043), Wandoan (35029) are ALL confirmed genuinely
      in SILO — the earlier Dalby/Tara substitutions turn out to have
      been unnecessary; both true stations were available in SILO all
      along (or added to SILO's PPD since the original substitution
      decision, years ago). Only **Goondiwindi (41507)** and
      **Moranbah (34035)** are confirmed NOT in SILO — genuinely need
      `manual_rainfall_data.toml` entries. `towns.toml` corrected for
      all of Tara/Goondiwindi/Wandoan/Dalby/Moranbah to match Steve's
      confirmed list exactly.
      **Still open, Goondiwindi specifically:** the published record
      shows a real station transition — "New Kildonan 041507 / WTP
      (2020→)" — meaning 2020-onward figures should come from the WTP
      station specifically, not New Kildonan continued. WTP's own
      station number isn't known yet. Since 41507 isn't in SILO either
      way, this doesn't block automation (manual entry is needed
      regardless), but whoever enters Goondiwindi's manual data needs
      to know which station's numbers to actually read for 2020+.
      **Wandoan's "TM (2021→)" transition is real per the published
      record too, but moot for automation** — 35029 (the pre-2021
      station) is confirmed IN SILO, so the fetcher works fine either
      way; the transition would only matter for exact continuity if
      SILO's own 35029 data happens to reflect the same underlying
      station change already (unconfirmed either way, not urgent).
- [x] **CONFIRMED LIVE (2026-09-21): 15/17 towns clean via SILO alone,
      no regressions.** Real run against the corrected `towns.toml`:
      Dalby (538.4mm) and Tara (475.8mm) both succeed directly through
      their true published stations, exactly as the direct SILO test
      predicted. The no-substitution policy held correctly even when
      tested under real pressure — SILO's own name-search surfaced
      `34038 MORANBAH WATER TREATMENT PLANT` as an available
      alternative for Moranbah, and the fetcher correctly did NOT
      silently substitute it, just logged it as information and still
      failed cleanly with the BOM link. Yarram briefly regressed back
      to the old 85151 in one run (likely an unsynced local edit,
      nothing wrong with the design) and was confirmed fixed on
      re-check — 85193 now used correctly. Genuinely done: `bom_rainfall`
      is now fully policy-correct (published-station continuity over
      automated convenience), with exactly two towns (Goondiwindi,
      Moranbah) waiting on manual entries, both understood and
      expected, not bugs.
- [x] **`update_rainfall.py` first real run found and fixed a genuine
      bug (2026-09-22): wrote 2025's data to column AA instead of Z.**
      Root cause confirmed directly against the real files: the
      Exogenous sheet's `used_range` reports the sheet-wide column
      extent (Z/26 in both the 2025 original and 2026 final files) --
      but that's from UNRELATED content elsewhere in the sheet
      (Education/Fuel sections further down), not from the Rainfall
      section's own year headers, which genuinely stop at Y (2024) in
      the pre-update file. Trusting `used_range.last_cell.column` as
      "where the year data ends" was the bug -- fixed to search row 1
      itself for the rightmost real year value and append directly
      after that, ignoring whatever else the sheet's used_range
      reports. Tested against a mock reproducing the exact real
      scenario (real years through Y, misleading used_range at Z) --
      confirms it now correctly creates 2025 at Z, not AA. Worth being
      aware `base.py`'s Population-sheet version uses the same
      used_range-trusting pattern -- hasn't caused a problem there so
      far (every Population write this session landed exactly where
      expected), so not touched without real evidence of a problem,
      but the same failure mode could in principle recur if that sheet
      ever develops similar unrelated-content pollution.
- [x] **"Historic Average" built and tested against real BOM data
      (2026-09-22).** Steve found BOM's "Climate Averages" tables
      (`bom.gov.au/climate/averages/tables/cw_{station}.shtml`) —
      confirmed genuinely accessible (independently verified via both
      `web_fetch` and a raw `curl` request, real HTTP 200) — a
      completely different BOM product from the interactive Climate
      Data Online portal that blocks automated access. This settled
      the definition question directly: "Historic Average" is BOM's
      own official Mean Annual Rainfall figure, not a self-computed
      rolling mean — confirmed by Steve's own manual reading for
      Moranbah (597.3mm), which also explained the existing workbook's
      unexplained 592.8mm (likely computed differently, from the
      substitute station or a different date range).
      `fetch_bom_rainfall.py` now fetches this automatically per
      station and stores it in the cache alongside everything else.
      **Real bug found and fixed during testing**: the real page has
      18 cells per row (including 2 trailing empty plot/map icon
      cells) — an isolated test snippet only had 16, so indexing from
      the END of the row (`cells[-3]`) silently grabbed the date-range
      cell instead of Annual on the first live test. Fixed to index
      from the START instead, which is stable regardless of trailing
      decorative cells. Re-tested against the real live page after the
      fix — exact match (597.3mm) confirmed.
      `update_rainfall.py` writes it with a genuinely different
      pattern from total/summer/winter: the SAME value across every
      year column (confirmed that's how this row is actually used —
      a flat constant for charting), with a 20%-difference sanity
      check against whatever's already there before overwriting, and
      per Steve's explicit instruction — when the page genuinely isn't
      available for a station (confirmed real, e.g. Chinchilla
      doesn't have one), the existing value is left untouched with a
      clear note, never blanked or guessed. Tested end-to-end:
      real Moranbah value fetches correctly, a wildly different value
      correctly gets flagged rather than silently overwriting, and an
      empty row writes cleanly with nothing to compare against.
- [x] **Real second bug found and fixed on the first live run
      (2026-09-22): "kept existing value" left a genuine gap at 2025
      specifically.** First live test against a real test file: 39/39
      total/summer/winter writes succeeded, but Historic Average came
      back completely blank for 2025 across every town, even ones
      with real, complete history through 2024. Root cause: a fresh
      BOM fetch mostly 404s (confirmed genuinely real, independently
      verified for Dalby/Toowoomba/Narrabri/Roma — BOM's Climate
      Averages compiled-table product has much narrower station
      coverage than hoped; Moranbah appears to be the exception among
      this town list, not the norm), and the old fallback logic did
      NOTHING when the fetch failed — but "keep the previous average"
      per Steve's original instruction actually means writing that old
      value into the newly-created 2025 column too, not leaving it
      untouched while every other column has data. New
      `_carry_forward_historic_average()` does this properly: reads
      whatever's already in the row (a flat constant, so any populated
      cell works) and writes it into just the new year's column,
      leaving the rest of the row alone. Tested against both the exact
      real bug scenario (real history through 2024, genuinely blank
      2025) and the genuinely-empty case (a town with no history at
      all yet, like Moranbah) — both confirmed correct.
- [x] **Declined a browser-automation (Selenium) route Steve found
      partial success with, and worth recording why clearly — not
      rejected for not working, rejected on principle (2026-09-21).**
      Using Selenium to drive a real browser through BOM's climate-data
      form and download files is still automated access to a service
      that's explicitly stated it doesn't want that (the same policy
      that blocked the CDO endpoint directly) — it just uses a page
      their bot-detection hasn't caught up to yet. Building around
      exploiting that gap would be circumventing an explicit access
      policy, not solving a technical problem, regardless of how
      legitimate the underlying research use is. Declined to help
      extend, debug, or integrate this into the pipeline. The actually-
      correct path — Steve reading BOM's site himself and entering
      figures into `manual_rainfall_data.toml` — is what's actually
      being used, and it's already tested and working.
- [x] **Confirmed real reason for the station-number-in-label
      convention (2026-09-21).** Directly diffed the 2025 original vs
      2026 final Exogenous sheets: 2025's Rainfall row labels have NO
      station numbers at all ("Harewood", "Dalby Airport", "New
      Kildonan/ WTP (2020→)") — 2026 adds them ("Harewood 042078",
      "Dalby Airport 041522", "New Kildonan 041507/ WTP (2020→)").
      Steve did this deliberately this year, specifically because
      chasing down the real station numbers (this whole rainfall
      thread) required real digging — embedding them in the label means
      nobody has to repeat that hunt. Confirmed this is exactly the
      convention `update_rainfall.py`'s row-finder already assumes
      (built and tested against the real 2026 structure) — matching by
      station number as a substring, which only works if the number is
      actually in the label. That's also the right safety net for "a
      future editor forgets to include the number" (Steve's own
      phrasing): the row-finder doesn't guess or fuzzy-match if the
      number's missing, it just fails loudly with a clear error, same
      as a genuinely-missing town.
- [ ] **New-town case explicitly deferred by Steve, not being solved
      now.** Current behavior: a brand new town (no existing Exogenous
      row at all) would just fail with "could not find station number
      X" the same as a bad/missing label — consistent with the "doesn't
      guess or create a row for you" principle everywhere else in this
      project, and a reasonable placeholder, but genuinely not designed
      for yet. Whenever this gets picked up: needs a real decision on
      whether the row gets created automatically (and by what template)
      or whether it's always a manual step first.
      SILO's own 35029 data happens to reflect the same underlying
      station change already (unconfirmed either way, not urgent).
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
- [x] **Chinchilla's formula cell — resolved, and confirms the audit
      design is working as intended, not a gap in it.** Checked whether
      anything else in the workbook depends on it: the pattern
      `=(Y-X)/X` is a real, intentional growth-rate calculation, but it
      belongs in row 53 (confirmed present there, consistently, in
      every version of the workbook checked — original 2025, current
      2025, and the 2026 truth document), not row 54 where it was
      flagged — the truth document's row 54 is a plain number, meaning
      whoever built it already cleared exactly this kind of stray
      formula in this exact spot. Nothing else references row 54
      specifically. **Steve's decision: delete it manually, re-run** —
      and confirmed this is the intended workflow going forward, not a
      one-off fix: audit flags an artifact, a human decides and cleans
      the sheet, re-run. More of this kind of leftover exploratory-
      analysis crud is expected across Income/Crime/Housing and the
      other sheets — connects directly to the existing "generalize
      row-finding to other sheets" item below, not a surprise when it
      turns up there too.
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
## Income indicator — wiring built (2026-09-23)

- [x] **`update_income.py` built and tested (2026-09-23).** Both ATO
      fetchers (`fetch_income.py` Table 8, `fetch_income_table6.py`
      Table 6) already worked at the fetch level (17/17, per project
      summary) -- this wires their output into the real Income sheet.
      Confirmed real sheet structure first: fiscal-year strings
      directly in row 1 ("2000/01"), no separate calendar-year row 2
      like Population has; 12-row town blocks (name, postcode, 10
      indicator rows) confirmed consistent across all 12 towns this
      sheet covers (same 12 as Rainfall -- VIC towns aren't covered
      here either); a state-level benchmark section below row ~147,
      out of scope; 4 "Individual ABN..." rows confirmed consistently
      empty for every town -- genuinely unsourced, not something
      either fetcher touches.
      **Real structural incompatibility found before writing any
      code, not after**: base.py's existing `_find_town_indicator_row`
      would have broken here -- it treats an empty column B as "left
      this town's block", which is exactly the postcode row's shape,
      so it would have incorrectly exited Chinchilla's block one row
      after entering it. Built a dedicated row-finder instead (same
      reasoning as Exogenous needing its own), with a real safety
      feature base.py's version doesn't have: cross-checks the
      postcode row against towns.toml before trusting a name match,
      not just trusting the town name string alone.
      Indicator mapping confirmed against real row labels: Table 8's
      full series -> "Average Taxable Income or Loss (all
      individuals)" (gets the ground-truth historical audit, full
      series available); Table 6's three latest-year-only values ->
      "...(taxable individuals)", "No. wage and salary earners",
      "Total wage and salary earnings" (shape-based audit only, like
      NRW's UCL writes).
      Tested (7 checks, all pass): town+postcode found correctly,
      wrong postcode correctly raises rather than silently
      proceeding, indicator found within the correct bounded window,
      a label that only exists in the NEXT town's block correctly
      fails to match (critical boundary test), fiscal-year column
      finding for both an existing year and a new one.
      **NOT yet tested against a live Excel instance** — next step is
      a real run against a throwaway test copy, same as every other
      wiring script's first live test.

## Income wiring — real bug found and fixed on first live run (2026-09-23)

- [x] **Serious bug: every write call created a NEW duplicate year
      column instead of reusing the one from the previous call.**
      First live run corrupted row 1 with 11 duplicate "2023/24"
      headers (columns Z through AJ) in a single run. Root cause: a
      genuine year-convention mismatch in `_find_year_column`'s
      comparison — `year` is always passed in as the CALENDAR year a
      fiscal year ENDS in (e.g. 2024 for FY "2023/24", matching
      `_fiscal_label`'s own convention), but the search loop compared
      it directly against the STARTING year parsed from each label
      ("2023/24" → 2023) — 2023 never equals 2024, so an already-
      existing column could never be found again, and every single
      write call (up to 17 towns × 4 indicators) created another new
      one. Fixed to parse the label the same way
      `_fy_to_calendar_year` does, so both sides of the comparison use
      one consistent convention. Tested against the exact real
      scenario: 5 repeated calls for the same target year now all
      correctly return the same column, with the sheet's real
      used_range pollution (AC-AF's growth-rate formula crud inflating
      it to column AG) reproduced in the mock too, confirmed not to
      matter — row 1 search is genuinely immune to it, was never
      actually the cause despite looking suspicious at first.
      Also directly confirmed (separately, since this bug happened to
      never trigger it) that the existing audit system's formula
      detection would correctly catch and block a write landing on
      one of the AC-AF-style formula cells, same as it did for
      Chinchilla's Population row — this protection was already
      correctly in place, just never exercised by the buggy version.
      **`test-copy.xlsx` is corrupted from this bug and should be
      discarded, not manually repaired** — unclear which data rows
      also got phantom writes scattered across the duplicate columns.
      Start fresh from the original once the fix is confirmed.

## Income methodology correction + ABN indicators (2026-09-23)

- [x] **Real methodology bug found via the delta check Steve asked
      for, and fixed.** Comparing test-copy.xlsx against the 2026
      reference file: "Average Taxable Income or Loss (all
      individuals)" for Chinchilla 2023/24 showed $72,289 (mine, from
      Table 8's own pre-published average) vs $73,581 (reference) —
      a real ~1.8% discrepancy, not rounding. Steve's `Notes_DD.docx`
      documents the actual intended methodology: this row should come
      from Table 6B (`Taxable income or loss $ / no.`), a completely
      different ATO product from Table 8. Verified directly against
      the real downloaded Table 6 file: the documented Table 6B
      calculation gives exactly $73,581 for Chinchilla — exact match.
      "Taxable individuals" was already correct (Table 6A, confirmed
      unchanged, still exact match at $91,431).
      Fixed: `fetch_income_table6.py` rebuilt to also extract Table
      6B's `taxable_no`/`taxable_income` (confirmed real column
      indices 4/5) and compute `avg_income_all`. `update_income.py`
      changed so Table 6B's value takes over entirely for whichever
      year it covers — Table 8's (confirmed less accurate) figure for
      that same year is never written at all, avoiding any same-run
      conflict. Table 8 still backfills older years Table 6's
      single-year-per-release snapshot doesn't reach.
- [x] **Six "Individual ABN..." rows built — previously entirely
      unsourced in the workbook.** Per Notes_DD.docx's documented
      mapping (the ATO file doesn't use the literal words "Individual
      ABN"): NPP/PP/Total business income fields from Table 6B,
      confirmed real column indices 132-137. Sanity-checked against
      real data: PP + NPP income exactly equals Total income for
      Chinchilla ($41,377,276 = $41,377,276), supporting these are
      genuinely the right fields.
      **One assumption worth Steve confirming, not just asserting**:
      the notes don't explicitly say Table 6A vs 6B for these six
      fields specifically. Used 6B (unfiltered/all-individuals),
      consistent with how wages already comes from 6B and none of
      the six ABN labels carry a "(taxable individuals)" qualifier
      the way the two average-income rows do — a reasonable inference
      from the documented pattern, not something explicitly stated.
- [x] **Corrected column indices confirmed directly against the real
      2023-24 file, not assumed to carry forward from the 2022-23
      file's indices the old docstring cited** — downloaded and
      inspected the real file's full header row for both 6A and 6B
      rather than trusting last year's positions still hold.

## Income wiring — real bugs found on second live run, plus a genuine methodology question (2026-09-23)

- [x] **Real bug: the "prevent double-write" logic never actually
      worked, due to a slug-vs-display-name key mismatch.** `t6_files`
      was keyed by filename slug ("chinchilla") but looked up by
      display name from Table 8's JSON ("Chinchilla") — guaranteed
      mismatch, so `skip_this_year` was silently always False. Table 8
      wrote into every "all individuals" cell first, every time, and
      only the audit's 1% tolerance masked it for towns where the two
      sources' values happened to land close together (Dysart,
      Moranbah, Toowoomba) — everywhere else (Chinchilla, Dalby,
      Goondiwindi, Miles, Roma, Tara, Wallumbilla, Wandoan) it surfaced
      as a confusing "cell already contains X" flag against a value
      this same run had just written moments earlier. Fixed by keying
      `t6_files` by display name (read from each file's own "town"
      field) instead of the filename slug. Directly confirmed via a
      reproduction: the old approach's lookup returns None for every
      real case, the fixed one matches correctly.
- [x] **Real, confirmed label variance: Narrabri's block uses
      different exact wording for three rows** — "Average taxable
      income (all)" not "...or Loss (all individuals)", same for
      "(taxable)", and "Total wage & salary earnings" (ampersand) not
      "...and salary earnings". Same shape of issue as Rainfall's
      Narrabri quirk. Fixed via an explicit, documented
      `ALTERNATE_LABELS` dict — the row-finder still requires an exact
      match against one of a specific, recorded set of strings, never
      a fuzzy "close enough" guess. Tested against a mock reproducing
      Narrabri's exact real labels — all four indicators (including
      the one with identical text, confirming nothing regressed)
      resolve correctly.
- [x] **Confirmed genuinely correct, not bugs**: Toowoomba (Central)/
      (Harlaxton)/(West) and Shepparton/Yarram all correctly fail
      "could not find town" — Income has exactly 12 town rows, only
      the combined "Toowoomba", confirmed directly against the
      reference file. Nothing to fix.
- [x] **Confirmed safe, via direct reference-file comparison**: Dysart
      ABN PP no. (fetched 11, reference Z=11), Goondiwindi earners_no
      (fetched 3,657, reference 3,657), Moranbah ABN PP no. (fetched
      30, reference 30) — all three flagged jumps are genuine, correct
      changes, not errors. Good candidates for
      `verified_overrides.toml` entries if Steve wants them written
      through rather than re-flagged on every future run.
- [ ] **Genuine methodology question, needs Steve's decision, not
      something to resolve unilaterally.** Toowoomba's four flagged
      figures (earners_no, all three ABN totals-no) are NOT just
      flagged-but-correct like the above — they're genuinely
      different from the reference, and precisely explained:
      `towns.toml` has `postcodes = ["4350", "4352"]` for Toowoomba;
      the fetcher correctly sums across both (confirmed: postcode 4350
      alone = 58,143, exactly the reference's figure; 4352 alone =
      16,954; 58,143 + 16,954 = 75,097, exactly the fetched combined
      total). The real question: is 4350+4352 the intended scope for
      Income (and potentially every other indicator that aggregates
      across Toowoomba's postcodes), or did 4352 get added to
      `towns.toml` after the reference file's income figures were last
      computed with a narrower 4350-only scope? Whichever way this
      goes could affect more than just Income.

## Income: state benchmarks + write-with-flag highlighting built (2026-09-23)

- [x] **State benchmark calculation built and verified against real
      data.** Per Notes_DD.docx's documented process: NOT an average
      of postcode averages -- sum the raw dollar/count fields across
      every row belonging to that state, then divide. Verified
      directly against the real downloaded Table 6 file before
      building anything: computed QLD/NSW all/taxable all four exactly
      match the 2026 reference workbook's existing figures (75,890 /
      90,846 / 83,306 / 100,533). Confirmed real benchmark-section
      structure: a 3-row block per state (name row + 2 indicator rows,
      no postcode row like town blocks have), with QLD and NSW using
      genuinely different exact label wording for their own two rows
      (confirmed directly, including "Queensland " with a trailing
      space). New dedicated row-finder (`_find_state_row`/
      `_find_state_indicator_row`) rather than reusing the town-block
      one, tested against a mock matching the real structure exactly
      (5/5 checks pass).
- [x] **Real bug, crashed on first live run, fixed immediately**: the
      glob pattern finding per-town Table 6 cache files
      (`*_income_t6.json`) also matched the new benchmark cache file
      (`benchmark_income_t6.json`), which has no `"town"` key (it has
      `"benchmarks"` instead by design) -- crashed with `KeyError:
      'town'` the first time a benchmark file actually existed to be
      swept in. Fixed by explicitly excluding that filename from the
      per-town glob. Confirmed fixed via a reproduction with both file
      types present together.
- [x] **New capability: flagged writes now visually highlight instead
      of silently not writing, per Steve's request** — "write it in
      bold red so it's easier to check directly in the sheet, rather
      than only refusing to write and requiring a separate log check."
      Added `WriteAuditReport.should_write_with_flag` to `audit.py`
      (shared, so this could extend to Population/Rainfall too, not
      just Income) -- deliberately scoped: a series-shape or
      historical-discrepancy concern now writes-with-flag, but a
      CELL-level block (a formula, unexpected existing content) still
      hard-blocks regardless -- overwriting structural content is a
      different, higher-stakes risk than a suspicious number, and that
      boundary must never soften. Tested exhaustively at the audit.py
      level (5/5 scenarios: clean, formula-blocked, unexpected-
      content-blocked even with a clean series, override-applied).
      `update_income.py`'s writer applies bold red font on a flagged
      write, and explicitly resets to plain formatting on a clean
      write (so a once-flagged, later-confirmed cell doesn't stay red
      forever). **Confirmed working at the write-decision and value
      level via a corrected mock** (the value genuinely gets written
      into a new cell despite the flag, not left blank) -- but the
      actual font-formatting assertion in that same test run couldn't
      be confirmed due to a separate mock limitation (the test mock
      doesn't persist font state across repeated `sheet.cells()`
      calls the way it does for `.value`), not a sign of a problem in
      the real code path. Worth a real (non-mocked) xlwings run to
      fully confirm the visual formatting before treating this as
      completely verified.

- [x] **Confirmed exactly the risk flagged earlier — real xlwings run
      crashed on `cell.font.color = None` (2026-09-23).** `TypeError:
      'NoneType' object is not subscriptable` inside xlwings'
      `rgb_to_int` -- confirms xlwings' real Font.color setter has no
      "reset to automatic" pathway via None, unlike what a mock
      assumed. Fixed to explicit black `(0, 0, 0)`, the standard
      default text colour. This is exactly why the earlier note said
      the formatting behaviour needed a real (non-mocked) run to fully
      confirm -- the mock genuinely could not have caught this.
      One line, isolated, well-understood fix -- worth a re-run to
      confirm clean end-to-end now, including whether the bold-red
      highlighting actually displays correctly for the genuinely
      flagged cells (Dysart, Goondiwindi, Moranbah, Toowoomba).

- [x] **CONFIRMED FULLY WORKING, real live run, zero crashes
      (2026-09-23).** 124 written (up from 113 -- now including the 4
      state benchmark writes), 27 correctly flagged-but-written in
      bold red (all previously confirmed correct via direct reference
      comparison), including Steve confirming the highlighting is
      visually working in real Excel. Income is genuinely done: both
      fetchers, the corrected all-individuals methodology, the six ABN
      indicators, QLD/NSW state benchmarks, Narrabri's label variance,
      the double-write bug, and the write-with-flag highlighting
      feature are all built, tested, and now confirmed live. Only
      remaining open item is the Toowoomba postcode-scope question
      (4350 alone vs 4350+4352) -- a real decision for Steve, not a
      bug, and not blocking anything.
## Crime indicator — two real, significant bugs found and fixed in the existing fetcher (2026-09-24)

- [x] **Major finding: the existing `fetch_crime_qps.py` (previously
      marked "14/14 QLD towns working") has NEVER produced correct
      values.** "Working" only ever meant "runs without crashing" --
      nobody had validated the actual numbers against the real
      reference workbook until now. Two real, confirmed bugs:
      1. Annual aggregation used `statistics.mean()` of the 12 monthly
         rates. An annual rate should be the SUM across the year (a
         rate PER YEAR, not an average monthly rate) -- confirmed via
         an almost-exact 12.024x ratio between old output and real
         workbook values, across four independent indicators
         (drug/good_order/theft/traffic). Every crime figure this
         fetcher ever produced, for all 14 towns, was wrong by roughly
         a factor of 12.
      2. "Total offences (person, property, other)" summed ALL 8 QPS
         summary categories, when the row's own name says literally
         "(person, property, other)" -- confirmed directly: summing
         just those three raw columns (with the sum-not-mean fix
         applied) matches the real value to within 0.2%; summing all 8
         does not come close.
      Both fixes verified together against all 11 individual real 2001
      values for Chinchilla plus Total -- every single one now matches
      the real workbook within a consistent, uniform 0.20% (the same
      tiny residual across every indicator, strongly suggesting a
      minor QPS-side population-denominator rounding quirk, not a
      remaining methodology error on our side).
      Also EXTENDED to cover all 12 rows the real sheet actually has
      per town (confirmed via direct inspection) -- the old version
      only produced 5 (all/drug/good_order/theft/traffic). Added:
      breach_dv, offences_property, offences_person, other_offences,
      prostitution, unlawful_entry, weapons -- confirmed via the raw
      column order these are genuinely independent categories, not
      double-counted sub-items of Total or of each other.

- [x] **Third real bug found and fixed, from checking the live run's
      "latest (2026)" output critically rather than accepting it at
      face value.** 2026 is the in-progress current year (today is
      2026-09-24) -- confirmed directly it only had 8 of 12 months in
      the raw CSV. Combined with the mean-to-sum fix, this meant the
      fetcher would silently write a partial-year sum as if it were a
      genuine annual total, misleadingly looking like a huge crime
      drop once compared against a real full year. Fixed: skip any
      division-year with fewer than 12 months present, same "skip
      incomplete years" pattern already proven correct in
      `fetch_bom_rainfall.py`. Tested against real data: 2026 now
      correctly excluded, 2025 correctly becomes the latest complete
      year, and the earlier fix's 2001 values are unaffected.

## Crime wiring script built (2026-09-24)

- [x] **Found and fixed: `apply_write_formatting`/`describe_write_outcome`
      (promoted to `audit.py` during the Business write-with-flag work)
      were never actually pushed to GitHub** -- `should_write_with_flag`
      and the `summary_line` structural fix were committed correctly,
      but the two shared functions added right after them, in the same
      sitting, apparently weren't. Found immediately (an ImportError)
      when `update_crime.py` tried to use them. Restored and verified
      both with syntax AND a real runtime import this time (not just
      `ast.parse`, learning from the earlier lesson that syntax
      validity doesn't guarantee correctness).
- [x] **`update_crime.py` built and tested for the 11 confirmed QLD
      towns.** Confirmed real structure: plain calendar-year integers
      in row 1 from column C (not fiscal strings); 13-row town blocks
      (name + 12 indicator rows) in a fixed order confirmed identical
      across every town checked. Tested (6 checks): town/indicator
      row-finding, correctly bounded search (doesn't spill into the
      next town), existing-year lookup, and new-year column creation.
      **Deliberately out of scope, confirmed real, not oversights:**
      Narrabri/NSW use a completely different 5-category structure
      (Assault, Malicious damage to property, Other offences, Other
      offences against the person, Robbery) -- genuinely need NSW
      BOCSAR as a data source, not QPS, not built yet. The Queensland
      state benchmark uses the same 12-category structure as towns,
      but QPS's data has no native statewide division -- would need
      its own cross-division aggregation, similar to Income's QLD/NSW
      benchmark work -- real, separate follow-up.
      **NOT yet tested against a live Excel instance.**

- [x] **Real bug found and fixed on the first live run: `isinstance(label, int)` silently rejected every real year value.** xlwings returns whole-number cells as floats (2001.0) via Excel's COM interface -- confirmed real, openpyxl (used only for inspection, never the live write path) happens to preserve int type but xlwings doesn't, so a mock built with plain ints never would have caught this. 0/168 written, every single row failed with "could not find any year values." Fixed in all three affected spots: broadened to `isinstance(label, (int, float))`, compared/stored via `int(label)`. Retested with a mock using actual floats this time (matching the real scenario) -- year-column finding, existing-series reading, and dict-key typing all confirmed correct.
- [x] Confirmed expected, not a bug: Toowoomba (Central)/(Harlaxton)/(West) correctly fail "could not find town" -- Crime has one combined "Toowoomba" row only, same pattern as Business and Income.