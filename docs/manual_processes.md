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
listed in `docs/data_sources/qgso_housing.md`. If QRSIS changes shape and the
fetcher breaks before a fix lands, the manual fallback is:

1. Go to http://www.qgso.qld.gov.au/products/tables/qld-regional-database/index.php
2. [NEEDS INPUT — the exact wizard steps: which theme, which region type,
   which date format, which series names to select]
3. Save the resulting export as `cache/qgso_and_bom_{YEAR}.xlsx`
4. Re-run `python run_update.py --only qgso_housing` — the fetcher should
   detect and read the manually-placed file
   [NEEDS INPUT — confirm the fetcher actually has this fallback-read
   behaviour implemented, or whether this is aspirational]

---

_Add further manual fallback processes below as they come up — same format:
numbered steps, exact URLs, exact filenames, and an explicit `[NEEDS INPUT]`
marker for anything not yet confirmed rather than a guess._