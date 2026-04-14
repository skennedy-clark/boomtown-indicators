# boomtown-indicators

Data pipeline for the UQ Centre for Natural Gas regional indicators project.
Fetches, processes and outputs social and economic indicator data for Queensland,
NSW and Victorian towns affected by CSG/resource development.

---

## Setup

```bash
git clone https://github.com/skennedy-clark/boomtown-indicators.git
cd boomtown-indicators
python -m venv venv
venv\Scripts\activate        # Windows
source venv/bin/activate     # *nix
pip install -r requirements.txt
```

---

## Commands

```bash
# Validate towns.toml config
python run_update.py --validate

# Run all fetchers
python run_update.py

# Run one fetcher
python run_update.py --only income

# Run one fetcher for specific towns
python run_update.py --only income --towns Roma Dalby

# Clear cache and re-download everything
python run_update.py --force

# Clear cache for one fetcher only
python run_update.py --only income --force

# Show what's in the cache
python run_update.py --list-cache

# Dry run — show what would run without fetching
python run_update.py --dry-run

# Skip a fetcher
python run_update.py --skip crime
```

---

## Websites

- https://gas-energy.centre.uq.edu.au/resources/tools
- https://boomtown-indicators.org/compare

---

## Fetchers

| Name | Source | Method | States | Status |
|------|--------|--------|--------|--------|
| `income` | ATO Taxation Statistics Table 8 (data.gov.au) | CKAN API | All | ✓ Working |
| `income_table6` | ATO Taxation Statistics Table 6 (data.gov.au) | CKAN API | All | ✓ Working |
| `population_ucl` | QGSO UCL Resident Population | Direct CSV download | QLD | ✓ Working |
| `population_erp` | ABS Estimated Resident Population (NSW/VIC) | ABS API | NSW, VIC | TODO |
| `population_nrw` | QGSO Surat/Bowen Basin Population Reports | QGSO web | QLD | TODO |
| `salm_unemployment` | ABS Small Area Labour Markets (SALM) | data.gov.au | QLD (14) | ✓ Working — NSW/VIC below threshold |
| `qgso_housing` | QGSO Regional Database (QRSIS) | QRSIS web scraper | QLD | ✓ Working — 14 towns |
| `housing_nsw` | NSW Valuer General / FACS rent tables | Web | NSW | TODO |
| `crime_qps` | QPS Reported Offences Rates | QPS S3 bucket CSV | QLD | ✓ Working |
| `crime_nsw` | BOCSAR LGA offences | BOCSAR web | NSW | TODO |
| `bom_rainfall` | SILO Patched Point Dataset | SILO API | All | ✓ Working — 13/14; Yarram parked |
| `business` | ABS Counts of Businesses by Turnover (8165) | ABS API | All | TODO |
| `fuel` | RACQ Annual Average Fuel Prices | RACQ web | QLD | TODO |
| `schools` | ACARA School Profile | ACARA data portal | All | TODO |

---

## Towns

Configured in `towns.toml`. Currently tracking:

**QLD:** Roma, Chinchilla, Dalby, Miles, Tara, Wandoan, Wallumbilla,
Goondiwindi, Moranbah, Dysart, Toowoomba (+ 3 sub-areas: Central, Harlaxton, West)

**NSW:** Narrabri

**VIC:** Shepparton, Yarram *(test towns)*

---

## Output

Processed data lands in `cache/{source}/` as JSON.
Final CSVs (one per indicator per town) will go in `output/{town}/`.

---

## Known issues / limitations

### Housing: Roma and Miles share SA2 boundaries
Roma and Wallumbilla share `qgso_sa2 = 307011176`. Miles and Wandoan share
`qgso_sa2 = 307011175`. All four towns correctly receive data from the shared SA2,
so Roma and Wallumbilla show identical sales/rent figures, as do Miles and Wandoan.
This is a QRSIS data granularity limitation, not a pipeline bug.

Building approvals for Toowoomba use the LGA boundary (not SA2), per previous booklets.

### Wallumbilla SA2 code probably wrong
`sa2_code = 307011178` (Tara SA2) in `towns.toml` is likely incorrect. The correct
code is probably 307011177 (Roma Surrounds). Do not use SA2-level data for Wallumbilla
until verified against ABS boundary files.

### Tara rainfall: synthetic data flag
SILO station 41099 has >10% synthetic data in 2004, 2005, 2010, 2012, 2013.
Years are flagged in JSON output. Data is usable but note in booklet.

### Yarram rainfall: no SILO station
Station 85151 is not in the SILO dataset. Parked pending manual investigation.

### NSW/VIC unemployment
Narrabri, Shepparton and Yarram are absent from the DEWR SALM smoothed SA2
dataset (below statistical reliability threshold). No unemployment data available
for these towns. LGA-level fallback not yet implemented.

### SALM URL changes quarterly
The DEWR SALM download URL contains the quarter and a file ID that change each
release. Update `FALLBACK_CSV_URL` in `fetch_salm_unemployment.py` each quarter.

---

## TODO

### Fetchers still needed
- [ ] `fetch_population_erp.py` — ABS ERP for NSW/VIC towns (Narrabri, Shepparton, Yarram)
  - ABS API: https://api.data.abs.gov.au/ (ERP dataset, filter by SA2/LGA code)
- [ ] `fetch_population_nrw.py` — QGSO Surat/Bowen Basin non-resident worker population
  - Surat: https://www.qgso.qld.gov.au/statistics/theme/population/population-estimates/surat-basin
  - Bowen: https://www.qgso.qld.gov.au/statistics/theme/population/population-estimates/bowen-basin
- [ ] `fetch_housing_nsw.py` — Narrabri housing (NSW Valuer General + FACS rent + ABS 8731)
- [ ] `fetch_crime_nsw.py` — BOCSAR LGA offences (Narrabri)
- [ ] `fetch_business.py` — ABS 8165 business counts by SA2
- [ ] `fetch_fuel.py` — RACQ average ULP prices (likely manual — annual PDF)
- [ ] `fetch_schools.py` — ACARA school enrolments by suburb/postcode

### Output
- [ ] `transform/to_csv.py` — wire up building_approvals transformer (18 of 25 CSVs done)

### Maintenance
- [ ] Update SALM fallback URL each quarter (see fetch_salm_unemployment.py)
- [ ] Verify Tara/Goondiwindi/Moranbah BOM station substitutions against previous booklets
  (stations changed from those used in 2020 booklets — continuity not yet confirmed)

### Geography
- [ ] Download ABS ASGS Edition 3 (2021) shapefiles for boundary maps
  - Source: https://www.abs.gov.au/statistics/standards/australian-statistical-geography-standard-asgs-edition-3
- [ ] Verify Wallumbilla SA2 code (307011178 vs 307011177)

### Booklets
- [ ] Booklet template (replaces manual Word/Excel/Acrobat process)
- [ ] Chart generation from CSVs
- [ ] PDF assembly pipeline

### Website
- [ ] Get sFTP credentials for boomtown-indicators.org server
- [ ] Automate CSV upload after successful fetch run

---


---

## Verification checklist — items to check before next publication

### Items to verify technically (pipeline team)

**1. Toowoomba building approvals: financial year vs calendar year aggregation**
Previous booklets (pre-2017) reported approvals as financial years (e.g. FY2016 = Jul 2015–Jun 2016),
switching to calendar years from 2017 onward. The current fetcher aggregates monthly approvals into
calendar years throughout the full series. This creates a methodological discontinuity with the 2018/19
booklet chart. Decide: adopt calendar years throughout (cleaner, consistent with other indicators) or
replicate the financial-year aggregation for historical data to match previous booklets.

**2. Toowoomba rainfall station**
The 2018/19 booklet (Appendix C) explicitly states "Toowoomba Airport Rainfall Station". The current
pipeline uses station 41529 (towns.toml `bom_station`). Verify whether 41529 is the Airport station
or a different Toowoomba station — if different, the rainfall series will not match previous booklets.
Check: https://www.bom.gov.au/climate/data/ → search Toowoomba → compare station names.

**3. Wallumbilla SA2 code**
`sa2_code = 307011178` in towns.toml is the Tara SA2 code — almost certainly wrong for Wallumbilla.
Correct code is likely 307011177 (Roma Surrounds, active from 01/07/2021). Verify against ABS ASGS
boundary files before using any SA2-level data for Wallumbilla. Note: `qgso_sa2 = 307011176` (Roma)
is intentional for QRSIS housing data — that is separate from sa2_code.

**4. BOM station substitutions — continuity with previous booklets**
Three towns changed station due to SILO availability:
- Tara: 42104 (Woodlea) → 41099 (Tara Township)
- Goondiwindi: 41507/41559 → 41038 (Post Office)
- Moranbah: 34035 (Airport) → 34038 (WTP)
Generate charts for each substituted station and compare overlapping years against previous booklet
charts before publishing. If continuity is poor, consider noting the break in the booklet text.

**5. Roma/Wallumbilla and Miles/Wandoan: shared SA2 housing data**
These town pairs share a QRSIS SA2 boundary, so sales count, median price and rent are identical
within each pair. The pipeline correctly applies the same data to both towns in each pair. Verify this
is consistent with how previous booklets handled these towns — check whether the 2020 Roma and
Wallumbilla booklets show identical or different housing figures.

**6. SALM URL — update each quarter**
The DEWR SALM fallback URL is hardcoded to the December 2025 quarter. Must be updated after each
new quarterly release. Check: https://www.dewr.gov.au/employment-research/small-area-labour-markets

---

### Items requiring academic/research team input

**A. Aggregation methodology: approvals financial year vs calendar year**
A decision is needed on whether to maintain the previous booklets' financial-year aggregation for
building approvals (pre-2017), or standardise to calendar years throughout. This affects comparability
with published booklets and is a methodological question, not purely technical.

**B. Toowoomba sub-area booklets: which geography for approvals?**
The 2018/19 booklet shows LGA-level approvals (~800/yr) for the single "Toowoomba" booklet. If
separate booklets are produced for Toowoomba Central, Harlaxton and West, should those show:
(a) the full LGA figure (same number, different booklet), or
(b) the SA2-level figure for each suburb (much smaller: Central ~10/yr, West ~150/yr, Harlaxton ~10/yr)?
Option (b) matches the SA2 boundary noted for sales/rent but SA2-level approvals are very sparse
for Central and Harlaxton.

**C. Wallumbilla housing data**
Wallumbilla's QRSIS housing data uses Roma's SA2 (307011176), meaning Wallumbilla and Roma show
identical sales/rent/approvals figures. Is this the correct approach, or should Wallumbilla housing be
flagged as unavailable (data is for the Roma SA2 area, not specifically Wallumbilla town)?

**D. Goondiwindi: moved from SA3/30703 to SA3/30701 in ASGS 2021**
Goondiwindi was reclassified in the 2021 restructure. If the academic team is comparing across
different booklet years using SA3-level indicators, they should be aware that Goondiwindi's SA3
changed. This may affect any regional-level analysis.

**E. Narrabri, Shepparton, Yarram: no unemployment data**
These three towns are below DEWR's statistical reliability threshold and have no SA2-level SALM data.
Confirm with the research team whether these towns should show no unemployment chart, use a
broader LGA-level figure (if obtainable), or use a different source (e.g. Census-based estimates).

**F. Benchmark geography for Toowoomba**
The 2018/19 booklet uses Brisbane as the housing benchmark. Confirm whether Brisbane remains the
appropriate benchmark for the 2025/26 booklets, or whether Queensland-wide or another benchmark
is preferred for housing indicators.

## Data sources — detailed notes

### QGSO Housing (QRSIS) — implementation notes

The QGSO Regional Database uses an Oracle PL/SQL Web Toolkit (QRSIS) wizard.

**Critical implementation details:**

1. **POST encoding** — Oracle reads `p_names`/`p_values` as interleaved parallel
   arrays in submission order. All POSTs use the `_q()` helper (list of tuples).
   Never use a dict — dicts group all p_names first then all p_values, which Oracle misreads.

2. **HTML parsing** — QRSIS serves unclosed `<OPTION>` tags. Use `lxml` parser;
   `html.parser` merges all options into one string.

3. **udqctl_id extraction** — Redirect URL format: `?p_names=udqctl_id&p_values=3908`
   (not standard `?udqctl_id=3908`).

4. **Sales and rent** — SA2 geography, Quarterly, `date_format=Y1`, no concorded field.

5. **Building approvals** — Mixed: SA2 for most towns, `LGA/36910` for Toowoomba.
   Monthly, `date_format=M1`, `p_concorded_data=Y` (checkbox checked by default —
   must send this or the region list returns empty).

6. **Exact series names** (case-sensitive, from live QRSIS):
   - `"Detached dwelling: number of sales (Number)"`
   - `"Detached dwelling: median sale price ($)"`
   - `"House - 3 bedrooms - median rent of lodgements ($/week)"`
   - `"Residential dwelling units (Private); New Houses (Number)"`

**Source:** http://www.qgso.qld.gov.au/products/tables/qld-regional-database/index.php

---

### QPS Crime Statistics (Queensland)

**Direct download (no auth, updated monthly):**
```
https://open-crime-data.s3-ap-southeast-2.amazonaws.com/Crime%20Statistics/division_Reported_Offences_Rates.csv
```

Rates per 100,000 population, monthly. Annual value = mean of 12 monthly rates.

| Website CSV | Source column |
|---|---|
| `Crime rate - all offences.csv` | Sum of all offence columns |
| `Drug offences.csv` | `Drug Offences` |
| `Good order offences.csv` | `Good Order Offences` |
| `Theft.csv` | `Other Theft (excl. Unlawful Entry)` |
| `Traffic offences.csv` | `Traffic and Related Offences` |

QPS Division mapping: Roma→Roma, Chinchilla/Dalby→Dalby, Miles→Miles,
Tara→Tara, Wandoan→Wandoan, Wallumbilla→Roma, Goondiwindi→Goondiwindi,
Moranbah→Moranbah, Dysart→Dysart, Toowoomba (all)→Toowoomba.

---

### QGSO UCL Population (Queensland)

**URL pattern:**
```
https://www.qgso.qld.gov.au/issues/{N}/estimated-resident-population-urban-centre-locality-qld-2001-{year}p.csv
```
Issue number `{N}` increments each release. Fetcher scrapes the QGSO page to
find the current URL automatically. Encoding: latin-1. Date format: `JAN01`-style.

---

### ATO Taxation Statistics (National)

**CKAN API:** `https://data.gov.au/data/api/3/action/package_show?id=taxation-statistics-2022-23`

Table 8 = median/average taxable income by postcode. Table 6A/6B = taxable status
split, wage/salary totals. ATO data lags ~18 months — 2022-23 is most recent as
of early 2026. Financial year YYYY-YY maps to calendar year YY+1 in output.

---

### ABS SALM Unemployment

**Source:** https://www.dewr.gov.au/employment-research/small-area-labour-markets

Quarterly smoothed SA2 CSV. URL contains the release quarter and a file ID that
changes each release. Update `FALLBACK_CSV_URL` in `fetch_salm_unemployment.py`
each quarter after DEWR publishes the new release.

---

### BOM Rainfall (SILO Patched Point Dataset)

**API:**
```
https://www.longpaddock.qld.gov.au/cgi-bin/silo/PatchedPointDataset.php
  ?format=csv&comment=R&station={N}&start=YYYYMMDD&finish=YYYYMMDD&username={email}
```

Source codes: `0` = target station, `25` = nearby station, `15` = synthetic.
Years with >10% code-15 days are flagged in JSON output.

**Station substitutions** (previous booklet stations not in SILO):

| Town | Previous | SILO station used | Note |
|---|---|---|---|
| Tara | 42104 (Woodlea) | 41099 (Tara Township) | Continuity unverified |
| Goondiwindi | 41507/41559 | 41038 (Post Office) | Continuity unverified |
| Moranbah | 34035 (Airport) | 34038 (WTP) | Continuity unverified |
| Yarram | 85151 | — | Parked: no SILO station found |

---

## Open questions

**1. SA1 vs SA2 for boundary maps** — SA2 is used for most indicators. Is SA2
sufficient for booklet boundary maps, or do we need finer SA1 granularity?

**2. NSW/VIC equivalent sources** — see fetcher table for gaps. ABS ERP UCL data
is national and could extend the population fetcher to NSW/VIC towns.