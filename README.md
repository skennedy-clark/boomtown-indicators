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

# Linux command line via Docker
docker run -it --rm -v ${PWD}:/workspace -w /workspace ubuntu bash

# View this document in preview
# markdown preview: ctrl + shift + v
```

---

## Websites

- https://gas-energy.centre.uq.edu.au/resources/tools
- https://boomtown-indicators.org/compare

---

## Fetchers

| Name | Source | States | Status |
|------|--------|--------|--------|
| `income` | ATO Taxation Statistics Table 8 (data.gov.au) | All | ✓ Working |
| `population` | ABS Estimated Resident Population | All | TODO |
| `unemployment` | ABS Small Area Labour Markets | All | TODO |
| `housing` | QGSO Regional Database | QLD | TODO |
| `crime` | QPS crime statistics | QLD | TODO |
| `crime_vic` | Victoria Police crime statistics | VIC | TODO |
| `rainfall` | Bureau of Meteorology | All | TODO |

---

## Towns

Configured in `towns.toml`. Currently tracking:

**QLD:** Roma, Chinchilla, Dalby, Miles, Tara, Wandoan, Wallumbilla,
Goondiwindi, Moranbah, Dysart, Toowoomba (+ 3 sub-areas)

**NSW:** Narrabri

**VIC:** Shepparton, Yarram *(test towns)*

Add new towns by adding a `[[towns]]` block to `towns.toml`.

---

## Output

Processed data lands in `cache/{source}/` as JSON.
Final CSVs (one per indicator per town) will go in `output/{town}/`.

---

## TODO

### Fetchers
- [ ] `fetch_population.py` — ABS ERP API, national
- [ ] `fetch_unemployment.py` — ABS Small Area Labour Markets, national
- [ ] `fetch_housing.py` — QGSO Regional Database (QLD), semi-automated download
- [ ] `fetch_crime.py` — QPS crime statistics CSV (QLD)
- [ ] `fetch_crime_vic.py` — Victoria Police crime statistics (VIC)
- [ ] `fetch_rainfall.py` — BOM climate data API, national
- [ ] `transform/to_csv.py` — convert cache JSON to town CSVs in correct format

### Geography
- [ ] Download ABS ASGS Edition 3 (2021) shapefiles for SA2, SA3, LGA boundaries
  - Source: https://www.abs.gov.au/statistics/standards/australian-statistical-geography-standard-asgs-edition-3/jul2021-jun2026/access-and-downloads/digital-boundary-files
  - Files needed: SA2_2021_AUST_GDA2020.zip, LGA_2021_AUST_GDA2020.zip
  - Needed for boundary maps in booklets and website
- [ ] Verify SA2 codes in towns.toml against downloaded shapefiles
- [ ] Generate boundary map images per town (SA2, UCL, postcode overlays)

### Booklets
- [ ] Create booklet template (replaces manual Word/Excel/Acrobat process)
  - Cover page + credits
  - CSG Development Story section (narrative — human input required)
  - Indicator sections: Population, Employment, Income, Housing, Safety & Wellbeing
  - Each section: summary box, data trends, community insights, chart page
  - Appendices: boundary maps, rainfall, project info
- [ ] Chart generation from CSVs (replace Excel charts with matplotlib/plotly)
- [ ] PDF assembly — combine indicator pages into final booklet
- [ ] Page numbering (Roman numerals for intro, Arabic from p.5)
- [ ] File naming and archiving conventions from original process doc

### Website
- [ ] Get sFTP credentials for boomtown-indicators.org server
- [ ] Understand current site structure and how CSVs are consumed
- [ ] Automate CSV upload to server after successful fetch run

---

## Notes

- ATO data lags ~18 months — 2022-23 is the most recent available as of early 2026
- QGSO housing/unemployment data requires semi-manual download (web form navigation)
- Cache is excluded from git — re-run fetchers after a fresh clone
- Toowoomba sub-areas (Central, Harlaxton, West) share postcode 4350 — income data
  is identical across them; differences appear in SA2-based indicators (crime, housing)