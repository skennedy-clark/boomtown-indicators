# Version 3 (draft): Victorian towns, first pass

Victorian towns are added like the New South Wales ones: copied from
Narrabri with `add_town.py`, filled by the annual run. Shepparton and
Yarram are in towns.toml.

    python regional-indicators/transform/add_town.py <workbook> --town Shepparton --like Narrabri
    python regional-indicators/transform/add_town.py <workbook> --town Yarram --like Narrabri

## What is filled

| Sheet | Source | Notes |
|---|---|---|
| Population: SA2 and LGA | ABS (as for NSW) | Shepparton's SA2 is now on 2021 boundaries: Shepparton - North (216031416). The city also covers Shepparton - South East (216031594); a composite of both, as for Toowoomba, is a decision for the research team. |
| Employment: SA2 | SALM (DEWR) | As for every town. |
| Business, Income, Education, Rainfall | ABS, ATO, ACARA, SILO | National sources; no change needed. |
| Crime | Crime Statistics Agency Victoria (`crime_csa`), 2012 onwards | By LGA: Greater Shepparton; Wellington for Yarram, a much larger area than the town. CSA's offence subdivisions are grouped into the eight New South Wales rows (see `fetch_crime_csa.py`); close, but not the same classification. The block sits in the New South Wales part of the Crime sheet, next to the NSW benchmark. |
| Housing: building approvals | ABS (`nsw_approvals`, now New South Wales and Victoria), 2019 onwards | The block is headed by the LGA, "Greater Shepparton (LGA)", "Wellington (LGA)". |
| Housing: Median Sale Price | Valuer-General Victoria, houses by suburb (`vic_sales_rent`), 2015 onwards (earlier years from older editions when they can be downloaded) | **Town locality, houses only** (towns.toml `locality`: Shepparton, Yarram), not the LGA and all dwellings as in the New South Wales rows: VGV publishes no LGA medians. Remove `locality` from a town to leave the row blank instead. |
| Housing: House 3-Bed rent | DFFH Rental Report, quarterly median rents by LGA, 3-bedroom houses (`vic_sales_rent`), 2000 onwards | Mean of the four quarterly medians, complete years only. DFFH's latest file is the September quarter 2025, so 2025 is not written until the December quarter is published. |

Fixed after the first test run (2026-10-29): `add_town.py` copied the
template's typed-in sum "=46+56+67+70" (Narrabri's 2025 sales count)
into the new Housing blocks, where it showed as a 2025 Sales No. of 239.
Formulas of typed-in numbers only are now cleared like other figures.
The ABS approvals query leaves the region type open (BA_LGA2026 holds
its data under LGA2025).

## Not yet filled

- **Mean Sale Price and Sales No.**: VGV's suburb time series has
  medians only. VGV's "Year summary" files (data.vic catalogue,
  `victorian-property-sales-report-time-series`) may give sales counts
  by suburb; not checked, because land.vic.gov.au could not be reached
  from the environment used to build this.
- **The all-dwellings 3-bedroom rent row** ("Total", below House 3-Bed):
  DFFH publishes 3-bedroom houses and flats separately, and medians
  cannot be combined. Left blank.
- **Manual downloads**: land.vic.gov.au and dffh.vic.gov.au refuse some
  automated requests. The fetcher finds the files in the data.vic
  catalogue and tries to download them; if that fails it uses the copies
  saved by hand as `1 Data/Housing/vgv_houses_by_suburb.xlsx` and
  `1 Data/Housing/dffh_rents_by_lga.xlsx` (instructions in that folder).
- **Population projections** (Victoria in Future, planning.vic.gov.au):
  the site blocks the cloud environment used to build this; the LGA and
  SA2 workbooks are needed to write the parser.
- **Benchmark rows for Victoria** (Crime, Employment, Income): the
  workbook has none. The Victorian towns are compared with the New South
  Wales benchmark rows until Victoria rows are added.
- **Yarram's rainfall historic average**: BOM has no Climate Averages
  table for station 085193 that the update can read, and the row had no
  figure to carry forward. Type BOM's figure into the row once; later
  runs carry it forward. (Shepparton, 081125, was read: 448.4 mm.)
- **Fuel prices** and **non-resident workers**: no Victorian source.
- `town_coverage.py` compares a town with others of its state, so it
  lists nothing for the first Victorian towns.
- The Narrabri Surrounds series on the copied pages, as for the New
  South Wales towns.
