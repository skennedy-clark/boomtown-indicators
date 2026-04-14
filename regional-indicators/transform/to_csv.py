"""
transform/to_csv.py
-------------------
Converts cached indicator JSON files into the 2-row CSV format used by
the boomtown-indicators.org website.

CSV format (from reference files):
  Row 1: "2000","2001","2002",...,"2024"   ← years as quoted strings
  Row 2: "","value","value",...,""          ← values, empty string = no data

Year conventions:
  ATO financial years map to the end calendar year:
    2003-04 → "2004",  2022-23 → "2023"
  All other sources use calendar year directly.

Usage:
  python transform/to_csv.py                     # all towns, all indicators
  python transform/to_csv.py --towns Roma Dalby  # specific towns
  python transform/to_csv.py --only income        # specific indicator
  python transform/to_csv.py --dry-run            # show what would be written
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from config import get_config, OUTPUT_DIR, CACHE_DIR, Town
from logger import get_logger, get_child_logger

log = get_child_logger("regional-indicators", "to_csv")


# ── Year range ─────────────────────────────────────────────────────────────────

YEAR_START = 2000
YEAR_END   = 2025   # update each cycle
ALL_YEARS  = [str(y) for y in range(YEAR_START, YEAR_END + 1)]


# ── ATO financial year → calendar year ─────────────────────────────────────────

def fy_to_year(fy: str) -> Optional[str]:
    """
    Convert "2022-23" or "2022–23" to "2023" (the end year).
    Returns None if not parseable.
    """
    fy = fy.replace("\u2013", "-").strip()   # en-dash → hyphen
    parts = fy.split("-")
    if len(parts) == 2:
        start, end_2 = parts
        try:
            end_full = str(int(start) + 1)
            # Sanity check: end_2 should match last 2 digits
            if end_full[-2:] == end_2.zfill(2)[-2:]:
                return end_full
        except ValueError:
            pass
    return None


# ── CSV writer ─────────────────────────────────────────────────────────────────

def write_csv(path: Path, values: dict[str, str | float | int | None],
              dry_run: bool = False) -> None:
    """
    Write the 2-row CSV.  `values` maps year string → value (or None/empty).
    """
    row1 = ALL_YEARS
    row2 = []
    for yr in ALL_YEARS:
        val = values.get(yr)
        if val is None or val == "":
            row2.append("")
        elif isinstance(val, float):
            # Drop trailing .0 for whole numbers
            row2.append(str(int(val)) if val == int(val) else str(val))
        else:
            row2.append(str(val))

    if dry_run:
        log.info(f"  [dry-run] would write {path.name}")
        return

    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, quoting=csv.QUOTE_ALL)
        w.writerow(row1)
        w.writerow(row2)
    log.info(f"  Wrote {path.relative_to(OUTPUT_DIR)}")


# ── Indicator transformers ──────────────────────────────────────────────────────

def transform_income_table8(town: Town, dry_run: bool = False) -> bool:
    """
    Cache: cache/ato/{slug}_income.json
    Produces: output/{town}/Income - taxable, incl. lowneg. incomes (ATO ave.).csv

    JSON structure:
      { "avg_taxable_income_by_year": { "2003-04": 38500.0, ..., "2022-23": 65642.0 } }

    Table 8 includes ALL individuals (taxable + non-taxable), so this is the
    "incl. lowneg. incomes" series.
    """
    cache_file = CACHE_DIR / "ato" / f"{town.slug}_income.json"
    if not cache_file.exists():
        log.warning(f"  [{town.name}] income cache missing: {cache_file.name}")
        return False

    with open(cache_file) as f:
        data = json.load(f)

    raw = data.get("avg_taxable_income_by_year", {})
    values: dict[str, float | None] = {}
    for fy, val in raw.items():
        yr = fy_to_year(str(fy))
        if yr and YEAR_START <= int(yr) <= YEAR_END:
            values[yr] = val

    out = town.output_dir / "Income - taxable, incl. lowneg. incomes (ATO ave.).csv"
    write_csv(out, values, dry_run=dry_run)
    return True




# ── Table 6 transformers ──────────────────────────────────────────────────────

def _load_t6_series(town: Town, indicator: str) -> dict:
    """Load all *_income_t6.json files for a town, merge across years."""
    cache_dir = CACHE_DIR / "ato"
    values = {}
    for pat in [f"{town.slug}_income_t6.json", f"{town.slug}_income_t6_*.json"]:
        for fpath in sorted(cache_dir.glob(pat)):
            with open(fpath) as f:
                data = json.load(f)
            for yr, val in data.get("indicators", {}).get(indicator, {}).items():
                if val is not None and YEAR_START <= int(yr) <= YEAR_END:
                    values[str(yr)] = val
    return values


def transform_income_table6_taxable(town: Town, dry_run: bool = False) -> bool:
    """Cache: cache/ato/{slug}_income_t6.json → Income - for taxable individuals.csv"""
    values = _load_t6_series(town, "avg_income_taxable")
    if not values:
        log.warning(f"  [{town.name}] income_t6 taxable cache missing")
        return False
    write_csv(town.output_dir / "Income - for taxable individuals.csv", values, dry_run=dry_run)
    return True


def transform_earners(town: Town, dry_run: bool = False) -> bool:
    """Cache: cache/ato/{slug}_income_t6.json → Number of earners.csv"""
    values = _load_t6_series(town, "earners_no")
    if not values:
        log.warning(f"  [{town.name}] income_t6 earners cache missing")
        return False
    write_csv(town.output_dir / "Number of earners.csv", values, dry_run=dry_run)
    return True


def transform_wages(town: Town, dry_run: bool = False) -> bool:
    """Cache: cache/ato/{slug}_income_t6.json → Wage & salary earnings (town total).csv"""
    values = _load_t6_series(town, "wages_total")
    if not values:
        log.warning(f"  [{town.name}] income_t6 wages cache missing")
        return False
    write_csv(town.output_dir / "Wage & salary earnings (town total).csv", values, dry_run=dry_run)
    return True



def transform_population_ucl(town: Town, dry_run: bool = False) -> bool:
    """
    Cache: cache/population/{slug}_population_ucl.json
    Produces: output/{town}/Population - town.csv
    Only applies to QLD towns with a UCL match (not Toowoomba sub-areas).
    """
    cache_file = CACHE_DIR / "population" / f"{town.slug}_population_ucl.json"
    if not cache_file.exists():
        log.warning(f"  [{town.name}] population_ucl cache missing")
        return False

    with open(cache_file) as f:
        data = json.load(f)

    raw = data.get("population_by_year", {})
    values = {yr: val for yr, val in raw.items()
              if YEAR_START <= int(yr) <= YEAR_END}

    out = town.output_dir / "Population - town.csv"
    write_csv(out, values, dry_run=dry_run)
    return True



# ── Crime transformers (QPS) ──────────────────────────────────────────────────

def _load_crime_series(town: Town, indicator: str) -> dict:
    """Load cache/crime/{slug}_crime_qps.json and return {year: value} dict."""
    cache_file = CACHE_DIR / "crime" / f"{town.slug}_crime_qps.json"
    if not cache_file.exists():
        return {}
    with open(cache_file) as f:
        data = json.load(f)
    return {
        yr: val for yr, val in data.get("indicators", {}).get(indicator, {}).items()
        if val is not None and YEAR_START <= int(yr) <= YEAR_END
    }


def transform_crime_all(town: Town, dry_run: bool = False) -> bool:
    """→ Crime rate - all offences.csv"""
    values = _load_crime_series(town, "all")
    if not values:
        log.warning(f"  [{town.name}] crime_qps cache missing")
        return False
    write_csv(town.output_dir / "Crime rate - all offences.csv", values, dry_run=dry_run)
    return True


def transform_crime_drugs(town: Town, dry_run: bool = False) -> bool:
    """→ Drug offences.csv"""
    values = _load_crime_series(town, "drug")
    if not values:
        log.warning(f"  [{town.name}] crime_qps cache missing")
        return False
    write_csv(town.output_dir / "Drug offences.csv", values, dry_run=dry_run)
    return True


def transform_crime_goodorder(town: Town, dry_run: bool = False) -> bool:
    """→ Good order offences.csv"""
    values = _load_crime_series(town, "good_order")
    if not values:
        log.warning(f"  [{town.name}] crime_qps cache missing")
        return False
    write_csv(town.output_dir / "Good order offences.csv", values, dry_run=dry_run)
    return True


def transform_crime_theft(town: Town, dry_run: bool = False) -> bool:
    """→ Theft.csv"""
    values = _load_crime_series(town, "theft")
    if not values:
        log.warning(f"  [{town.name}] crime_qps cache missing")
        return False
    write_csv(town.output_dir / "Theft.csv", values, dry_run=dry_run)
    return True


def transform_crime_traffic(town: Town, dry_run: bool = False) -> bool:
    """→ Traffic offences.csv"""
    values = _load_crime_series(town, "traffic")
    if not values:
        log.warning(f"  [{town.name}] crime_qps cache missing")
        return False
    write_csv(town.output_dir / "Traffic offences.csv", values, dry_run=dry_run)
    return True



# ── QGSO housing / unemployment / NRW / rainfall transformers ─────────────────

def _load_qgso_series(town: Town, indicator: str) -> dict:
    """
    Load QGSO housing cache for a town.

    Matches two filename patterns (both are produced by fetch_qgso_housing.py):
      {slug}_qgso.json          ← current format (single file per town)
      {slug}_qgso_*.json        ← legacy/split format (multiple files per town)

    Values from all matching files are merged; later files win on key collisions.
    """
    cache_dir = CACHE_DIR / "housing"
    values = {}

    # Collect all matching files, deduplicated, in sorted order
    matched: list[Path] = []
    seen: set[Path] = set()
    for pat in [f"{town.slug}_qgso.json", f"{town.slug}_qgso_*.json"]:
        for fpath in sorted(cache_dir.glob(pat)):
            if fpath not in seen:
                matched.append(fpath)
                seen.add(fpath)
    matched.sort()

    for fpath in matched:
        with open(fpath) as f:
            data = json.load(f)
        for yr, val in data.get("indicators", {}).get(indicator, {}).items():
            if val is not None and YEAR_START <= int(yr) <= YEAR_END:
                values[str(yr)] = val

    return values


def _make_housing_transformer(indicator: str, filename: str):
    """Factory for housing/NRW/rainfall transformer functions."""
    def transformer(town: Town, dry_run: bool = False) -> bool:
        values = _load_qgso_series(town, indicator)
        if not values:
            log.warning(f"  [{town.name}] qgso cache missing for {indicator}")
            return False
        write_csv(town.output_dir / filename, values, dry_run=dry_run)
        return True
    transformer.__name__ = f"transform_{indicator}"
    return transformer


transform_house_price = _make_housing_transformer(
    "housing_median_price",          # JSON key in {slug}_qgso.json
    "House sale price (median).csv",
)
transform_house_sales = _make_housing_transformer(
    "housing_sales_count",           # JSON key in {slug}_qgso.json
    "House sales (number).csv",
)
transform_rent = _make_housing_transformer(
    "rent_3bed_median",              # JSON key in {slug}_qgso.json
    "Rent (3-bedroom house; median).csv",
)
transform_approvals = _make_housing_transformer(
    "building_approvals",            # JSON key in {slug}_qgso.json
    "Residential building approvals.csv",
)


# Unemployment comes from SALM cache (not QGSO housing cache)
def transform_unemployment(town: Town, dry_run: bool = False) -> bool:
    """→ Unemployment rate.csv  (source: SALM or QGSO housing cache)"""
    # Try SALM cache first (automated, national)
    salm_file = CACHE_DIR / "unemployment" / f"{town.slug}_salm.json"
    if salm_file.exists():
        with open(salm_file) as f:
            data = json.load(f)
        values = {yr: val for yr, val in data.get("indicators", {}).get("unemployment", {}).items()
                  if YEAR_START <= int(yr) <= YEAR_END}
        if values:
            write_csv(town.output_dir / "Unemployment rate.csv", values, dry_run=dry_run)
            return True
    # Fallback: QGSO housing cache (manual)
    return _make_housing_transformer("unemployment", "Unemployment rate.csv")(town, dry_run)


transform_nrw_town = _make_housing_transformer("nrw_town", "Non-resident workers in town.csv")
transform_nrw_lga  = _make_housing_transformer("nrw_lga",  "Non-resident workers in local govt area.csv")


# Rainfall comes from BOM/SILO cache (not QGSO housing cache)
def transform_rainfall(town: Town, dry_run: bool = False) -> bool:
    """→ Environment - total rainfall.csv  (source: BOM/SILO or QGSO housing cache)"""
    bom_file = CACHE_DIR / "rainfall" / f"{town.slug}_bom_rainfall.json"
    if bom_file.exists():
        with open(bom_file) as f:
            data = json.load(f)
        values = {yr: val for yr, val in data.get("indicators", {}).get("rainfall", {}).items()
                  if YEAR_START <= int(yr) <= YEAR_END}
        if values:
            write_csv(town.output_dir / "Environment - total rainfall.csv", values, dry_run=dry_run)
            return True
    # Fallback: QGSO housing cache (manual)
    return _make_housing_transformer("rainfall", "Environment - total rainfall.csv")(town, dry_run)


# ── Registry ───────────────────────────────────────────────────────────────────

# Maps indicator key → transform function signature: fn(town, dry_run) -> bool
TRANSFORMERS: dict[str, callable] = {
    "income":              transform_income_table8,
    "income_taxable":      transform_income_table6_taxable,
    "earners":             transform_earners,
    "wages":               transform_wages,
    "population_town":     transform_population_ucl,
    "crime_all":           transform_crime_all,
    "crime_drugs":         transform_crime_drugs,
    "crime_goodorder":     transform_crime_goodorder,
    "crime_theft":         transform_crime_theft,
    "crime_traffic":       transform_crime_traffic,
    "house_price":         transform_house_price,
    "house_sales":         transform_house_sales,
    "rent":                transform_rent,
    "approvals":           transform_approvals,
    "unemployment":        transform_unemployment,
    "nrw_town":            transform_nrw_town,
    "nrw_lga":             transform_nrw_lga,
    "rainfall":            transform_rainfall,
}


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Convert cached data to CSVs")
    parser.add_argument("--towns",   nargs="+", help="Filter to specific towns")
    parser.add_argument("--only",    nargs="+", help="Filter to specific indicators")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be written")
    args = parser.parse_args()

    config  = get_config()
    towns   = config.study_towns()

    if args.towns:
        wanted = {t.lower() for t in args.towns}
        towns  = [t for t in towns if t.name.lower() in wanted]
        if not towns:
            log.error(f"No towns matched: {args.towns}")
            sys.exit(1)

    indicators = list(TRANSFORMERS.keys())
    if args.only:
        wanted_ind = {i.lower() for i in args.only}
        indicators = [i for i in indicators if i.lower() in wanted_ind]
        if not indicators:
            log.error(f"No indicators matched: {args.only}. Available: {list(TRANSFORMERS)}")
            sys.exit(1)

    if args.dry_run:
        log.info("DRY RUN — no files will be written")

    ok = fail = skip = 0
    for town in towns:
        for ind in indicators:
            fn = TRANSFORMERS[ind]
            try:
                success = fn(town, dry_run=args.dry_run)
                if success:
                    ok += 1
                else:
                    skip += 1
            except Exception as exc:
                log.error(f"  [{town.name}] {ind}: {exc}", exc_info=True)
                fail += 1

    log.info(
        f"\nDone — {ok} written, {skip} skipped (no cache), {fail} errors"
        + (" [dry-run]" if args.dry_run else "")
    )
    sys.exit(0 if fail == 0 else 1)


if __name__ == "__main__":
    get_logger()
    main()