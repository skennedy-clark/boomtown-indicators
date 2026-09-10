"""
fetch_sa2.py -- Download ABS ASGS Edition 3 (2021) SA2 boundaries and load
them into the project's DuckDB spatial geodatabase.

Fits into the existing boomtown-indicators fetcher pattern: one fetcher per
data source, downloading raw data into cache/ before any transformation.

Source: Australian Bureau of Statistics, ASGS Edition 3, July 2021 - June 2026
        Digital boundary files (ESRI Shapefile, GDA2020)
        https://www.abs.gov.au/statistics/standards/australian-statistical-geography-standard-asgs/edition-3-july-2021-june-2026/access-and-downloads/digital-boundary-files

Usage:
    # Fetch SA2 boundaries for QLD + NSW (default) and load into DuckDB
    python fetch_sa2.py

    # Fetch for a specific state, several states, or all of Australia
    python fetch_sa2.py --states QLD
    python fetch_sa2.py --states QLD,NSW,VIC
    python fetch_sa2.py --states ALL

    # Re-download even if a cached copy exists
    python fetch_sa2.py --force

    # Just refresh the cached shapefile without touching the database
    python fetch_sa2.py --skip-load

Output:
    cache/sa2/SA2_2021_AUST_GDA2020.zip   (raw ABS download, ~48 MB, cached)
    cache/sa2/SA2_2021_AUST_GDA2020.shp   (extracted shapefile + sidecars)
    geodata.duckdb -> table sa2_boundaries (filtered to requested states)

Notes:
    - ASGS Edition 3 (2021) is the current stable release through June 2026
      and is a completed dataset. Edition 4 (2026 vintage) boundaries are
      being rolled out progressively from 2026 and are not fully published
      yet -- stay on 2021 for now and revisit fetch_sa2.py when Edition 4
      is complete across all structures.
    - ABS state codes: 1=NSW, 2=VIC, 3=QLD, 4=SA, 5=WA, 6=TAS, 7=NT, 8=ACT,
      9=Other Territories.
    - This fetcher targets SA2 only. LGA and UCL are separate ABS products
      (Non ABS Structures, and SUA/UCL/SOS/SOSR respectively) with their
      own file naming -- add fetch_lga.py / fetch_ucl.py following the same
      pattern when those are needed, rather than overloading this one.
"""

import argparse
import zipfile
from pathlib import Path

import duckdb
import requests

ABS_SA2_URL = (
    "https://www.abs.gov.au/statistics/standards/australian-statistical-"
    "geography-standard-asgs/edition-3-july-2021-june-2026/access-and-"
    "downloads/digital-boundary-files/SA2_2021_AUST_SHP_GDA2020.zip"
)

# ABS state/territory codes as used in the STE_CODE21 field
STATE_CODES = {
    "NSW": "1", "VIC": "2", "QLD": "3", "SA": "4", "WA": "5",
    "TAS": "6", "NT": "7", "ACT": "8", "OT": "9",
}

CACHE_DIR = Path("cache/sa2")
DB_PATH = Path("geodata.duckdb")


def download_sa2_shapefile(force: bool = False) -> Path:
    """Download the national SA2 2021 GDA2020 shapefile zip from ABS.

    Cached in cache/sa2/ so re-runs don't re-download the ~48 MB file
    unless force=True. Returns the path to the extracted .shp file.

    Example:
        >>> shp = download_sa2_shapefile()
        >>> shp.name
        'SA2_2021_AUST_GDA2020.shp'
    """
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    zip_path = CACHE_DIR / "SA2_2021_AUST_GDA2020.zip"
    shp_path = CACHE_DIR / "SA2_2021_AUST_GDA2020.shp"

    if not zip_path.exists() or force:
        print(f"Downloading {ABS_SA2_URL} ...")
        resp = requests.get(ABS_SA2_URL, timeout=180)
        resp.raise_for_status()
        zip_path.write_bytes(resp.content)
        print(f"Saved {len(resp.content) / 1e6:.1f} MB to {zip_path}")
    else:
        print(f"Using cached {zip_path}")

    if not shp_path.exists() or force:
        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(CACHE_DIR)
        print(f"Extracted to {CACHE_DIR}")

    return shp_path


def load_into_duckdb(shp_path: Path, states: list) -> None:
    """Load the SA2 shapefile into geodata.duckdb, filtered to the given
    states, replacing the sa2_boundaries table.

    Requires the DuckDB spatial extension (installed automatically on
    first use -- needs network access the first time only).

    Example:
        >>> load_into_duckdb(Path("cache/sa2/SA2_2021_AUST_GDA2020.shp"), ["QLD", "NSW"])
        Loaded 611 SA2 boundaries into geodata.duckdb (sa2_boundaries)
    """
    con = duckdb.connect(str(DB_PATH))
    con.execute("INSTALL spatial; LOAD spatial;")

    if "ALL" in states:
        state_filter = ""
    else:
        codes = ", ".join(f"'{STATE_CODES[s]}'" for s in states)
        state_filter = f"WHERE STE_CODE21 IN ({codes})"

    con.execute(f"""
        CREATE OR REPLACE TABLE sa2_boundaries AS
        SELECT
            SA2_CODE21 AS sa2_code,
            SA2_NAME21 AS sa2_name,
            SA3_NAME21 AS sa3_name,
            SA4_NAME21 AS sa4_name,
            STE_NAME21 AS state_name,
            AREASQKM21 AS area_sqkm,
            geom
        FROM ST_Read('{shp_path}')
        {state_filter}
    """)

    count = con.execute("SELECT COUNT(*) FROM sa2_boundaries").fetchone()[0]
    print(f"Loaded {count} SA2 boundaries into {DB_PATH} (sa2_boundaries)")
    con.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "--states", default="QLD,NSW",
        help="Comma-separated state codes (QLD,NSW,VIC,...) or ALL. Default: QLD,NSW",
    )
    parser.add_argument("--force", action="store_true", help="Re-download even if cached")
    parser.add_argument("--skip-load", action="store_true", help="Download only, skip DuckDB load")
    args = parser.parse_args()

    states = [s.strip().upper() for s in args.states.split(",")]
    shp_path = download_sa2_shapefile(force=args.force)

    if not args.skip_load:
        load_into_duckdb(shp_path, states)


if __name__ == "__main__":
    main()
