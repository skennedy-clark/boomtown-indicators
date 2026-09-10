"""
fetch_qld_tenure.py -- Download current Queensland petroleum tenure
boundaries (exploration permits and production leases) from the
Department of Resources' live ArcGIS REST service, and load them into
the project's DuckDB spatial geodatabase.

Fits into the existing boomtown-indicators fetcher pattern: one fetcher
per data source, downloading raw data into cache/ before any
transformation, following the fine-grained-modularity principle so a
non-programmer maintainer can fix just this one piece if it breaks.

Source: Queensland Department of Resources, "Economy/MinesPermitsCurrent"
        ArcGIS REST map service. Updated nightly from the MERLIN tenures
        system. GDA2020 source data.
        https://spatial-gis.information.qld.gov.au/arcgis/rest/services/Economy/MinesPermitsCurrent/MapServer

Tenure types fetched:
    EPP  Exploration Permit for Petroleum (the current name for what used
         to be called an Authority to Prospect / ATP -- same concept:
         exploration-stage tenure, no production allowed).
    PL   Petroleum Lease (production tenure; granted once a discovery is
         proven; production must generally commence within a set period
         after grant).

Each type has an 'application' and a 'granted' layer on the service --
this fetcher pulls both and tags every record with tenure_type and
tenure_status columns so pending vs. current tenure is distinguishable
in one table.

Usage:
    # Fetch current QLD petroleum tenure (EPP + PL, application + granted)
    python fetch_qld_tenure.py

    # Just resolve and print the current layer IDs, don't download features
    # (useful first check if the service is reorganised and IDs shift)
    python fetch_qld_tenure.py --list-layers

    # Re-download even if a cached copy exists
    python fetch_qld_tenure.py --force

Output:
    cache/qld_tenure/epp_granted.geojson
    cache/qld_tenure/epp_application.geojson
    cache/qld_tenure/pl_granted.geojson
    cache/qld_tenure/pl_application.geojson
    geodata.duckdb -> table qld_tenure (all four, combined, with
                       tenure_type and tenure_status columns)

Notes:
    - Layer IDs on this service are NOT hardcoded -- they're looked up by
      name each run via /MapServer/layers, because ArcGIS service layer
      numbers can shift when the Department reorganises the service. This
      costs one extra request but is far more robust than a fetcher that
      silently starts pulling the wrong layer after a reshuffle.
    - The service paginates at 2000 features per request; this fetcher
      follows resultOffset until a page comes back short of the page
      size, so it stays correct as tenure counts grow.
    - A granted PL does not by itself mean gas is flowing -- cross-check
      against the QLD petroleum and gas production statistics
      (data.qld.gov.au) at the PL level, and/or well status from the QLD
      borehole series, to work out whether a given lease has actually
      started producing.
    - NSW petroleum titles (PEL/PPL/PAL) live on a parallel ArcGIS REST
      service at spatial.industry.nsw.gov.au/arcgis/rest/services/
      Minerals/CurrentTitles/MapServer -- a fetch_nsw_tenure.py following
      this same pattern is the natural next fetcher for Narrabri.
"""

import argparse
import json
from pathlib import Path

import duckdb
import requests

BASE_URL = "https://spatial-gis.information.qld.gov.au/arcgis/rest/services/Economy/MinesPermitsCurrent/MapServer"

# Layer names to look for, matched case-insensitively against the 'name'
# field returned by /MapServer/layers.
WANTED_LAYERS = {
    "epp_granted": "EPP granted",
    "epp_application": "EPP application",
    "pl_granted": "PL granted",
    "pl_application": "PL application",
}

CACHE_DIR = Path("cache/qld_tenure")
DB_PATH = Path("geodata.duckdb")
PAGE_SIZE = 2000


def list_layers() -> dict:
    """Fetch the current layer list from the ArcGIS service and return a
    {layer_name: layer_id} lookup. Prints the resolved mapping for the
    wanted layers so it can be sanity-checked against the source before
    pulling any features.

    Example:
        >>> layers = list_layers()
        Resolved layer IDs:
          epp_granted      -> layer 42 ('EPP granted')
          ...
    """
    resp = requests.get(f"{BASE_URL}/layers", params={"f": "json"}, timeout=60)
    resp.raise_for_status()
    layers = resp.json().get("layers", [])
    name_to_id = {layer["name"]: layer["id"] for layer in layers}

    print("Resolved layer IDs:")
    for key, wanted_name in WANTED_LAYERS.items():
        matches = [n for n in name_to_id if n.lower() == wanted_name.lower()]
        if matches:
            print(f"  {key:16s} -> layer {name_to_id[matches[0]]} ('{matches[0]}')")
        else:
            print(f"  {key:16s} -> NOT FOUND (looked for '{wanted_name}')")
    return name_to_id


def fetch_layer_features(layer_id: int) -> dict:
    """Page through an ArcGIS feature layer and return a combined GeoJSON
    FeatureCollection. Uses resultOffset pagination at PAGE_SIZE per
    request until a page comes back short of the page size.

    Example:
        >>> fc = fetch_layer_features(42)
        >>> fc["type"]
        'FeatureCollection'
    """
    features = []
    offset = 0
    while True:
        params = {
            "where": "1=1",
            "outFields": "*",
            "f": "geojson",
            "resultRecordCount": PAGE_SIZE,
            "resultOffset": offset,
        }
        resp = requests.get(f"{BASE_URL}/{layer_id}/query", params=params, timeout=120)
        resp.raise_for_status()
        page = resp.json()
        page_features = page.get("features", [])
        features.extend(page_features)
        if len(page_features) < PAGE_SIZE:
            break
        offset += PAGE_SIZE
    return {"type": "FeatureCollection", "features": features}


def download_tenure(force: bool = False) -> dict:
    """Download all four tenure layers (EPP/PL x application/granted) as
    GeoJSON into cache/qld_tenure/. Returns {key: path} for the files
    that were successfully downloaded or already cached.
    """
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    name_to_id = list_layers()
    paths = {}

    for key, wanted_name in WANTED_LAYERS.items():
        out_path = CACHE_DIR / f"{key}.geojson"
        if out_path.exists() and not force:
            print(f"Using cached {out_path}")
            paths[key] = out_path
            continue

        matches = [n for n in name_to_id if n.lower() == wanted_name.lower()]
        if not matches:
            print(f"WARNING: skipping {key}, layer '{wanted_name}' not found on service")
            continue

        layer_id = name_to_id[matches[0]]
        print(f"Fetching {key} (layer {layer_id}) ...")
        collection = fetch_layer_features(layer_id)
        out_path.write_text(json.dumps(collection))
        print(f"  {len(collection['features'])} features -> {out_path}")
        paths[key] = out_path

    return paths


def load_into_duckdb(paths: dict) -> None:
    """Union all fetched tenure layers into a single qld_tenure table in
    geodata.duckdb, tagged with tenure_type (EPP/PL) and tenure_status
    (application/granted). Uses UNION ALL BY NAME so the four layers
    don't need identical column sets.
    """
    con = duckdb.connect(str(DB_PATH))
    con.execute("INSTALL spatial; LOAD spatial;")

    selects = []
    for key, path in paths.items():
        tenure_type, tenure_status = key.split("_")
        selects.append(f"""
            SELECT '{tenure_type.upper()}' AS tenure_type,
                   '{tenure_status}' AS tenure_status,
                   *
            FROM ST_Read('{path}')
        """)

    if not selects:
        print("No tenure layers downloaded -- nothing to load.")
        con.close()
        return

    union_sql = "\nUNION ALL BY NAME\n".join(selects)
    con.execute(f"CREATE OR REPLACE TABLE qld_tenure AS {union_sql}")

    count = con.execute("SELECT COUNT(*) FROM qld_tenure").fetchone()[0]
    print(f"Loaded {count} tenure records into {DB_PATH} (qld_tenure)")
    con.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--list-layers", action="store_true", help="Only print resolved layer IDs and exit")
    parser.add_argument("--force", action="store_true", help="Re-download even if cached")
    args = parser.parse_args()

    if args.list_layers:
        list_layers()
        return

    paths = download_tenure(force=args.force)
    load_into_duckdb(paths)


if __name__ == "__main__":
    main()
