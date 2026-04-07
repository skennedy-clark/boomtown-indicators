import requests
import pandas as pd
from pathlib import Path

# ---------------------------------------------
# CONFIGURATION
# ---------------------------------------------

# Station numbers (must be 6‑digit padded strings)
stations = [
    "041240",  # Dalby
    "043091",  # Roma
    "042078",  # Chinchilla
]

# Product ID for monthly rainfall
product = "IDCJAC0001"

# Output folder
out_dir = Path("bom_data")
out_dir.mkdir(exist_ok=True)

# ---------------------------------------------
# FUNCTION TO DOWNLOAD A SINGLE STATION
# ---------------------------------------------

def download_station(product, station):
    """
    Downloads a single IDCJAC0001 CSV from BOM Climate Data Online API.
    Returns a pandas DataFrame.
    """
    url = (
        f"https://api.bom.gov.au/v1/climate/data/{product}"
        f"?station={station}&format=csv"
    )

    print(f"Downloading {station} ...")

    r = requests.get(url)
    r.raise_for_status()

    csv_path = out_dir / f"{product}_{station}.csv"
    csv_path.write_bytes(r.content)

    df = pd.read_csv(csv_path)
    df["station"] = station  # ensure station column exists

    return df


# ---------------------------------------------
# DOWNLOAD ALL STATIONS AND MERGE
# ---------------------------------------------

all_data = []

for stn in stations:
    try:
        df = download_station(product, stn)
        all_data.append(df)
    except Exception as e:
        print(f"Failed for station {stn}: {e}")

# Combine into one long-format DataFrame
if all_data:
    combined = pd.concat(all_data, ignore_index=True)
    combined.to_csv(out_dir / "IDCJAC0001_all_stations.csv", index=False)
    print("Saved combined dataset to bom_data/IDCJAC0001_all_stations.csv")
else:
    print("No data downloaded.")