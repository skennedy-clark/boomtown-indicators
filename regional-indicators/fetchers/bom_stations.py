"""
regional-indicators/fetchers/bom_stations.py

Names of Bureau of Meteorology stations, by station number, for the
rainfall station label on the Exogenous sheet ("Mitchell Post Office
043020"). The station a town uses is `bom_station` in towns.toml; this
module supplies its name.

Source: BOM's list of all stations,
    https://www.bom.gov.au/climate/data/lists_by_element/stations.txt
(product IDCJMC0014, about 2.6 MB, some 19,500 stations, updated by BOM
from time to time). It is a fixed-width text file:

     Site  Dist  Site name                                 Start     End ...
    ------- ----- ---------------------------------------- ------- ------- ...
     043020 43    MITCHELL POST OFFICE                        1884      .. ...

The station number is in characters 1-7 and the name in characters
15-54. The file is saved as cache/rainfall/bom_stations.txt and
downloaded again only when it is missing (or with refresh=True).

Labels: the name is written in title case ("MITCHELL POST OFFICE" ->
"Mitchell Post Office"), with a few abbreviations kept in capitals
(AWS, TM, ...), followed by the six-digit station number.

Used by transform/add_town.py (label for a new town's station) and
transform/xlsx_update/update_rainfall.py (suggested label when the
sheet's label names another station).
"""

from __future__ import annotations

import re
from pathlib import Path

URL = "https://www.bom.gov.au/climate/data/lists_by_element/stations.txt"
CACHE_FILE = Path(__file__).resolve().parent.parent / "cache" / "rainfall" / "bom_stations.txt"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

# Abbreviations kept in capitals in a label.
CAPITALS = {"AWS", "TM", "WTP", "STP", "PO", "DPI", "CSIRO", "RAAF", "AMO", "HS", "SS", "ADF", "QLD", "NSW", "NT"}

_LINE = re.compile(r"^\s*(\d{4,6})\s+\S+\s+(.+?)\s{2,}\d{4}\s")


def parse(text: str) -> dict[int, str]:
    """{station number: name as BOM writes it} from the stations list."""
    names: dict[int, str] = {}
    for line in text.splitlines():
        match = _LINE.match(line)
        if match:
            names[int(match.group(1))] = match.group(2).strip()
    return names


def station_names(download: bool = True, refresh: bool = False, cache_file: Path = CACHE_FILE) -> dict[int, str]:
    """The stations list, read from the cached copy. When there is no copy
    (or `refresh`), it is downloaded if `download` is set. Returns {} if
    the list is not available."""
    if (refresh or not cache_file.exists()) and download:
        try:
            import requests
            response = requests.get(URL, headers=HEADERS, timeout=120)
            response.raise_for_status()
            if b"Site name" not in response.content[:2000]:
                raise ValueError("not the stations list")
            cache_file.parent.mkdir(parents=True, exist_ok=True)
            cache_file.write_bytes(response.content)
        except Exception:
            pass
    if not cache_file.exists():
        return {}
    return parse(cache_file.read_text(encoding="latin-1"))


def tidy_name(name: str) -> str:
    """BOM's capitals in title case, keeping known abbreviations."""
    words = []
    for word in name.split():
        words.append(word if word.upper() in CAPITALS else word.capitalize() if word.isalpha()
                     else re.sub(r"[A-Za-z]+", lambda m: m.group(0).upper() if m.group(0).upper() in CAPITALS
                                 else m.group(0).capitalize(), word))
    return " ".join(words)


def station_label(number, names: dict[int, str]) -> str | None:
    """'Mitchell Post Office 043020' for station 43020, or None if the
    station is not in the list."""
    try:
        key = int(str(number).strip())
    except ValueError:
        return None
    name = names.get(key)
    return f"{tidy_name(name)} {key:06d}" if name else None
