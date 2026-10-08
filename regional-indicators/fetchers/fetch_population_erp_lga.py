"""
regional-indicators/fetchers/fetch_population_erp_lga.py

Fetches LGA-level estimated resident population (ERP): the "Population
(ERP)" row in each block of the Population sheet's LGA section.

Source: ABS Data API, dataflow "ERP by LGA (<year>), 2001 to <year>"
(id ERP_LGA<year>, e.g. ERP_LGA2025). One national dataset covering
every LGA in Australia, with annual history from 2001 on the current
boundary vintage.

The ABS dataflow is used in preference to the QGSO regional database
(QRSIS) because it covers every state from a single query and needs no
multi-step session. Its Queensland figures are identical to QRSIS's,
which republishes the ABS series.

API notes (same host and conventions as fetch_narrabri_approvals.py):
  - Host: https://data.api.abs.gov.au/rest/
  - A new ERP_LGA<year> dataflow is published each year on that year's
    LGA boundaries. The available dataflows are listed at run time and
    the newest is used, so no annual code change is needed.
  - Dimension order: MEASURE . REGION_TYPE . REGION . FREQ
    Query key: "ERP..<code>+<code>+....A". REGION_TYPE is left blank
    (wildcard) because its code is vintage-specific ("LGA2025").
  - Values are persons at 30 June of TIME_PERIOD, which is the
    workbook's year header (2025 = the 2024/25 column).
  - A code that does not exist in the dataflow is omitted from the
    response without an error. This is reported as a failure for that
    region so that an incorrect code in towns.toml is visible.

Regions: every distinct LGA across the towns in towns.toml, benchmark
towns included (config.lgas()), each fetched once however many towns
share it. A town's LGA name is its `lga` field, which must match the
workbook's LGA block label. The code is taken from `qgso_lga`
("LGA/37310" -> 37310; QGSO uses the ABS code) or, for towns outside
Queensland, from `lga_code`. To add an LGA, add or edit a town.

Output: cache/population/lga_<slug>_population_erp_lga.json, one per
LGA:
  {"region": "Brisbane", "level": "LGA", "state": "QLD",
   "abs_lga": "31000", "source": "...", "dataflow": "ERP_LGA2025",
   "year": 2025, "value": 1375301,
   "series_by_year": {"2001": 885787, ..., "2025": 1375301}}
Read by transform/xlsx_update/update_population_erp_lga.py. The file
name suffix "_population_erp_lga.json" is not matched by the SA2
writer's "*_population_erp.json" pattern.

Usage:
    python run_update.py --only population_erp_lga
    python fetchers/fetch_population_erp_lga.py        # standalone
"""

from __future__ import annotations

import csv
import io
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import YEAR_START

try:
    import requests
except ImportError:
    raise ImportError("pip install requests")


API_BASE        = "https://data.api.abs.gov.au/rest"
DATAFLOW_PREFIX = "ERP_LGA"          # ERP_LGA2025, ERP_LGA2026, ...
MEASURE         = "ERP"
FREQ            = "A"
CSV_ACCEPT      = "application/vnd.sdmx.data+csv"
TIMEOUT_S       = 60


def _discover_newest_dataflow(log) -> str | None:
    """List the ABS dataflows and return the newest ERP_LGA<year> id.

    Matches ERP_LGA followed by exactly four digits, which excludes the
    similarly named ABS_ERP_LGA2018, ABS_ANNUAL_ERP_LGA2025 (age and sex
    breakdown) and ERP_COMP_LGA2025 (components of change) dataflows.
    """
    try:
        resp = requests.get(
            f"{API_BASE}/dataflow/ABS",
            params={"detail": "allstubs"},
            timeout=TIMEOUT_S,
        )
        resp.raise_for_status()
    except requests.RequestException as exc:
        log.error(f"  Could not list ABS dataflows: {exc}")
        return None

    ids = sorted(set(re.findall(rf'id="({DATAFLOW_PREFIX}\d{{4}})"', resp.text)))
    if not ids:
        log.error(f"  No {DATAFLOW_PREFIX}<year> dataflow found in the ABS listing")
        return None
    log.info(f"  Discovered {len(ids)} ERP-by-LGA dataflows: {ids[0]}..{ids[-1]}; using {ids[-1]}")
    return ids[-1]


def _parse_csv(text: str) -> dict[str, dict[int, int]]:
    """Parse SDMX-CSV into {lga_code: {year: persons}}, using the REGION,
    TIME_PERIOD and OBS_VALUE columns.
    """
    result: dict[str, dict[int, int]] = {}
    for row in csv.DictReader(io.StringIO(text)):
        code = (row.get("REGION") or "").strip()
        period = (row.get("TIME_PERIOD") or "").strip()
        raw = (row.get("OBS_VALUE") or "").strip()
        if not code or not period[:4].isdigit() or not raw:
            continue
        try:
            value = int(round(float(raw)))
        except ValueError:
            continue
        result.setdefault(code, {})[int(period[:4])] = value
    return result


class ABSPopulationERPLGAFetcher(BaseFetcher):

    SOURCE_NAME      = "population_erp_lga"
    SUPPORTED_STATES = []   # national dataset: any state or territory

    def fetch_all(self):
        # (name, ABS code, state) for each distinct LGA, benchmarks included
        regions = self.config.lgas()
        if not regions:
            self.log.info("No towns in towns.toml have an LGA code — nothing to fetch")
            return

        dataflow = _discover_newest_dataflow(self.log)
        if not dataflow:
            self.result.add_error("ALL", "Could not discover an ERP_LGA<year> dataflow")
            return

        codes = "+".join(code for _, code, _ in regions)
        key = f"{MEASURE}..{codes}.{FREQ}"
        url = f"{API_BASE}/data/ABS,{dataflow}/{key}"
        try:
            resp = requests.get(
                url,
                params={"startPeriod": str(YEAR_START)},
                headers={"Accept": CSV_ACCEPT},
                timeout=TIMEOUT_S,
            )
            resp.raise_for_status()
        except requests.RequestException as exc:
            self.result.add_error("ALL", f"ABS query failed ({url}): {exc}")
            return

        data = _parse_csv(resp.text)
        self.log.info(f"  Parsed {len(data)} of {len(regions)} LGAs from {dataflow}")

        out_dir = Path(__file__).parent.parent / "cache" / "population"
        out_dir.mkdir(parents=True, exist_ok=True)

        for name, code, state in regions:
            label = f"{name} (LGA)"
            slug = name.lower().replace(" ", "_").replace("-", "_")
            series = data.get(code)
            if not series:
                self.result.add_error(
                    label,
                    f"LGA code {code} returned no data from {dataflow} -- "
                    f"check qgso_lga / lga_code in towns.toml against the ABS LGA codelist "
                    f"(codes can change when a council is amalgamated or renamed)."
                )
                continue

            latest_year = max(series)
            out = {
                "region":   name,
                "level":    "LGA",
                "state":    state,
                "abs_lga":  code,
                "source":   f"ABS Data API, ERP by LGA ({dataflow})",
                "dataflow": dataflow,
                "year":     latest_year,
                "value":    series[latest_year],
                "series_by_year": {str(y): series[y] for y in sorted(series)},
            }
            out_path = out_dir / f"lga_{slug}_population_erp_lga.json"
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2)

            self.log.info(
                f"  {label}: {len(series)} years "
                f"({min(series)}-{latest_year}), {latest_year} = {series[latest_year]:,}"
            )
            self.result.towns_ok.append(label)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = ABSPopulationERPLGAFetcher().run()
    sys.exit(0 if result.success else 1)
