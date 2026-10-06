"""
regional-indicators/fetchers/fetch_population_erp_lga.py
-----------------------------------------------------------
Fetches the LGA-level "Population (ERP)" figure -- the "Residents (LGA)"
row in each block of the Population sheet's LGA section (Brisbane,
Goondiwindi, Isaac, Maranoa, Narrabri, Toowoomba, Western Downs).

Built 2026-10-06. Before this, NOTHING fetched or wrote LGA-level ERP:
fetch_population_erp.py only ever asks QRSIS for the SA2 region type,
and its writer only writes the SA2 section, so the whole LGA
"Population (ERP)" row (Population!AA4, AA6, AA9, AA12, AA14, AA17,
AA20 for 2025) stayed empty on every run.

SOURCE: ABS Data API, dataflow "ERP by LGA (<year>), 2001 to <year>"
(id ERP_LGA<year>, e.g. ERP_LGA2025). One national dataset, every LGA
in Australia, full annual history back to 2001 on the current boundary
vintage.

WHY ABS AND NOT QRSIS (both were tested live, 2026-10-06):
  - QRSIS collection 1961 does offer an "LGA - Local Government Area"
    region type and returns the six Queensland LGAs correctly. But it
    is Queensland-only, so Narrabri (NSW) would need a second source
    anyway.
  - The ABS dataflow returns all seven -- including Narrabri -- from
    one query, and its Queensland figures are IDENTICAL to QRSIS's
    (QGSO republishes the ABS series).
  - It needs no multi-step wizard session, and it already covers every
    other state and territory, so adding an NT or WA LGA later is a
    towns.toml edit, not new code.

CONFIRMED LIVE 2026-10-06 against the hand-built 2026 answer key
(Indicators Data-Charts 2026.xlsx): all seven LGAs, every year
2001-2025 present in the workbook, zero mismatches.
  2025: Brisbane 1,375,301 / Goondiwindi 10,438 / Isaac 23,186 /
        Maranoa 13,370 / Narrabri 12,797 / Toowoomba 186,276 /
        Western Downs 35,452

API details, all confirmed live (same host and conventions as
fetch_narrabri_approvals.py):
  - Host: https://data.api.abs.gov.au/rest/
  - A new ERP_LGA<year> dataflow appears each year, on that year's LGA
    boundaries. This fetcher lists the dataflows live and uses the
    newest, so it picks up ERP_LGA2026 next year with no code change.
  - Dimension order: MEASURE . REGION_TYPE . REGION . FREQ
    Query key used: "ERP..<code>+<code>+....A" -- REGION_TYPE is left
    blank (wildcard) on purpose, because its code is vintage-specific
    ("LGA2025" this year) and would otherwise need updating annually.
  - Values are persons at 30 June of TIME_PERIOD (the workbook's
    calendar-year header: 2025 = the "2024/25" column).
  - A code that doesn't exist in the dataflow is silently omitted from
    the response rather than raising -- handled below as a per-region
    failure so a bad towns.toml code is reported, not ignored.

WHICH LGAs: derived from the towns in towns.toml (config.lgas()) --
every distinct LGA across all towns, BENCHMARKS INCLUDED, each fetched
once however many towns share it. A town's LGA comes from its existing
fields: `lga` (the name, which must match the workbook's LGA block
label) and the code from `qgso_lga` ("LGA/37310" -> 37310; QGSO uses
the ABS code). Non-Queensland towns have no qgso_lga, so they carry
`lga_code` instead (Narrabri: "15750"). Brisbane is an ordinary
towns.toml entry with benchmark = true.
Nothing is hardcoded here -- to add an LGA, add or edit a town.

OUTPUT: cache/population/lga_<slug>_population_erp_lga.json, one per
LGA:
  {"region": "Brisbane", "level": "LGA", "state": "QLD",
   "abs_lga": "31000", "source": "...", "dataflow": "ERP_LGA2025",
   "year": 2025, "value": 1375301,
   "series_by_year": {"2001": 885787, ..., "2025": 1375301}}
Read by transform/xlsx_update/update_population_erp_lga.py.

(The filename ends "_population_erp_lga.json", which the SA2 writer's
"*_population_erp.json" glob does NOT match -- the two never collide.)

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
    """Live-list ABS dataflows and return the newest ERP_LGA<year> id.

    Matches ERP_LGA followed by exactly four digits and nothing else, so
    the similarly named ABS_ERP_LGA2018, ABS_ANNUAL_ERP_LGA2025 (age and
    sex breakdown) and ERP_COMP_LGA2025 (components of change) dataflows
    are NOT picked up -- confirmed live, all of those exist alongside.
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
    """SDMX-CSV -> {lga_code: {year: persons}}. Column names confirmed
    live: REGION, TIME_PERIOD, OBS_VALUE."""
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
    SUPPORTED_STATES = []   # national dataset -- any state/territory

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
