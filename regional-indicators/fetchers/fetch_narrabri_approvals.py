"""
regional-indicators/fetchers/fetch_narrabri_approvals.py

Fetches building approvals for Narrabri (LGA) for the Narrabri block of
the Housing sheet (rows 142-145: "Building approvals: new residential
building" and "Building approvals: new houses").

The Narrabri block has its own structure, separate from the Queensland
LGA/SA2/State sections of the Housing sheet. Before this fetcher it was
filled by hand from a manually downloaded "Narrabri housing 2024.xlsx"
file and was populated to 2023.

Source: ABS SDMX REST API (free, no key). QGSO covers Queensland only.
  - The API host is https://data.api.abs.gov.au/rest/. The older
    https://api.data.abs.gov.au/, which appears in some documentation
    and libraries, no longer resolves.
  - Building Approvals LGA data is published as a separate dataflow per
    financial year and ASGS boundary vintage (BA_LGA2018 to BA_LGA2026
    at the time of writing). The list is read from the dataflow listing
    at run time, not hardcoded.
  - Each BA_LGA<year> dataflow covers one financial year (July to June)
    on that year's LGA boundaries; BA_LGA2024 holds 2024-07 to 2025-06.
    A complete calendar year therefore spans two adjacent dataflows
    (January-June from the earlier one, July-December from the later).
    Every available dataflow is fetched and merged by month before
    aggregating to calendar years.
  - Dimension order, from the dataflow's structure document:
    MEASURE.SECTOR.WORK_TYPE.BUILDING_TYPE.REGION_TYPE.REGION.FREQ.
    Codes, from the structure's codelists:
      MEASURE=1 (Number of dwelling units), SECTOR=1 (Private Sector),
      WORK_TYPE=1 (New), FREQ=M (Monthly).
    BUILDING_TYPE: 110 = Houses ("new houses"); 100 = Total Residential
      ("new residential building"). The sheet's "new residential
      building" label means all dwelling types combined, as with the
      "Total (Number)" series used for the Queensland approvals.
    REGION: Narrabri is 15750 in the BA_LGA2024 and BA_LGA2025
      codelists. The code is looked up per dataflow in case ABS
      reclassifies LGA boundaries.

Known limitations:
  - BA_LGA2018, BA_LGA2019 and BA_LGA2020 do not resolve Narrabri in
    their codelists (probably a boundary naming difference before the
    2021 ASGS edition). Calendar year 2021 is therefore not produced
    (its January-June half needs BA_LGA2020), and history assembles
    reliably only from 2022. The sheet already holds data back to 2010
    from the original manual source, so only 2024 onward is needed.
  - BA_LGA2026 appears in the dataflow listing but its data query
    returns 404: a dataflow for a newly opened financial year is
    registered before it is populated.

Output: cache/housing/regions/narrabri_approvals.json, read by the
Narrabri writer in update_housing.py. Values are calendar-year sums (not
means) of the monthly approval counts for both measures, complete years
only, as for the Queensland approvals. The schema follows the per-town
JSON convention used elsewhere in the pipeline.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR, YEAR_START, YEAR_END

try:
    import requests
except ImportError:
    raise ImportError("pip install requests")


API_BASE = "https://data.api.abs.gov.au/rest"
# Dataflows are discovered from the listing at run time (BA_LGA2018 to
# BA_LGA2026 at the time of writing). Update each cycle only if ABS
# changes this prefix.
DATAFLOW_PREFIX = "BA_LGA"

NARRABRI_NAME = "Narrabri"
BUILDING_TYPES = {
    "110": "new_houses",            # Housing sheet rows 144/145
    "100": "new_residential_building",  # rows 142/143; "Total Residential" = all dwelling types
}
OUTPUT_LABELS = {
    "new_houses": "Building approvals: new houses",
    "new_residential_building": "Building approvals: new  residential building",  # the double space matches the sheet label
}


def _discover_lga_dataflows(log) -> list[str]:
    """Return every BA_LGA<year> dataflow that ABS currently publishes,
    sorted oldest to newest.

    The list is read from the API rather than hardcoded. Each dataflow
    covers one financial year on one boundary vintage.
    """
    try:
        resp = requests.get(
            f"{API_BASE}/dataflow/ABS",
            params={"detail": "allstubs"},
            timeout=30,
        )
        resp.raise_for_status()
        ids = sorted(set(re.findall(rf'id="({DATAFLOW_PREFIX}\d+)"', resp.text)))
        log.info(f"  Discovered {len(ids)} Building Approvals LGA dataflows: {ids[0]}..{ids[-1]}")
        return ids
    except requests.RequestException as exc:
        log.error(f"  Could not list ABS dataflows: {exc}")
        return []


def _find_narrabri_code(log, dataflow_id: str, region_type: str) -> str | None:
    """Return Narrabri's region code from the given dataflow's codelist,
    or None if it is not found.

    The code is looked up per dataflow rather than hardcoded, because it
    is not guaranteed to stay the same across boundary vintages (it is
    15750 in both the 2024 and 2025 vintages).

    The structure is requested as JSON through the Accept header, as for
    the data query. Matching the raw XML structure response does not
    find the codelist.
    """
    region_codelist = f"CL_LGA_{region_type[3:]}"   # e.g. "LGA2025" -> "CL_LGA_2025"
    try:
        resp = requests.get(
            f"{API_BASE}/datastructure/ABS/{dataflow_id}/1.0.0",
            params={"references": "codelist"},
            headers={"Accept": "application/vnd.sdmx.structure+json"},
            timeout=30,
        )
        resp.raise_for_status()
        payload = resp.json()
        for cl in payload.get("data", {}).get("codelists", []):
            if cl["id"] == region_codelist:
                for code in cl["codes"]:
                    if code["name"] == NARRABRI_NAME:
                        return code["id"]
    except (requests.RequestException, ValueError, KeyError) as exc:
        log.warning(f"  Could not look up Narrabri's code for {dataflow_id}: {exc}")
    return None


class NarrabriApprovalsFetcher(BaseFetcher):

    SOURCE_NAME      = "narrabri_approvals"
    SUPPORTED_STATES = ["NSW"]

    def fetch_all(self):
        dataflows = _discover_lga_dataflows(self.log)
        if not dataflows:
            self.result.add_error("ALL", "Could not discover any BA_LGA dataflows")
            return

        # month -> {building_type_key: value}
        by_month: dict[str, dict[str, float]] = {}

        for flow_id in dataflows:
            m = re.match(rf"{DATAFLOW_PREFIX}(\d+)", flow_id)
            boundary_year = m.group(1)
            region_type = f"LGA{boundary_year}"

            code = _find_narrabri_code(self.log, flow_id, region_type)
            if not code:
                self.log.warning(f"  [{flow_id}] Narrabri not found in this vintage's region list -- skipping")
                continue

            key = f"1.1.1.{'+'.join(BUILDING_TYPES)}.{region_type}.{code}.M"
            try:
                resp = requests.get(
                    f"{API_BASE}/data/ABS,{flow_id},1.0.0/{key}",
                    params={"format": "jsondata"},
                    timeout=60,
                )
                resp.raise_for_status()
                payload = resp.json()
            except (requests.RequestException, ValueError) as exc:
                self.log.warning(f"  [{flow_id}] Query failed: {exc}")
                continue

            data = payload.get("data", {})
            datasets = data.get("dataSets", [])
            structs = data.get("structures", [])
            if not datasets or not structs:
                continue

            # Map the TIME_PERIOD index to a month string, and the series index
            # to a building type code.
            time_values, btype_values = None, None
            for s in structs:
                for scope, dl in s.get("dimensions", {}).items():
                    for dim in dl:
                        if dim["id"] == "TIME_PERIOD":
                            time_values = [v["id"] for v in dim["values"]]
                        if dim["id"] == "BUILDING_TYPE":
                            btype_values = [v["id"] for v in dim["values"]]
            if time_values is None:
                continue

            series_dict = datasets[0].get("series", {})
            for series_key, series_val in series_dict.items():
                parts = series_key.split(":")
                # BUILDING_TYPE is the 4th dimension (index 3) of the series key.
                btype_idx = int(parts[3])
                btype_code = btype_values[btype_idx] if btype_values else None
                type_key = BUILDING_TYPES.get(btype_code)
                if not type_key:
                    continue
                for obs_idx, obs_val in series_val.get("observations", {}).items():
                    month = time_values[int(obs_idx)]
                    v = obs_val[0] if obs_val else None
                    if v is None:
                        continue
                    by_month.setdefault(month, {})[type_key] = v

            self.log.info(f"  [{flow_id}] {len(series_dict)} series, {len(time_values)} months parsed")

        if not by_month:
            self.result.add_error("ALL", "No Building Approvals data assembled for Narrabri across any dataflow")
            return

        # Aggregate by calendar year: sum of monthly counts, complete years only.
        by_year: dict[str, dict[str, float]] = {}
        for month, vals in by_month.items():
            yr = month[:4]
            for type_key, v in vals.items():
                by_year.setdefault(yr, {}).setdefault(type_key, []).append(v)

        indicators = {}
        for type_key, label in OUTPUT_LABELS.items():
            values = {}
            for yr, series in by_year.items():
                monthly = series.get(type_key, [])
                if len(monthly) == 12 and YEAR_START <= int(yr) <= YEAR_END:
                    values[yr] = int(sum(monthly))
            indicators[type_key] = {"label": label, "values": values}

        out = {
            "region":     NARRABRI_NAME,
            "section":    "Narrabri",
            "source":     "ABS Building Approvals, Australia (SDMX Data API) — LGA, by financial-year/boundary-vintage dataflow",
            "source_url": f"{API_BASE}/data/ABS,BA_LGA<year>,1.0.0/1.1.1.110+100.LGA<year>.15750.M",
            "note": (
                "Calendar-year sums of monthly approval counts, complete "
                "years only (12 months present) -- assembled by merging every "
                "available BA_LGA<year> dataflow version (each covers one "
                "financial year on that year's LGA boundaries) month by month, "
                "since no single dataflow spans a full calendar year on its own."
            ),
            "indicators": indicators,
        }

        out_dir = CACHE_DIR / "housing" / "regions"
        out_dir.mkdir(parents=True, exist_ok=True)
        with open(out_dir / "narrabri_approvals.json", "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2)

        for type_key, entry in indicators.items():
            vals = entry["values"]
            if vals:
                latest = max(vals, key=int)
                self.log.info(f"  [Narrabri:{type_key}] {len(vals)} years, latest ({latest}) = {vals[latest]}")
        self.result.towns_ok.append("Narrabri:approvals")


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = NarrabriApprovalsFetcher().run()
    sys.exit(0 if result.success else 1)
