"""
fetchers/fetch_narrabri_approvals.py
------------------------------------
Fetches Building Approvals for Narrabri (LGA), for the Housing sheet's
separate Narrabri block (rows 142-145: "Building approvals: new
residential building" / "Building approvals: new houses"). Confirmed via
direct inspection of the real starting file that this block has its own
distinct structure from the rest of Housing (QLD LGA/SA2/State), last
populated through 2023, and was never previously automated -- sourced
from a manually-downloaded "Narrabri housing 2024.xlsx" file, per TODO.md.

SOURCE: ABS's SDMX REST API (free, no key needed), NOT QGSO -- QGSO is
Queensland-government-specific and doesn't cover NSW. Confirmed live,
2026-09-30/10-01:
  - Correct current API host is https://data.api.abs.gov.au/rest/ -- the
    OLDER https://api.data.abs.gov.au/ (seen in some older docs and
    libraries) doesn't resolve at all anymore (DNS failure, confirmed).
  - ABS publishes Building Approvals LGA data as a SEPARATE dataflow per
    financial year and ASGS boundary vintage: BA_LGA2018 through
    BA_LGA2026 confirmed live via the dataflow listing (discovered
    dynamically below, not hardcoded -- same self-updating principle as
    the rest of this project's "update each cycle" fixes).
  - Each BA_LGAxxxx dataflow covers ONE financial year (Jul-Jun) on that
    year's LGA boundaries, confirmed directly: BA_LGA2024 = exactly
    2024-07 through 2025-06 (12 months), BA_LGA2025 similarly. A complete
    CALENDAR year therefore needs TWO adjacent dataflow versions merged
    (e.g. 2024's Jan-Jun comes from the PRIOR FY's dataflow, Jul-Dec from
    the matching one) -- handled below by fetching every available
    dataflow and merging by month before aggregating to calendar years,
    rather than trying to map "which FY dataflow has which half" by hand.
  - Dimension order confirmed from the dataflow's own structure document:
    MEASURE.SECTOR.WORK_TYPE.BUILDING_TYPE.REGION_TYPE.REGION.FREQ.
    Codes confirmed from the structure's codelists, not guessed:
      MEASURE=1 (Number of dwelling units), SECTOR=1 (Private Sector),
      WORK_TYPE=1 (New), FREQ=M (Monthly).
    BUILDING_TYPE: 110=Houses (matches "new houses" exactly); 100=Total
      Residential (matches "new residential building" -- confirmed by the
      same QLD-vs-"Total (Number)" pattern already proven correct for the
      QLD approvals fix earlier this session: the sheet's "new residential
      building" label means ALL dwelling types combined, not a narrower
      category).
    Narrabri's REGION code (15750) confirmed from BA_LGA2025's own
    codelist, and confirmed working against BA_LGA2024's data too --
    looked up fresh per dataflow below rather than assumed stable forever,
    in case ABS ever reclassifies LGA boundaries.

KNOWN LIMITATION, confirmed live and not chased further since it doesn't
block the actual goal: BA_LGA2018/2019/2020 don't resolve Narrabri in
their own codelists (likely a pre-2021 ASGS-edition boundary naming
difference -- 2021 was a major ASGS edition). This means calendar year
2021 drops out entirely (its Jan-Jun half needs BA_LGA2020, which fails),
and full history only reliably assembles from 2022 onward. Not a problem
for this fetcher's actual purpose -- the sheet already has real data back
to 2010 from the original manual source; this only needs to ADD 2024
onward, which it does. BA_LGA2026 also 404s on the data query despite
appearing in the dataflow listing -- confirmed live, the dataflow is
registered but not yet populated (expected for a just-opened FY).

Output: calendar-year sums (not means) for both measures, complete years
only -- matches the same "sum of real monthly approval counts" principle
already proven correct for QLD's approvals fix. Schema matches this
project's per-town JSON convention; written into
cache/housing/regions/narrabri_approvals.json for update_housing.py's
Narrabri-specific writer to pick up (that writer, a separate piece of
work, reads this file directly -- see update_housing.py's Narrabri
handling).
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
# update each cycle if this ever stops matching -- confirmed live
# 2026-10-01 this ranges BA_LGA2018 through BA_LGA2026; discovered fresh
# below via the dataflow listing, this is only documentation of what was
# true when built, not a value the code actually depends on.
DATAFLOW_PREFIX = "BA_LGA"

NARRABRI_NAME = "Narrabri"
BUILDING_TYPES = {
    "110": "new_houses",            # matches Housing sheet row 144/145 label exactly
    "100": "new_residential_building",  # matches row 142/143 -- "Total Residential" = all dwelling types
}
OUTPUT_LABELS = {
    "new_houses": "Building approvals: new houses",
    "new_residential_building": "Building approvals: new  residential building",  # double space matches the real sheet label exactly, confirmed by direct inspection
}


def _discover_lga_dataflows(log) -> list[str]:
    """Live-list every BA_LGA<year> dataflow ABS currently publishes,
    rather than hardcode a year range -- confirmed live 2026-10-01 these
    run BA_LGA2018 through BA_LGA2026, each a separate financial year +
    boundary-vintage dataflow. Returns them sorted oldest to newest.
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
    """Look up Narrabri's region code fresh from this specific dataflow's
    own codelist, rather than assume a hardcoded code stays valid across
    every boundary vintage forever. Confirmed live: 15750 works for both
    2024 and 2025 vintages, but this still looks it up live rather than
    hardcoding that number, matching this project's "verify live, fall
    back to the proven value" pattern used for the other self-updating
    fetchers.

    BUG FOUND AND FIXED 2026-10-01: the first version of this function
    requested the structure as raw XML and regex-matched it -- found
    nothing, on every single dataflow. Confirmed live that asking for
    JSON via the Accept header instead (same approach already proven for
    the main data query) returns a clean, parseable codelist -- the XML
    endpoint may format things differently or need a different query
    param than guessed; JSON is simply the approach already confirmed
    working elsewhere in this file, so this now matches that rather than
    inventing a second, untested parsing path.
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

            # Build TIME_PERIOD index -> month string, and series index -> building type code
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
                # BUILDING_TYPE is the 4th dimension (index 3) per the confirmed order above
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

        # Aggregate by calendar year -- sum of real monthly counts, complete years only
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
                "Calendar-year sums of real monthly approval counts, complete "
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
