"""
regional-indicators/fetchers/fetch_population_erp_abs.py

Fetches SA2-level estimated resident population (ERP) for towns outside
Queensland: the "Population (ERP)" row in each town's SA2 block of the
Population sheet. Queensland towns are fetched from QGSO by
fetch_population_erp.py; this module writes the same output for the
other states, so update_population_erp.py writes both without change.

Source: ABS Data API, dataflow ABS_ANNUAL_ERP_ASGS2021 ("ERP by SA2 and
above (ASGS Edition 3), 2001 onwards"). One national dataset with annual
figures (30 June) from 2001, on 2021 SA2 boundaries. The query names the
towns' SA2 codes (sa2_code in towns.toml):
    https://data.api.abs.gov.au/rest/data/ABS,ABS_ANNUAL_ERP_ASGS2021,/
        ERP.SA2.<code>+<code>.A?startPeriod=2001&format=csvfile
Same host and conventions as fetch_population_erp_lga.py and
fetch_nsw_approvals.py.

Checked against Narrabri's hand-made row in the 2026 workbook: 2001-2016
equal; 2017-2024 differ by 1 to 29 persons, which are ABS's revisions
after the 2021 Census (the workbook holds the earlier figures); 2025 is
not in the workbook.

Output: cache/population/<slug>_population_erp.json
    {"town", "state", "source", "sa2_name", "year", "value", "series_by_year"}
as written by fetch_population_erp.py.
"""

from __future__ import annotations

import csv
import io
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR, YEAR_START, YEAR_END

API_BASE = "https://data.api.abs.gov.au/rest"
DATAFLOW = "ABS_ANNUAL_ERP_ASGS2021"
SOURCE = "ABS Data API, ERP by SA2 (ASGS Edition 3), ABS_ANNUAL_ERP_ASGS2021"


def parse(text: str) -> dict[str, dict[str, float]]:
    """{SA2 code: {year: persons}} from the API's CSV response."""
    out: dict[str, dict[str, float]] = {}
    for row in csv.DictReader(io.StringIO(text)):
        code = row.get("ASGS_2021") or row.get("REGION") or ""
        try:
            value = float(row["OBS_VALUE"])
        except (KeyError, ValueError):
            continue
        out.setdefault(code.strip(), {})[str(row["TIME_PERIOD"]).strip()] = value
    return out


class ABSPopulationERPFetcher(BaseFetcher):

    SOURCE_NAME      = "population_erp_abs"
    SUPPORTED_STATES = ["NSW", "VIC", "SA", "WA", "TAS", "NT", "ACT"]   # Queensland: fetch_population_erp.py

    def fetch_all(self):
        import requests

        towns = [t for t in self.applicable_towns() if t.sa2_code]
        if not towns:
            self.log.info("  no towns outside Queensland with an sa2_code")
            return
        codes = sorted({str(t.sa2_code) for t in towns})
        url = f"{API_BASE}/data/ABS,{DATAFLOW},/ERP.SA2.{'+'.join(codes)}.A"
        try:
            response = requests.get(url, params={"startPeriod": str(YEAR_START), "format": "csvfile"},
                                    headers={"Accept": "application/vnd.sdmx.data+csv"}, timeout=120)
            response.raise_for_status()
        except requests.RequestException as exc:
            self.result.add_error("ALL", f"ABS Data API query failed: {exc}")
            return
        by_code = parse(response.text)

        out_dir = CACHE_DIR / "population"
        out_dir.mkdir(parents=True, exist_ok=True)
        for town in towns:
            series = {y: int(round(v)) for y, v in sorted(by_code.get(str(town.sa2_code), {}).items())
                      if YEAR_START <= int(y) <= YEAR_END}
            if not series:
                # A warning, not a failure: a town whose sa2_code is not on 2021
                # boundaries (an older edition's code) should not fail the run.
                self.result.add_warning(town.name, f"SA2 {town.sa2_code} is not in {DATAFLOW} (2021 "
                                                   f"boundaries); check sa2_code in towns.toml")
                continue
            latest = max(series, key=int)
            out = {
                "town":     town.name,
                "state":    town.state,
                "source":   SOURCE,
                "sa2_name": town.sa2_name or town.name,
                "year":     int(latest),
                "value":    series[latest],
                "series_by_year": series,
            }
            with open(out_dir / f"{town.slug}_population_erp.json", "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2)
            self.log.info(f"  {town.name} (SA2 {town.sa2_code}): {min(series)}-{latest}, {latest} = {series[latest]:,}")
            self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = ABSPopulationERPFetcher().run()
    sys.exit(0 if result.success else 1)
