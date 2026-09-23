"""
fetchers/fetch_income_table6.py
--------------------------------
Fetches ATO individual income data (Table 6) from data.gov.au.

Table 6 is published annually with two sheets:
  Table 6A — rows split by taxable status (Non Taxable / Taxable)
  Table 6B — combined totals (both statuses merged, no status column)

We use:
  Table 6A (Taxable rows only) -> avg_income_taxable
    = taxable_income_$ / taxable_income_no.
  Table 6B (combined) -> avg_income_all, earners_no, wages_total,
    abn_npp_income_$/no., abn_pp_income_$/no., abn_total_income_$/no.

Confirmed column indices (2023-24 file, verified 2026-09-23 against
the real downloaded xlsx, not assumed from the 2022-23 file's indices
which may not carry forward year to year):
  6A: 0=status  3=postcode  5=taxable_income_no  6=taxable_income_$
      19=wages_no  20=wages_$
  6B: 2=postcode  4=taxable_income_no  5=taxable_income_$
      18=wages_no  19=wages_$
      132=abn_pp_income_no    133=abn_pp_income_$
      134=abn_npp_income_no   135=abn_npp_income_$
      136=abn_total_income_no 137=abn_total_income_$

avg_income_all METHODOLOGY CORRECTED 2026-09-23, per Notes_DD.docx
(Steve's documented process for how the 2026 reference workbook was
actually built): "Average Taxable Income or Loss (all individuals)"
is Table 6B's Taxable income or loss $ / Taxable income or loss no. --
NOT fetch_income.py's Table 8 "Average taxable income" column, which
is a confirmed DIFFERENT ATO product with a real, non-trivial
discrepancy (Chinchilla 2023-24: Table 8 gives $72,289, but the
documented Table 6B method gives $73,581, matching the reference
workbook exactly -- verified directly against the real downloaded
file, not just trusting the notes). Table 8 remains useful for
backfilling older years Table 6 doesn't cover in a single release
(Table 6 is a single-year snapshot each cycle, Table 8 spans multiple
non-contiguous years) -- update_income.py decides which source to use
per year, not this fetcher.

ABN / business income fields, per Notes_DD.docx's mapping table --
the ATO file doesn't use the literal words "Individual ABN":
  Individual ABN NPP Total Income $/no.  -> Total business income,
    non-primary production $/no.
  Individual ABN PP Total Income $/no.   -> Total business income,
    primary production $/no.
  Individual ABN Total Income $/no.      -> Total business income $/no.
  (NPP = non-primary production, PP = primary production)
The notes don't explicitly say 6A vs 6B for these three field pairs --
using 6B (unfiltered, all individuals), consistent with how wages
already comes from 6B and none of the six ABN row labels carry a
"(taxable individuals)" qualifier the way the two average-income rows
do. Worth Steve confirming this is the intended reading.

Confirmed real values, Chinchilla (4413) 2023-24 (verified directly
against the real downloaded file, 2026-09-23):
  avg_income_all     = $73,581   (exact match vs the 2026 reference workbook)
  avg_income_taxable = $91,431   (exact match, was already correct before this change)

Website CSVs produced:
  Income - for taxable individuals.csv
  Income - all individuals.csv
  Number of earners.csv
  Wage & salary earnings (town total).csv
  Individual ABN NPP Total Income.csv
  Individual ABN PP Total Income.csv
  Individual ABN Total Income.csv
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from fetchers.ato_release import discover_latest_ato_release

try:
    import openpyxl
    import requests
except ImportError:
    raise ImportError("pip install openpyxl requests")


CKAN_API = "https://data.gov.au/data/api/3/action/package_show"


def _score_resource(name: str) -> int:
    n = name.lower()
    score = 0
    if "table" in n and ("6" in n or "06" in n): score += 3
    if "postcode" in n:                           score += 2
    if "individual" in n:                         score += 1
    if "selected" in n and "item" in n:           score += 1
    if "taxable" in n and "status" in n:          score += 1
    return score


class ATOTable6Fetcher(BaseFetcher):

    SOURCE_NAME      = "ato_income_table6"
    SUPPORTED_STATES = []

    def fetch_all(self):
        try:
            release = discover_latest_ato_release()
        except RuntimeError as exc:
            self.result.add_error("ALL", str(exc))
            return

        year = release.financial_year
        slug = release.slug

        self.log.info(
            f"Using ATO Taxation Statistics {year} "
            f"(package modified {release.modified or 'unknown'})"
        )

        try:
            resp = requests.get(CKAN_API, params={"id": slug}, timeout=30)
            resp.raise_for_status()
            resources = resp.json().get("result", {}).get("resources", [])
        except Exception as exc:
            self.log.error(f"CKAN API unreachable: {exc}")
            self.result.add_error("ALL", f"CKAN API unreachable: {exc}")
            return

        self.log.info(f"  Dataset has {len(resources)} resources")
        best = max(resources, key=lambda r: _score_resource(r.get("name", "")), default=None)
        if not best or _score_resource(best.get("name", "")) == 0:
            self.log.error("Could not identify Table 6 resource")
            return

        t6_url = best["url"]
        self.log.info(f"  Table 6 URL: {t6_url}")

        cache_key = f"ato_table6_{year}"
        t6_path   = self.download(t6_url, cache_key, suffix=".xlsx")
        if not t6_path:
            self.result.add_error("ALL", "Table 6 download failed")
            return

        self.log.info(f"  Parsing {t6_path.name} ({t6_path.stat().st_size // 1024} KB)")

        taxable_data  = self._parse_6a_taxable(t6_path)
        combined_data = self._parse_6b_combined(t6_path)

        if not taxable_data or not combined_data:
            self.result.add_error("ALL", "Table 6 parse returned no data")
            return

        for town in self.applicable_towns():
            self._extract_town(town, taxable_data, combined_data, year)

        self._compute_state_benchmarks(t6_path, year)

    def _compute_state_benchmarks(self, path: Path, year: str):
        """QLD and NSW state-level benchmark figures, per Notes_DD.docx's
        documented process: NOT an average of postcode averages -- sum
        the raw dollar and count fields across every row belonging to
        that state, then divide. Confirmed real state codes in the file
        (2026-09-23): QLD, NSW, VIC, WA, SA, TAS, ACT, NT, Overseas.
        Verified directly against the real downloaded file before
        building this: computed QLD/NSW all/taxable all four exactly
        match the 2026 reference workbook's existing benchmark figures
        (75,890 / 90,846 / 83,306 / 100,533).
        """
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            ws_a = wb["Table 6A"]
            ws_b = wb["Table 6B"]

            def state_avg_all(state_code: str) -> int | None:
                total_no, total_dollars = 0.0, 0.0
                for row in ws_b.iter_rows(min_row=3, values_only=True):
                    if row[0] == state_code:
                        total_no += (row[4] or 0)
                        total_dollars += (row[5] or 0)
                return round(total_dollars / total_no) if total_no > 0 else None

            def state_avg_taxable(state_code: str) -> int | None:
                total_no, total_dollars = 0.0, 0.0
                for row in ws_a.iter_rows(min_row=3, values_only=True):
                    status = str(row[0] or "").strip().lower()
                    if "taxable" not in status or "non" in status:
                        continue
                    if row[1] == state_code:
                        total_no += (row[5] or 0)
                        total_dollars += (row[6] or 0)
                return round(total_dollars / total_no) if total_no > 0 else None

            fy_parts = year.replace("\u2013", "-").split("-")
            cal_year = str(int(fy_parts[0]) + 1) if len(fy_parts) == 2 else year

            benchmarks = {}
            for state_code in ("QLD", "NSW"):
                avg_all = state_avg_all(state_code)
                avg_taxable = state_avg_taxable(state_code)
                benchmarks[state_code] = {
                    "avg_income_all":     {cal_year: avg_all},
                    "avg_income_taxable": {cal_year: avg_taxable},
                }
                self.log.info(
                    f"  Benchmark {state_code}: avg_all=${avg_all:,}  avg_taxable=${avg_taxable:,}"
                    if avg_all and avg_taxable else f"  Benchmark {state_code}: no data"
                )

            out = {
                "source": "ATO Taxation Statistics Table 6 (state benchmark)",
                "latest_year": year,
                "cal_year": cal_year,
                "benchmarks": benchmarks,
            }
            out_dir = Path(__file__).parent.parent / "cache" / "ato"
            out_dir.mkdir(parents=True, exist_ok=True)
            out_path = out_dir / "benchmark_income_t6.json"
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2, default=str)

        except Exception as exc:
            self.log.error(f"State benchmark computation error: {exc}", exc_info=True)

    def _parse_6a_taxable(self, path: Path) -> dict:
        """Table 6A -- Taxable rows only. Returns {pc: {taxable_no, taxable_income}}"""
        try:
            ws   = openpyxl.load_workbook(path, read_only=True)["Table 6A"]
            rows = list(ws.iter_rows(values_only=True))
            result, skipped = {}, 0
            for row in rows[2:]:
                status = str(row[0] or "").strip().lower()
                if "taxable" not in status or "non" in status:
                    continue
                pc = self._clean_pc(row[3])
                if not pc:
                    skipped += 1
                    continue
                try:
                    result[pc] = {
                        "taxable_no":     float(row[5] or 0),
                        "taxable_income": float(row[6] or 0),
                    }
                except (TypeError, ValueError):
                    skipped += 1
            self.log.info(f"    6A: {len(result)} taxable postcodes (skipped {skipped})")
            return result
        except Exception as exc:
            self.log.error(f"Table 6A parse error: {exc}", exc_info=True)
            return {}

    def _parse_6b_combined(self, path: Path) -> dict:
        """Table 6B -- combined (all-individuals) totals. Returns
        {pc: {taxable_no, taxable_income, wages_no, wages_total,
              abn_pp_no, abn_pp_income, abn_npp_no, abn_npp_income,
              abn_total_no, abn_total_income}}
        Confirmed real column indices (2026-09-23, see module docstring).
        """
        try:
            ws   = openpyxl.load_workbook(path, read_only=True)["Table 6B"]
            rows = list(ws.iter_rows(values_only=True))
            result, skipped = {}, 0
            for row in rows[2:]:
                pc = self._clean_pc(row[2])
                if not pc:
                    skipped += 1
                    continue
                try:
                    result[pc] = {
                        "taxable_no":      float(row[4] or 0),
                        "taxable_income":  float(row[5] or 0),
                        "wages_no":        float(row[18] or 0),
                        "wages_total":     float(row[19] or 0),
                        "abn_pp_no":       float(row[132] or 0),
                        "abn_pp_income":   float(row[133] or 0),
                        "abn_npp_no":      float(row[134] or 0),
                        "abn_npp_income":  float(row[135] or 0),
                        "abn_total_no":    float(row[136] or 0),
                        "abn_total_income": float(row[137] or 0),
                    }
                except (TypeError, ValueError):
                    skipped += 1
            self.log.info(f"    6B: {len(result)} postcodes (skipped {skipped})")
            return result
        except Exception as exc:
            self.log.error(f"Table 6B parse error: {exc}", exc_info=True)
            return {}

    def _clean_pc(self, raw) -> str | None:
        if raw is None:
            return None
        pc = str(int(raw)) if isinstance(raw, float) else str(raw).strip()
        pc = pc.replace(".0", "").zfill(4)
        return pc if (pc.isdigit() and len(pc) == 4) else None

    def _extract_town(self, town, taxable_data: dict, combined_data: dict, year: str):
        agg_t = {"taxable_no": 0.0, "taxable_income": 0.0}
        agg_c = {
            "taxable_no": 0.0, "taxable_income": 0.0,
            "wages_no": 0.0, "wages_total": 0.0,
            "abn_pp_no": 0.0, "abn_pp_income": 0.0,
            "abn_npp_no": 0.0, "abn_npp_income": 0.0,
            "abn_total_no": 0.0, "abn_total_income": 0.0,
        }
        found = False

        for pc in town.postcodes:
            key = pc.zfill(4)
            if key in taxable_data:
                found = True
                for k in agg_t:
                    agg_t[k] += taxable_data[key][k]
            if key in combined_data:
                for k in agg_c:
                    agg_c[k] += combined_data[key][k]

        if not found:
            self.log.warning(f"  [{town.name}] no Table 6 data for {town.postcodes}")
            self.result.towns_failed.append(town.name)
            return

        avg_taxable = (
            round(agg_t["taxable_income"] / agg_t["taxable_no"])
            if agg_t["taxable_no"] > 0 else None
        )
        avg_all = (
            round(agg_c["taxable_income"] / agg_c["taxable_no"])
            if agg_c["taxable_no"] > 0 else None
        )
        earners_no  = int(agg_c["wages_no"])
        wages_total = int(agg_c["wages_total"])
        abn_pp_no        = int(agg_c["abn_pp_no"])
        abn_pp_income     = int(agg_c["abn_pp_income"])
        abn_npp_no        = int(agg_c["abn_npp_no"])
        abn_npp_income    = int(agg_c["abn_npp_income"])
        abn_total_no      = int(agg_c["abn_total_no"])
        abn_total_income  = int(agg_c["abn_total_income"])

        fy_parts = year.replace("\u2013", "-").split("-")
        cal_year = str(int(fy_parts[0]) + 1) if len(fy_parts) == 2 else year

        out = {
            "town": town.name, "state": town.state, "postcodes": town.postcodes,
            "source": "ATO Taxation Statistics Table 6",
            "latest_year": year, "cal_year": cal_year,
            "indicators": {
                "avg_income_all":       {cal_year: avg_all},
                "avg_income_taxable":   {cal_year: avg_taxable},
                "earners_no":           {cal_year: earners_no},
                "wages_total":          {cal_year: wages_total},
                "abn_pp_income_no":     {cal_year: abn_pp_no},
                "abn_pp_income_total":  {cal_year: abn_pp_income},
                "abn_npp_income_no":    {cal_year: abn_npp_no},
                "abn_npp_income_total": {cal_year: abn_npp_income},
                "abn_total_income_no":    {cal_year: abn_total_no},
                "abn_total_income_total": {cal_year: abn_total_income},
            }
        }

        out_dir  = Path(__file__).parent.parent / "cache" / "ato"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{town.slug}_income_t6.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2, default=str)

        self.log.info(
            f"  {town.name}: avg_all=${avg_all:,}  avg_taxable=${avg_taxable:,}  "
            f"earners={earners_no:,}  wages=${wages_total:,}  "
            f"abn_total=${abn_total_income:,} ({abn_total_no:,} filers)"
            if avg_all and avg_taxable else f"  {town.name}: no data"
        )
        self.result.towns_ok.append(town.name)


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = ATOTable6Fetcher().run()
    sys.exit(0 if result.success else 1)