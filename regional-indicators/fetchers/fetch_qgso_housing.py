"""
regional-indicators/fetchers/fetch_qgso_housing.py

Fetches housing indicators from the QGSO Regional Database (QRSIS).
Queensland only.

Source:
  https://www.qgso.qld.gov.au/statistics/queensland-regions/regional-tools-statistics/queensland-regional-database

Collections fetched (coverage as last recorded):
  1925  Residential land and dwelling sales   (Sep 2000 - Sep 2025, quarterly, SA2)
  1929  Median rent                            (Dec 1989 - Mar 2026, quarterly, SA2)
  2075  Building Approvals (Historical)        (Jul 2001 - Dec 2018, monthly)
  2031  Building Approvals (Current)           (Jan 2019 - present, monthly)

Indicators produced (per town, annual):
  housing_sales_count    Detached dwelling: number of sales
  housing_median_price   Detached dwelling: median sale price ($)
  rent_3bed_median       House - 3 bedrooms - median rent of lodgements ($/week)
  building_approvals     Residential dwelling units (Private); New Houses (Number)

The same four series are also fetched for the LGA, SA2 and State regions
that the Housing sheet lists in its own sections (see _fetch_regions).

Output:
  cache/housing/<slug>_qgso.json                 one per town
  cache/housing/regions/<section>_<slug>.json    one per region

Implementation notes:
  1. POST encoding: Oracle PL/SQL WebTK reads p_names/p_values as
     interleaved parallel arrays. The form data must be a list of
     (key, value) tuples; a dict groups all p_names before all p_values,
     which the server misreads. The _q() helper builds the list.

  2. HTML parsing: QRSIS serves HTML with unclosed <OPTION> tags.
     BeautifulSoup's html.parser merges the option texts into one
     string, so the lxml parser is required.

  3. udqctl_id extraction: the redirect URL uses the interleaved form
     ?p_names=udqctl_id&p_values=3908, not ?udqctl_id=3908.

  4. Building approvals geography: towns are queried by SA2, except the
     Toowoomba towns (TOOWOOMBA_LGA_SLUGS), which use the Toowoomba LGA
     (LGA/36910) as in previous booklets.

  5. Series names must match QRSIS exactly:
     - 'Detached dwelling: number of sales (Number)'
     - 'Detached dwelling: median sale price ($)'
     - 'House - 3 bedrooms - median rent of lodgements ($/week)'
     - 'Residential dwelling units (Private); New Houses (Number)'

  6. Annual sales count: the per-town output takes the "Year Ended
     31 Dec" value only, whereas the region-level output sums the four
     quarterly values (see _fetch_regions for the evidence). The two
     methods give different figures; the per-town method has not been
     reconciled with the region-level one.
"""

from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from fetchers.base import BaseFetcher
from config import CACHE_DIR

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    raise ImportError("pip install requests beautifulsoup4 lxml")


# ── QRSIS constants ─────────────────────────────────────────────────────────────

BASE_URL     = "https://statistics.qgso.qld.gov.au/pls/qis_public/"
PUBLIC_USER  = "edtert"
ACCESS_LEVEL = "85"
COLLGRP_ID   = "22"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": "https://statistics.qgso.qld.gov.au/",
    "Origin":  "https://statistics.qgso.qld.gov.au",
}

# Series names exactly as QRSIS returns them.
SERIES_SALES_COUNT = "Detached dwelling: number of sales (Number)"
SERIES_SALES_PRICE = "Detached dwelling: median sale price ($)"
SERIES_RENT        = "House - 3 bedrooms - median rent of lodgements ($/week)"
SERIES_APPROVALS   = "Residential dwelling units (Private); New Houses (Number)"

COLLECTIONS = {
    "sales": {
        "id":         "1925",
        "series":     [SERIES_SALES_COUNT, SERIES_SALES_PRICE],
        "from_date":  "Year Ended 30 Sep 2000",
        # Fallback To Date, used when the newest option cannot be read from
        # the Time Periods page (see _set_time_period). It must reach the
        # latest quarter the collection offers. A value that stops before
        # the December quarter leaves that year's price as the mean of 3 of
        # 4 quarters, which understates a rising market (in 2025 the
        # December quarter was the highest). With all four quarters the
        # computed 2025 price is within 0.1% of the workbook (Toowoomba LGA:
        # 693637 computed, 692825 in the workbook). Kept equal to "rent"'s
        # to_date so both collections stay in step. The collection returns
        # HTTP 500 for a to_date beyond what its own dropdown offers, so the
        # value cannot be set far ahead.
        "to_date":    "Year Ended 31 Mar 2026",   # update each cycle
        "geo":        "SA2",       # region type to select
    },
    "rent": {
        "id":         "1929",
        "series":     [SERIES_RENT],
        "from_date":  "Year Ended 31 Dec 2000",
        "to_date":    "Year Ended 31 Mar 2026",   # update each cycle; see "sales" above
        "geo":        "SA2",
    },
    "approvals_hist": {
        "id":         "2075",
        "series":     [SERIES_APPROVALS],
        "from_date":  "Jul 2001",
        "to_date":    "Dec 2018",
        "period":     "Monthly",
        "date_fmt":   "M1",
        # This collection's Time Periods page carries a hidden
        # p_concorded_data="N" default. Sending "Y" (which makes
        # _set_time_period append p_concorded_data=Y to the time-period
        # POST) empties the server's region-type list: 0 region types are
        # offered, against 8 with "N" (GCCSA/LGA/RESREG/S/SA2/SA3/SA4/SED),
        # and the fetch then fails with "No SA2 regions matched". The same
        # applies to approvals_curr (2031) below. With "N" the list has 548
        # SA2 regions, including Roma (SA2/307011176). "N" is also the
        # default in _set_time_period; it is set explicitly here because
        # the value matters.
        "concorded":  "N",
        "geo":        "SA2",
        "approvals":  True,
    },
    "approvals_curr": {
        "id":         "2031",
        "series":     [SERIES_APPROVALS],
        "from_date":  "Jan 2019",
        "to_date":    "Jan 2026",   # update each cycle
        "period":     "Monthly",
        "date_fmt":   "M1",
        "concorded":  "N",   # see approvals_hist above
        "geo":        "SA2",
        "approvals":  True,
    },
}



# Toowoomba approvals use the LGA boundary, as in previous booklets.
# All other towns use SA2. Toowoomba's LGA code is LGA/36910.
TOOWOOMBA_LGA_SLUGS = {"toowoomba", "toowoomba_central", "toowoomba_harlaxton", "toowoomba_west"}
TOOWOOMBA_LGA_CODE  = "LGA/36910"

# ── Region-level series for the Housing sheet's LGA/SA2/State sections ─────────
# Sales (1925) and rent (1929) support the region types "LGA - Local
# Government Area" and "S - State" as well as the SA2 type used for the
# per-town output. Brisbane = LGA/31000, Queensland = S/3 (the same state
# code as the Labour collection). The per-town output is SA2-level only,
# whereas the Housing sheet's "LGA" section (6 blocks) and the Queensland
# benchmark rows need LGA and State figures, which these regions supply.
#
# Building approvals (2075, 2031) are fetched for the same regions. All
# three region types are available for those collections provided
# p_concorded_data is "N" (see COLLECTIONS["approvals_hist"]).

LGA_REGIONS = {
    "Brisbane":       "LGA/31000",   # benchmark, not a project town
    "Goondiwindi":    "LGA/33610",
    "Isaac":          "LGA/33980",
    "Maranoa":        "LGA/34860",
    "Toowoomba":      "LGA/36910",
    "Western Downs":  "LGA/37310",
}

# SA2 codes as verified for the Business and Employment fetchers.
SA2_REGIONS = {
    "Broadsound-Nebo":             "312011338",
    "Chinchilla":                  "307011172",
    "Goondiwindi":                 "307011173",
    "Miles-Wandoan":               "307011175",
    "Moranbah":                    "312011341",
    "North Toowoomba - Harlaxton": "317011454",
    "Roma":                        "307011176",
    "Roma Surrounds":              "307011177",
    "Tara":                        "307011178",
    "Toowoomba - Central":         "317011456",
    "Toowoomba - East":            "317011457",
    "Toowoomba - West":            "317011458",
    "Wambo":                       "307021183",
}

STATE_REGIONS = {"Queensland": "S/3"}

REGION_SERIES_LABELS = {
    SERIES_SALES_PRICE: "Detached dwelling: median sale price ($)",
    SERIES_SALES_COUNT: "Detached dwelling: number of sales (Number)",
    SERIES_RENT:        "House - 3 bedrooms - median rent of lodgements ($/week)",
    # Region-level approvals. With p_concorded_data="N" the approvals
    # collections offer the LGA, SA2 and State region types (80 LGA regions,
    # including Brisbane and Toowoomba, and S/3 Queensland). The data is
    # monthly, so _fetch_regions sums it in a separate branch from the
    # quarterly logic used for the other three series.
    SERIES_APPROVALS:   "Building Approvals: Residential dwelling units (Private) New Houses (Number)",
}

def _q(*pairs) -> list[tuple]:
    """Build the interleaved p_names/p_values list from alternating
    (name, value) arguments.
    """
    result = []
    it = iter(pairs)
    for name in it:
        value = next(it)
        result.append(("p_names", name))
        result.append(("p_values", value))
    return result


def _parse_options(html: str, select_name: str) -> list[str]:
    """Return the <option> texts of the named <select>.

    The lxml parser is required: html.parser merges unclosed <OPTION>
    tags into one string.
    """
    soup = BeautifulSoup(html, "lxml")
    for sel in soup.find_all("select", {"name": select_name}):
        opts = [o.get_text(strip=True) for o in sel.find_all("option")]
        return [o for o in opts if o]
    return []


class QGSOHousingFetcher(BaseFetcher):

    SOURCE_NAME      = "qgso_housing"
    SUPPORTED_STATES = ["QLD"]

    def fetch_all(self):
        towns = self.applicable_towns()
        if not towns:
            self.log.info("No QLD towns configured — nothing to fetch")
            return

        # SA2 map used for sales and rent.
        sa2_map: dict[str, list] = {}  # sa2_code -> [town, ...]; several towns may share an SA2
        for town in towns:
            sa2 = getattr(town, "qgso_sa2", None) or town.sa2_code
            if sa2:
                sa2_map.setdefault(str(sa2), []).append(town)

        if not sa2_map:
            self.result.add_error("ALL", "No SA2 codes available for QGSO lookup")
            return

        # Building approvals: most towns use SA2 (Roma, Chinchilla, etc.);
        # Toowoomba uses LGA, as in previous booklets.
        approvals_sa2_map: dict[str, list] = {}
        approvals_lga_map: dict[str, list] = {}
        for town in towns:
            if town.slug in TOOWOOMBA_LGA_SLUGS:
                approvals_lga_map.setdefault(TOOWOOMBA_LGA_CODE, []).append(town)
            else:
                sa2 = getattr(town, "qgso_sa2", None) or town.sa2_code
                if sa2:
                    approvals_sa2_map.setdefault(str(sa2), []).append(town)
                    # Towns that share an SA2 (Roma and Wallumbilla, Miles and
                    # Wandoan) are both added, so both receive the data.

        all_data: dict[str, dict] = {}

        for coll_key, cfg in COLLECTIONS.items():
            self.log.info(f"  Collection: {coll_key} (id={cfg['id']})")
            if cfg.get("approvals"):
                # Approvals use the SA2 and LGA region maps together.
                region_map = {**{k: v for k, v in approvals_sa2_map.items()},
                              **{k: v for k, v in approvals_lga_map.items()}}
            else:
                region_map = sa2_map  # values are already lists of towns
            if not region_map:
                self.log.warning(f"  No regions for {coll_key} — skipping")
                continue
            try:
                coll_data = self._fetch_collection(coll_key, cfg, region_map)
                for slug, indicators in coll_data.items():
                    if slug not in all_data:
                        all_data[slug] = {}
                    for ind_name, values in indicators.items():
                        if ind_name in all_data[slug] and isinstance(values, dict):
                            all_data[slug][ind_name].update(values)
                        else:
                            all_data[slug][ind_name] = values
            except Exception as exc:
                self.log.error(f"  Collection {coll_key} failed: {exc}", exc_info=True)
                self.result.add_error("ALL", f"Collection {coll_key} failed: {exc}")

        out_dir = CACHE_DIR / "housing"
        out_dir.mkdir(exist_ok=True)

        for town in towns:
            indicators = all_data.get(town.slug, {})
            if not indicators:
                self.log.warning(f"  [{town.name}] no housing data retrieved")
                self.result.towns_failed.append(town.name)
                continue

            out = {
                "town":       town.name,
                "state":      town.state,
                "source":     "QGSO Regional Database (QRSIS)",
                "source_url": (
                    "https://www.qgso.qld.gov.au/statistics/queensland-regions/"
                    "regional-tools-statistics/queensland-regional-database"
                ),
                "note": (
                    "Sales count = Dec-quarter rolling 12-month total. "
                    "Median price and rent = mean of 4 quarterly medians. "
                    "Building approvals = sum of 12 monthly values (LGA level)."
                ),
                "indicators": indicators,
            }

            out_path = out_dir / f"{town.slug}_qgso.json"
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2)

            summary = ", ".join(f"{k}({len(v)}yr)" for k, v in indicators.items())
            self.log.info(f"  {town.name}: {summary}")
            self.result.towns_ok.append(town.name)

        # Region-level output for the Housing sheet's LGA/SA2/State
        # sections. It is written separately from the per-town files above
        # and covers sales, price, rent and building approvals.
        try:
            self._fetch_regions()
        except Exception as exc:
            self.log.error(f"  Region-level fetch failed: {exc}", exc_info=True)
            self.result.add_error("REGIONS", f"Region-level fetch failed: {exc}")

    def _fetch_regions(self):
        """Fetch the region-level series and write one JSON file per region
        into cache/housing/regions/.

        Sales (1925), rent (1929) and building approvals (2075, 2031) are
        each queried at the LGA, SA2 and State region types together, in
        one session per collection. The output uses the same schema as the
        Employment region files, so update_housing.py can use the same
        section-driven wiring.
        """
        import statistics as stats
        from config import YEAR_START, YEAR_END

        # region_map: code -> (label, section). The map is keyed by code and
        # built from each source dict separately, never by merging the dicts
        # on label: the same label can occur at two geography levels
        # ("Goondiwindi" is both an LGA and an SA2 on the sheet), and a merge
        # by label would drop one of them. Codes are unique across the three
        # dicts (LGA/..., SA2/..., S/...). by_region below is keyed by code
        # for the same reason, so the LGA and the SA2 named Goondiwindi never
        # share a bucket.
        region_map = {}
        for label, code in LGA_REGIONS.items():
            region_map[code] = (label, "LGA")
        for label, code in SA2_REGIONS.items():
            region_map[f"SA2/{code}"] = (label, "SA2")
        for label, code in STATE_REGIONS.items():
            region_map[code] = (label, "State")

        by_region: dict[str, dict[str, dict[int, list]]] = {}  # code -> series -> year -> {period: value}

        for coll_key in ("sales", "rent"):
            cfg = COLLECTIONS[coll_key]
            session = requests.Session()
            session.headers.update(HEADERS)
            udqctl_id = self._select_collection(session, cfg["id"])
            if not udqctl_id:
                raise RuntimeError(f"Failed to get udqctl_id for collection {cfg['id']}")

            avail_series = self._get_options(session, "QIS1110W$UDQSER.ProcessSeries",
                                              udqctl_id, "infoser.htm", "p_new_multi")
            wanted = [s for s in cfg["series"] if s in REGION_SERIES_LABELS]
            matched = self._match_series(wanted, avail_series)
            if not matched:
                raise RuntimeError(f"No region series matched for collection {cfg['id']}")
            time_periods_html = self._select_series(session, udqctl_id, matched)
            self._set_time_period(session, udqctl_id, cfg["id"], cfg, time_periods_html)

            self._select_region_type(session, udqctl_id, "LGA - Local Government Area")
            self._select_region_type(session, udqctl_id, "SA2 - Statistical Area Level 2")
            self._select_region_type(session, udqctl_id, "S - State")

            avail_regions = self._get_options(session, "QIS1110W$UDQREG.ProcessRegions",
                                               udqctl_id, "inforeg.htm", "p_new_multi")
            wanted_codes = list(region_map.keys())
            matched_regions = [r for r in avail_regions if any(r.startswith(c) for c in wanted_codes)]
            if not matched_regions:
                raise RuntimeError(f"No regions matched for collection {cfg['id']}")
            self._select_regions(session, udqctl_id, matched_regions)

            html = self._submit_report(session, udqctl_id, cfg["id"])
            raw = self._parse_output_html(html)
            self.log.info(f"  Region fetch {coll_key}: {len(raw)} regions parsed")

            for region_code, periods in raw.items():
                match = next(((c, lbl_sec) for c, lbl_sec in region_map.items()
                              if region_code.startswith(c)), None)
                if not match:
                    continue
                code, _ = match
                for period, vals in periods.items():
                    yr = self._parse_year(period)
                    if not yr or not (YEAR_START <= yr <= YEAR_END):
                        continue
                    for series_name, v in vals.items():
                        if v is None or series_name not in REGION_SERIES_LABELS:
                            continue
                        # Each value is stored under its period label, not in a
                        # flat list, so the number of quarters present in a year
                        # can be checked when the annual sales count is built.
                        by_region.setdefault(code, {}).setdefault(series_name, {}) \
                                 .setdefault(yr, {})[period] = v

        # Building approvals are fetched in a separate pass because the data
        # is monthly (parsed with _parse_month_year and summed, unlike the
        # quarterly series above) and comes from two collections (2075:
        # 2001-2018, 2031: 2019 onward) that are merged into one continuous
        # series per region, as _aggregate() does for the per-town output.
        for coll_key in ("approvals_hist", "approvals_curr"):
            cfg = COLLECTIONS[coll_key]
            session = requests.Session()
            session.headers.update(HEADERS)
            udqctl_id = self._select_collection(session, cfg["id"])
            if not udqctl_id:
                raise RuntimeError(f"Failed to get udqctl_id for collection {cfg['id']}")

            avail_series = self._get_options(session, "QIS1110W$UDQSER.ProcessSeries",
                                              udqctl_id, "infoser.htm", "p_new_multi")
            matched = self._match_series(cfg["series"], avail_series)
            if not matched:
                raise RuntimeError(f"No series matched for collection {cfg['id']}")
            time_periods_html = self._select_series(session, udqctl_id, matched)
            self._set_time_period(session, udqctl_id, cfg["id"], cfg, time_periods_html)

            self._select_region_type(session, udqctl_id, "LGA - Local Government Area")
            self._select_region_type(session, udqctl_id, "SA2 - Statistical Area Level 2")
            self._select_region_type(session, udqctl_id, "S - State")

            avail_regions = self._get_options(session, "QIS1110W$UDQREG.ProcessRegions",
                                               udqctl_id, "inforeg.htm", "p_new_multi")
            wanted_codes = list(region_map.keys())
            matched_regions = [r for r in avail_regions if any(r.startswith(c) for c in wanted_codes)]
            if not matched_regions:
                raise RuntimeError(f"No regions matched for collection {cfg['id']}")
            self._select_regions(session, udqctl_id, matched_regions)

            html = self._submit_report(session, udqctl_id, cfg["id"])
            raw = self._parse_output_html(html)
            self.log.info(f"  Region fetch {coll_key}: {len(raw)} regions parsed")

            for region_code, periods in raw.items():
                match = next(((c, lbl_sec) for c, lbl_sec in region_map.items()
                              if region_code.startswith(c)), None)
                if not match:
                    continue
                code, _ = match
                for period, vals in periods.items():
                    parsed = self._parse_month_year(period)
                    if not parsed:
                        continue
                    yr, _mo = parsed
                    if not (YEAR_START <= yr <= YEAR_END):
                        continue
                    v = vals.get(SERIES_APPROVALS)
                    if v is None:
                        continue
                    by_region.setdefault(code, {}).setdefault(SERIES_APPROVALS, {}) \
                             .setdefault(yr, {})[period] = v

        out_dir = CACHE_DIR / "housing" / "regions"
        out_dir.mkdir(parents=True, exist_ok=True)
        for code, series_data in by_region.items():
            label, section = region_map[code]
            indicators = {}
            for series_name, by_year in series_data.items():
                # Annual figures. Price and rent are per-quarter medians and are
                # averaged over the quarters present. Sales count and approvals
                # are counts and are summed, as described in each branch.
                if series_name == SERIES_SALES_COUNT:
                    # In this collection a "Year Ended DD Mon YYYY" period is a
                    # quarter-end reading whose value is that quarter's count,
                    # not a rolling 12-month total, so the annual figure is the
                    # sum of the four quarters. Check: Toowoomba LGA's four 2024
                    # values (3307/3325/3441/3435) sum to 13508, against 13484
                    # in the workbook (0.2%). Taking the mean of the four, or
                    # the December value alone, gives about a quarter of the
                    # workbook's figures. Years without all four quarters are
                    # omitted.
                    #
                    # Known limitation: the per-town code (_aggregate, "sales")
                    # takes only the December value and so does not follow this
                    # rule. It has not been changed because the effect on the
                    # per-town output has not been verified.
                    values = {
                        str(yr): int(sum(period_vals.values()))
                        for yr, period_vals in by_year.items() if len(period_vals) == 4
                    }
                elif series_name == SERIES_APPROVALS:
                    # Sum of the monthly values present for the year. Twelve
                    # months are not required: the current, in-progress year
                    # has fewer, and the per-town code applies the same rule.
                    # A partial-year count is still a valid total of approvals
                    # to date, unlike a partial-year mean.
                    values = {
                        str(yr): int(sum(period_vals.values()))
                        for yr, period_vals in by_year.items() if period_vals
                    }
                else:
                    round_fn = round if series_name == SERIES_SALES_PRICE else (lambda x: round(x, 4))
                    values = {
                        str(yr): round_fn(stats.mean(period_vals.values()))
                        for yr, period_vals in by_year.items() if period_vals
                    }
                indicators[series_name] = {"label": REGION_SERIES_LABELS[series_name], "values": values}
            out = {
                "region":     label,
                "section":    section,
                "source":     "QGSO Regional Database (QRSIS)",
                "source_url": (
                    "https://www.qgso.qld.gov.au/statistics/queensland-regions/"
                    "regional-tools-statistics/queensland-regional-database"
                ),
                "note": (
                    "Median price and rent = mean of 4 quarterly medians. Sales count = "
                    "sum of the 4 quarterly values for the calendar year (each 'Year "
                    "Ended...' period is a per-quarter count despite its label). "
                    "Building approvals = sum of the monthly values present for the "
                    "year."
                ),
                "indicators": indicators,
            }
            slug = re.sub(r"[^a-z0-9]+", "_", label.lower()).strip("_")
            with open(out_dir / f"{section.lower()}_{slug}.json", "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2)
            self.log.info(f"  [{section}:{label}] {len(indicators)} series written")
            self.result.towns_ok.append(f"{section}:{label}")

    # ── Per-collection fetch ─────────────────────────────────────────────────────

    def _fetch_collection(self, coll_key, cfg, region_map) -> dict[str, dict]:
        """Run the QRSIS query wizard for one collection and return
        {town slug: {indicator: {year: value}}}.

        region_map: { region_code: [town, ...] }
          For SA2: { "307011176": [roma_town] }
          For LGA: { "LGA/34860": [roma_town, wallumbilla_town] }
        """
        session = requests.Session()
        session.headers.update(HEADERS)

        # Step 1: select collection
        udqctl_id = self._select_collection(session, cfg["id"])
        if not udqctl_id:
            raise RuntimeError(f"Failed to get udqctl_id for collection {cfg['id']}")
        self.log.info(f"    udqctl_id={udqctl_id}")

        # Step 2: series (list available, match exact names, select)
        avail_series  = self._get_options(session, "QIS1110W$UDQSER.ProcessSeries",
                                          udqctl_id, "infoser.htm", "p_new_multi")
        self.log.info(f"    Available series: {avail_series}")
        matched = self._match_series(cfg["series"], avail_series)
        if not matched:
            raise RuntimeError(f"No series matched for collection {cfg['id']}")
        time_periods_html = self._select_series(session, udqctl_id, matched)
        self.log.info(f"    Series selected: {matched}")

        # Step 3: time period
        self._set_time_period(session, udqctl_id, cfg["id"], cfg, time_periods_html)

        # Step 4: region type. SA2 and/or LGA is selected according to
        # the codes in region_map; LGA codes occur in the approvals map
        # (Toowoomba).
        has_lga = any(str(k).startswith("LGA/") for k in region_map)
        has_sa2 = any(not str(k).startswith("LGA/") for k in region_map)
        geo_label = "SA2 - Statistical Area Level 2"  # default
        if has_sa2:
            self._select_region_type(session, udqctl_id, "SA2 - Statistical Area Level 2")
        if has_lga:
            self._select_region_type(session, udqctl_id, "LGA - Local Government Area")

        # Step 5: regions
        avail_regions   = self._get_options(session, "QIS1110W$UDQREG.ProcessRegions",
                                            udqctl_id, "inforeg.htm", "p_new_multi")
        matched_regions = self._match_regions(region_map, avail_regions, cfg["geo"])
        if not matched_regions:
            raise RuntimeError(f"No {cfg['geo']} regions matched in QRSIS region list")
        self._select_regions(session, udqctl_id, matched_regions)
        self.log.info(f"    Regions selected: {len(matched_regions)}")

        # Step 6: submit
        html = self._submit_report(session, udqctl_id, cfg["id"])
        raw  = self._parse_output_html(html)
        self.log.info(f"    Parsed data for {len(raw)} regions")

        return self._aggregate(coll_key, matched, raw, region_map, cfg["geo"])

    # ── QRSIS wizard steps ───────────────────────────────────────────────────────

    def _select_collection(self, session, coll_id) -> Optional[str]:
        resp = session.post(
            BASE_URL + "QIS1110W$COLL.ProcessCollection",
            data=_q("usr_id", PUBLIC_USER, "access_lvl", ACCESS_LEVEL,
                    "coll_id", "", "collgrp_id", COLLGRP_ID,
                    "sel_coll_name", coll_id, "op_mode", "Next"),
            timeout=30, allow_redirects=True,
        )
        resp.raise_for_status()
        return self._extract_udqctl_id(resp.url, resp.text)

    def _get_options(self, session, proc, udqctl_id, info_page, select_name) -> list[str]:
        """GET a QRSIS page and return the options from the named select."""
        resp = session.get(
            BASE_URL + proc,
            params=_q("op_mode", "VIEW", "info_page", info_page,
                      "udqctl_id", udqctl_id, "error_msg", ""),
            timeout=30,
        )
        resp.raise_for_status()
        return _parse_options(resp.text, select_name)

    def _match_series(self, wanted: list[str], available: list[str]) -> list[str]:
        matched = []
        for want in wanted:
            if want in available:
                matched.append(want)
            else:
                self.log.error(f"    Series not found: '{want}'")
                self.log.error(f"    Available: {available}")
        return matched

    def _select_series(self, session, udqctl_id, series) -> str:
        for s in series:
            session.post(
                BASE_URL + "QIS1110W$UDQSER.ProcessActions",
                data=_q("udqctl_id", udqctl_id, "info_page", "infoser.htm",
                        "error_msg", "", "op_mode", "->") + [("p_new_multi", s)],
                timeout=30,
            )
            time.sleep(0.2)
        resp = session.post(
            BASE_URL + "QIS1110W$UDQSER.ProcessActions",
            data=_q("udqctl_id", udqctl_id, "info_page", "infoser.htm",
                    "error_msg", "", "op_mode", "Next"),
            timeout=30,
        )
        # Oracle PL/SQL WebTK returns the next screen (Time Periods) as the
        # response to this POST, so no separate request is needed. The page
        # is returned so that _set_time_period can read the available To
        # Date range from it (see _discover_max_to_date).
        return resp.text

    def _discover_max_to_date(self, time_periods_html: str) -> str | None:
        """Return the newest option of the 'To Date' dropdown on the Time
        Periods page, or None if it cannot be found.

        The options are free text ("Qtr Ended 31 Mar 2026", "Year Ended
        31 Dec 2025"), listed most recent first, with the newest marked
        selected="selected". Any non-empty option text is accepted (the
        equivalent function in fetch_population_erp.py accepts only plain
        digit years, because the ERP collection uses a bare year
        <select>). The selected option is preferred; if none is marked
        selected, the first option is returned. The date strings are not
        parsed or sorted, because the label styles ("Year Ended", "Qtr
        Ended") differ between collections.
        """
        soup = BeautifulSoup(time_periods_html, "lxml")
        to_date_marker = soup.find("input", {"name": "p_names", "value": "to_date"})
        if not to_date_marker:
            return None
        select = to_date_marker.find_next("select")
        if not select:
            return None
        options = [o.get_text(strip=True) for o in select.find_all("option") if o.get_text(strip=True)]
        if not options:
            return None
        selected = select.find("option", selected=True)
        if selected and selected.get_text(strip=True):
            return selected.get_text(strip=True)
        return options[0]

    def _set_time_period(self, session, udqctl_id, coll_id, cfg, time_periods_html: str = ""):
        from_date  = cfg["from_date"]
        # The To Date is read from the live page where possible, as in
        # fetch_population_erp.py, so the query extends to the newest period
        # without a config change. If that fails, cfg["to_date"] is used; it
        # still needs the manual "update each cycle" edit, and keeps the
        # fetcher working if the page structure changes.
        to_date = self._discover_max_to_date(time_periods_html) if time_periods_html else None
        if to_date:
            self.log.info(f"    Latest To Date option on the page: {to_date}")
        else:
            to_date = cfg["to_date"]
            self.log.warning(
                f"    Could not read the To Date option from the page -- "
                f"falling back to the hardcoded {to_date!r}, which may now be stale."
            )
        period     = cfg.get("period", "Quarterly")
        date_fmt   = cfg.get("date_fmt", "Y1")
        concorded  = cfg.get("concorded", "N")
        data = _q("udqctl_id", udqctl_id, "coll_id", coll_id, "error_msg", "",
                  "date_format", date_fmt, "period", period,
                  "from_date", from_date, "to_date", to_date)
        # p_concorded_data is sent only when the collection config asks for
        # "Y". All collections currently use "N", the pages' own hidden default.
        if concorded == "Y":
            data.append(("p_concorded_data", "Y"))
        data.append(("p_op_mode", "Next"))
        session.post(BASE_URL + "QIS1110W$UDQCTL.ProcessActions", data=data, timeout=30)

    def _select_region_type(self, session, udqctl_id, geo_label):
        session.post(
            BASE_URL + "QIS1110W$REGTYP.ProcessRegType",
            data=_q("udqctl_id", udqctl_id, "op_mode", "->")
                + [("p_multi", geo_label)],
            timeout=30,
        )
        time.sleep(0.2)
        session.post(
            BASE_URL + "QIS1110W$REGTYP.ProcessRegType",
            data=_q("udqctl_id", udqctl_id, "op_mode", "Next"),
            timeout=30,
        )

    def _match_regions(self, region_map, available, geo) -> list[str]:
        if available:
            self.log.info(f"    First available region: {available[0][:80]}")
        matched = []
        for code, towns in region_map.items():
            town_names = [t.name for t in towns] if isinstance(towns, list) else [towns.name]
            if str(code).startswith("LGA/"):
                # already has its prefix
                prefix = str(code)
            elif geo == "SA3":
                prefix = f"SA3/{code}"
            elif geo == "LGA":
                prefix = code  # already "LGA/NNNNN"
            else:
                prefix = f"SA2/{code}"
            hits = [r for r in available if r.startswith(prefix)]
            if hits:
                matched.extend(hits)
                self.log.info(f"    Matched {'/'.join(town_names)}: {hits[0][:60]}")
            else:
                self.log.warning(f"    [{'/'.join(town_names)}] {prefix} not in QRSIS list")
        return matched

    def _select_regions(self, session, udqctl_id, regions):
        session.post(
            BASE_URL + "QIS1110W$UDQREG.ProcessActions",
            data=_q("udqctl_id", udqctl_id, "info_page", "inforeg.htm",
                    "error_msg", "", "op_mode", "->")
                + [("p_new_multi", r) for r in regions],
            timeout=60,
        )
        time.sleep(0.5)
        session.post(
            BASE_URL + "QIS1110W$UDQREG.ProcessActions",
            data=_q("udqctl_id", udqctl_id, "info_page", "inforeg.htm",
                    "error_msg", "", "op_mode", "Next"),
            timeout=30,
        )

    def _submit_report(self, session, udqctl_id, coll_id) -> str:
        resp = session.post(
            BASE_URL + "QIS1110W$UDQCTL1.ProcessActions",
            data=_q("udqctl_id", udqctl_id, "coll_id", coll_id, "error_msg", "",
                    "ser_sort_col", "Sort Number", "reg_sort_col", "Region Code",
                    "display_style", "For each Region display Time Period by Series",
                    "op_mode", "QRSIS Query"),
            timeout=120,
        )
        resp.raise_for_status()
        return resp.text

    # ── Helpers ──────────────────────────────────────────────────────────────────

    def _extract_udqctl_id(self, url: str, html: str) -> Optional[str]:
        m = re.search(r'p_names=udqctl_id&p_values=(\d+)', url)
        if m:
            return m.group(1)
        m = re.search(r'[?&]udqctl_id=(\d+)', url)
        if m:
            return m.group(1)
        m = re.search(
            r'NAME="p_names"\s+VALUE="udqctl_id"[^>]*>.*?NAME="p_values"\s+VALUE="(\d+)"',
            html, re.IGNORECASE | re.DOTALL
        )
        if m:
            return m.group(1)
        m = re.search(r'udqctl_id.*?(\d{4,6})', url + html, re.DOTALL)
        if m:
            return m.group(1)
        return None

    # ── HTML parsing ─────────────────────────────────────────────────────────────

    def _parse_output_html(self, html: str) -> dict:
        """Parse QRSIS output HTML. Returns:
          { region_code: { period: { series_name: float|None } } }
        The region code keeps its type prefix as shown in the output,
        e.g. "SA2/307011176" or "LGA/34860".
        """
        result: dict[str, dict] = {}

        # The output has one section per region, headed for example
        # "Region : SA2/307011176 - Roma" or "Region : LGA/34860 - Maranoa (R)".
        # Splitting on "Region :" gives one section per region.
        sections = re.split(r'Region\s*:\s*', html)
        for section in sections[1:]:
            # Region code: SA2/..., SA3/..., SA4/..., LGA/... or S/...
            m = re.match(r'((?:SA2|SA3|SA4|LGA|S)/[\w]+)', section)
            if not m:
                continue
            region_code = m.group(1)
            region_data: dict[str, dict] = {}

            soup  = BeautifulSoup(section, "lxml")
            table = soup.find("table", border=True) or soup.find("table")
            if not table:
                result[region_code] = region_data
                continue

            rows = table.find_all("tr")
            if not rows:
                result[region_code] = region_data
                continue

            header_cells = rows[0].find_all(["th", "td"])
            series_names = [c.get_text(strip=True) for c in header_cells[1:]]

            for row in rows[1:]:
                cells = row.find_all(["td", "th"])
                if not cells:
                    continue
                period = cells[0].get_text(strip=True)
                if not period:
                    continue
                period_data: dict = {}
                for i, sname in enumerate(series_names, start=1):
                    if i < len(cells):
                        raw = cells[i].get_text(strip=True).replace(",", "").replace("$", "").strip()
                        try:
                            period_data[sname] = float(raw)
                        except ValueError:
                            period_data[sname] = None
                region_data[period] = period_data

            result[region_code] = region_data

        return result

    # ── Annual aggregation ───────────────────────────────────────────────────────

    def _aggregate(self, coll_key, matched_series, raw, region_map, geo) -> dict:
        from config import YEAR_START, YEAR_END
        import statistics as stats

        result: dict[str, dict] = {}

        # Resolve each region code to its towns.
        for code, towns in region_map.items():
            if str(code).startswith("LGA/"):
                region_key = str(code)
            elif geo == "SA3":
                region_key = f"SA3/{code}"
            elif geo == "LGA":
                region_key = code
            else:
                region_key = f"SA2/{code}"
            periods = raw.get(region_key) or raw.get(code) or {}
            if not periods:
                continue

            town_list = towns if isinstance(towns, list) else [towns]

            for town in town_list:
                slug = town.slug

                if coll_key == "sales":
                    count_yr: dict[str, int] = {}
                    for period, vals in periods.items():
                        if "31 Dec" not in period:
                            continue
                        yr = self._parse_year(period)
                        if yr and YEAR_START <= yr <= YEAR_END:
                            v = vals.get(SERIES_SALES_COUNT)
                            if v is not None:
                                count_yr[str(yr)] = int(v)

                    price_yr: dict[str, list] = {}
                    for period, vals in periods.items():
                        yr = self._parse_year(period)
                        if not yr or not (YEAR_START <= yr <= YEAR_END):
                            continue
                        v = vals.get(SERIES_SALES_PRICE)
                        if v is not None:
                            price_yr.setdefault(str(yr), []).append(v)
                    price_annual = {yr: round(stats.mean(vs)) for yr, vs in price_yr.items() if vs}

                    result[slug] = {
                        "housing_sales_count":  count_yr,
                        "housing_median_price": price_annual,
                    }

                elif coll_key == "rent":
                    rent_yr: dict[str, list] = {}
                    for period, vals in periods.items():
                        yr = self._parse_year(period)
                        if not yr or not (YEAR_START <= yr <= YEAR_END):
                            continue
                        v = vals.get(SERIES_RENT)
                        if v is not None:
                            rent_yr.setdefault(str(yr), []).append(v)
                    result[slug] = {"rent_3bed_median": {yr: round(stats.mean(vs)) for yr, vs in rent_yr.items() if vs}}

                elif coll_key in ("approvals_hist", "approvals_curr"):
                    app_yr: dict[str, int] = {}
                    for period, vals in periods.items():
                        parsed = self._parse_month_year(period)
                        if not parsed:
                            continue
                        yr, _ = parsed
                        if not (YEAR_START <= yr <= YEAR_END):
                            continue
                        v = vals.get(SERIES_APPROVALS)
                        if v is not None:
                            app_yr[str(yr)] = app_yr.get(str(yr), 0) + int(v)
                    # Merge the historical and current collections into one series.
                    if slug in result and "building_approvals" in result[slug]:
                        result[slug]["building_approvals"].update(app_yr)
                    else:
                        result.setdefault(slug, {})["building_approvals"] = app_yr

        return result

    @staticmethod
    def _parse_year(period: str) -> Optional[int]:
        m = re.search(r'\b(20\d{2}|19\d{2})\b', period)
        return int(m.group(1)) if m else None

    @staticmethod
    def _parse_month_year(period: str) -> Optional[tuple[int, int]]:
        months = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
                  "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
        m = re.match(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})', period.strip())
        return (int(m.group(2)), months[m.group(1)]) if m else None


if __name__ == "__main__":
    from logger import get_logger
    get_logger()
    result = QGSOHousingFetcher().run()
    sys.exit(0 if result.success else 1)