"""
run_update.py
-------------
Master orchestrator for the regional-indicators data pipeline.

Usage:
    python run_update.py                    # run all fetchers
    python run_update.py --only income      # run one fetcher by name
    python run_update.py --skip crime       # skip one fetcher
    python run_update.py --towns Roma Dalby # only process specific towns
    python run_update.py --force            # ignore cache, re-download everything
    python run_update.py --validate         # validate towns.toml only, no fetching
    python run_update.py --list-cache       # show what's in the cache index

Run from the regional-indicators/ directory:
    cd regional-indicators
    python run_update.py
"""

from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

# ── Local imports ──────────────────────────────────────────────────────────────
from config import get_config, get_cache, Config
from logger import get_logger
from fetchers.base import FetchResult

# ── Register all fetchers here ─────────────────────────────────────────────────
# Add new fetchers to this dict as they are built.
# key = short name used with --only / --skip flags

from fetchers.fetch_income import ATOIncomeFetcher
from fetchers.fetch_income_table6 import ATOTable6Fetcher
from fetchers.fetch_population_ucl import QGSOPopulationUCLFetcher
from fetchers.fetch_crime_qps import QPSCrimeFetcher

FETCHER_REGISTRY: dict[str, type] = {
    "income":       ATOIncomeFetcher,
    "income_table6":  ATOTable6Fetcher,
    "population_ucl": QGSOPopulationUCLFetcher,
    "crime_qps":      QPSCrimeFetcher,
    # "population":   ABSPopulationFetcher,       # TODO
    # "unemployment": ABSLabourFetcher,           # TODO
    # "housing":      QGSOHousingFetcher,         # TODO (QLD only, semi-manual)
    # "crime":        QPSCrimeFetcher,            # TODO (QLD only)
    # "crime_vic":    VicPolCrimeFetcher,         # TODO (VIC only)
    # "rainfall":     BOMRainfallFetcher,         # TODO
    # "fuel":         FuelPriceFetcher,           # TODO
}


# ── CLI ────────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Regional Indicators — data update pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument(
        "--only", nargs="+", metavar="FETCHER",
        help=f"Run only these fetchers. Choices: {', '.join(FETCHER_REGISTRY)}"
    )
    p.add_argument(
        "--skip", nargs="+", metavar="FETCHER",
        help="Skip these fetchers"
    )
    p.add_argument(
        "--towns", nargs="+", metavar="TOWN",
        help="Only process these towns (exact names from towns.toml)"
    )
    p.add_argument(
        "--force", action="store_true",
        help="Ignore cache and re-download all files"
    )
    p.add_argument(
        "--validate", action="store_true",
        help="Validate towns.toml and exit without fetching"
    )
    p.add_argument(
        "--list-cache", action="store_true",
        help="Show all cached files and exit"
    )
    p.add_argument(
        "--dry-run", action="store_true",
        help="Show what would run without actually fetching"
    )
    return p.parse_args()


# ── Helpers ────────────────────────────────────────────────────────────────────

def print_config_summary(config: Config, log):
    log.info("=" * 60)
    log.info("towns.toml summary")
    log.info("=" * 60)
    by_state: dict[str, list] = {}
    for t in config.towns:
        by_state.setdefault(t.state, []).append(t)

    for state, towns in sorted(by_state.items()):
        benchmarks = [t for t in towns if t.benchmark]
        study      = [t for t in towns if not t.benchmark]
        log.info(f"  {state}: {len(study)} study towns, {len(benchmarks)} benchmark")
        for t in study:
            issues = []
            if not t.sa2_code:
                issues.append("NO SA2")
            if not t.postcodes:
                issues.append("NO POSTCODES")
            flag = f"  ⚠ {', '.join(issues)}" if issues else ""
            log.info(f"    {t.name} ({t.postcode}){flag}")


def print_cache_summary(log):
    cache  = get_cache()
    entries = cache.list_entries()
    if not entries:
        log.info("Cache is empty.")
        return
    log.info(f"Cache entries ({len(entries)}):")
    for key, meta in sorted(entries.items()):
        path   = Path(meta["path"])
        exists = "✓" if path.exists() else "✗ MISSING"
        size   = f"{path.stat().st_size / 1024:.1f} KB" if path.exists() else ""
        log.info(f"  {exists}  {key}  ({meta.get('downloaded_at','?')})  {size}")


def build_run_summary(results: list[FetchResult], elapsed: float) -> str:
    lines = [
        "",
        "=" * 60,
        "RUN SUMMARY",
        "=" * 60,
        f"Total time : {elapsed:.1f}s",
        f"Fetchers   : {len(results)}",
        "",
    ]
    all_ok    = [r for r in results if r.success]
    all_fail  = [r for r in results if not r.success]

    lines.append(f"  ✓ Passed : {len(all_ok)}")
    for r in all_ok:
        lines.append(
            f"    {r.source:<20} "
            f"{len(r.towns_ok)} ok, "
            f"{len(r.towns_skipped)} cached, "
            f"{len(r.towns_failed)} failed"
            f"  ({r.elapsed_s:.1f}s)"
        )

    if all_fail:
        lines.append(f"  ✗ Failed : {len(all_fail)}")
        for r in all_fail:
            lines.append(f"    {r.source}")
            for e in r.errors:
                lines.append(f"      {e}")

    all_warnings = [w for r in results for w in r.warnings]
    if all_warnings:
        lines.append(f"\nWarnings ({len(all_warnings)}):")
        for w in all_warnings:
            lines.append(f"  {w}")

    lines.append("=" * 60)
    return "\n".join(lines)


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    ts   = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    log  = get_logger("regional-indicators", ts)
    args = parse_args()

    # ── Load & validate config ─────────────────────────────────────────────────
    log.info("Loading towns.toml …")
    try:
        config = get_config()
        log.info(f"Loaded: {config}")
    except (FileNotFoundError, ValueError) as exc:
        log.error(f"Config error: {exc}")
        sys.exit(1)

    print_config_summary(config, log)

    if args.validate:
        log.info("Validation passed — exiting (--validate flag set)")
        sys.exit(0)

    if args.list_cache:
        print_cache_summary(log)
        sys.exit(0)

    # ── Build list of fetchers to run ──────────────────────────────────────────
    to_run: dict[str, type] = dict(FETCHER_REGISTRY)

    if args.only:
        unknown = [k for k in args.only if k not in FETCHER_REGISTRY]
        if unknown:
            log.error(f"Unknown fetcher(s): {unknown}. "
                      f"Available: {list(FETCHER_REGISTRY)}")
            sys.exit(1)
        to_run = {k: v for k, v in to_run.items() if k in args.only}

    if args.skip:
        to_run = {k: v for k, v in to_run.items() if k not in args.skip}

    if not to_run:
        log.warning("No fetchers selected — nothing to do.")
        sys.exit(0)

    log.info(f"Fetchers to run: {list(to_run)}")

    # ── Override cache for --force ─────────────────────────────────────────────
    if args.force:
        log.warning("--force flag set: invalidating all cache entries")
        cache = get_cache()
        for key in list(cache.list_entries()):
            cache.invalidate(key)

    # ── Dry run ────────────────────────────────────────────────────────────────
    if args.dry_run:
        log.info("DRY RUN — no fetching will occur")
        for name in to_run:
            log.info(f"  Would run: {name}")
        sys.exit(0)

    # ── Run fetchers ───────────────────────────────────────────────────────────
    results:    list[FetchResult] = []
    run_start = time.time()

    for name, FetcherClass in to_run.items():
        log.info(f"\n{'─' * 40}")
        log.info(f"Running: {name}")
        log.info(f"{'─' * 40}")

        try:
            fetcher = FetcherClass()

            # Apply --towns filter if specified
            if args.towns:
                # Patch config to only include specified towns
                # (simple approach: filter in applicable_towns)
                original_study = fetcher.config.study_towns
                specified = args.towns
                fetcher.config.study_towns = lambda: [
                    t for t in original_study()
                    if t.name in specified
                ]

            result = fetcher.run()
            results.append(result)

        except Exception as exc:
            log.error(f"Failed to instantiate or run {name}: {exc}", exc_info=True)
            results.append(FetchResult(
                source  = name,
                success = False,
                errors  = [str(exc)],
            ))

    # ── Summary ────────────────────────────────────────────────────────────────
    elapsed = time.time() - run_start
    summary = build_run_summary(results, elapsed)
    log.info(summary)

    # Write summary to a separate file for easy access
    summary_path = Path(__file__).parent / "logs" / f"{ts}_summary.txt"
    summary_path.write_text(summary, encoding="utf-8")
    log.info(f"Summary written to: {summary_path}")

    # Exit code reflects whether any fetcher failed
    failed = any(not r.success for r in results)
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()