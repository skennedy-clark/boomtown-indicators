"""
regional-indicators/run_end_to_end.py

Runs the complete annual update as a single command.

Steps, in order:
  1. Copy the starting workbook (last year's delivered workbook) to the
     working copy, replacing any existing working copy. The starting
     workbook is never modified.
  2. Run every fetcher (run_update.py). Skipped with --skip-fetch, which
     reuses the existing cache.
  3. Run every writer in transform/xlsx_update/ against the working
     copy, one sheet at a time, each in its own Excel session.
  4. Build the website folder from the working copy
     (transform/web_export.py), compared with an earlier export if
     --previous-web is given.
  5. Print one line per step with its status and summary.

A step that fails is recorded and the run continues, so a single run
reports the state of every step. All output is also written to
regional-indicators/logs/end_to_end_<date>_<time>.log.

Each step is an ordinary command line, printed before it runs, and can
be re-run on its own.

The writers require Microsoft Excel (Windows or macOS). The fetch and
website steps run on any platform.

Usage (from the repository root):
    python regional-indicators/run_end_to_end.py <starting workbook.xlsx>
        [--test-copy test-copy.xlsx]
        [--previous-web <earlier export folder>]
        [--web-out "3 Web Content"]
        [--business-year 2025]      (default: the previous calendar year)
        [--skip-fetch] [--visible]
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

HERE      = Path(__file__).resolve().parent            # regional-indicators/
REPO_ROOT = HERE.parent
WRITERS   = HERE / "transform" / "xlsx_update"
CACHE     = HERE / "cache"
LOG_DIR   = HERE / "logs"
PAUSE_BETWEEN_EXCEL_STEPS_S = 3     # pause for Excel and file-sync clients to release the file

# (step name, writer script, cache folder), in sheet order.
WRITER_STEPS = [
    ("Population: town (UCL)",            "update_population_ucl.py",     "population"),
    ("Population: non-residents, town",   "update_population_nrw.py",     "population"),
    ("Population: non-residents, LGA",    "update_population_nrw_lga.py", "population"),
    ("Population: district (SA2)",        "update_population_erp.py",     "population"),
    ("Population: LGA",                   "update_population_erp_lga.py", "population"),
    ("Exogenous: rainfall",               "update_rainfall.py",           "rainfall"),
    ("Exogenous: education",              "update_education.py",          "schools"),
    ("Exogenous: fuel",                   "update_fuel.py",               "fuel"),
    ("Crime",                             "update_crime.py",              "crime"),
    ("Employment",                        "update_employment.py",         "unemployment"),
    ("Housing",                           "update_housing.py",            "housing"),
    ("Income",                            "update_income.py",             "ato"),
    ("Business",                          "update_business.py",           "business"),   # also takes a year argument
]


class Log:
    """Writes each line to the console and to the log file."""

    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.path = path
        self._file = open(path, "w", encoding="utf-8")

    def __call__(self, text: str = "") -> None:
        print(text, flush=True)
        self._file.write(text + "\n")
        self._file.flush()

    def close(self) -> None:
        self._file.close()


def run_step(name: str, command: list[str], log: Log) -> dict:
    """Run one command, echoing its output as it is produced.

    Never raises: a failure is reported in the returned result.
    """
    log("")
    log("=" * 78)
    log(f"STEP: {name}")
    log("  " + " ".join(f'"{c}"' if " " in c else c for c in command))
    log("=" * 78)
    started = time.time()
    summary = ""
    crashed = False
    try:
        process = subprocess.Popen(
            command, cwd=REPO_ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace",
        )
        for line in process.stdout:
            line = line.rstrip("\n")
            log(line)
            stripped = line.strip()
            if stripped.startswith(("Summary:", "Wrote", "Would write")):
                summary = stripped                      # summary line printed by a writer or by the export
            elif stripped.startswith(("✓ Passed :", "✗ Failed :")):
                # pass/fail totals printed by run_update.py
                summary = (summary + ", " if summary else "") + stripped[2:].replace(" :", ":").lower()
            if "Traceback (most recent call last)" in line:
                crashed = True
        code = process.wait()
        if crashed and code != 0:
            summary = "crashed (see the traceback above)"
    except OSError as exc:
        log(f"Could not start: {exc}")
        code, summary = -1, f"could not start: {exc}"
    return {
        "name": name, "ok": code == 0, "code": code,
        "summary": summary, "seconds": time.time() - started,
    }


def build_steps(args, test_copy: Path) -> list[tuple[str, list[str]]]:
    py = sys.executable
    steps: list[tuple[str, list[str]]] = []
    if not args.skip_fetch:
        steps.append(("Fetch everything", [py, str(HERE / "run_update.py")]))
    for name, script, cache in WRITER_STEPS:
        command = [py, str(WRITERS / script), str(test_copy), str(CACHE / cache)]
        if script == "update_business.py":
            command.append(str(args.business_year))
        if args.visible:
            command.append("--visible")
        steps.append((name, command))
    web = [py, str(HERE / "transform" / "web_export.py"), str(test_copy), "--out", str(args.web_out)]
    if args.previous_web:
        web += ["--compare-with", str(args.previous_web)]
    steps.append(("Website folder", web))
    return steps


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run the whole update end to end, for testing.")
    parser.add_argument("original", type=Path, help="last year's delivered workbook (the starting point)")
    parser.add_argument("--test-copy", type=Path, default=REPO_ROOT / "test-copy.xlsx",
                        help="working copy to create and write into (default: test-copy.xlsx)")
    parser.add_argument("--previous-web", type=Path, default=None,
                        help="an earlier '3 Web Content' folder to compare the new one with")
    parser.add_argument("--web-out", type=Path, default=REPO_ROOT / "3 Web Content",
                        help="where to build the website folder (default: '3 Web Content')")
    parser.add_argument("--business-year", type=int, default=datetime.now().year - 1,
                        help="year for the Business sheet (financial year ending in it); "
                             "default: last calendar year")
    parser.add_argument("--skip-fetch", action="store_true", help="reuse what is already in cache/")
    parser.add_argument("--visible", action="store_true", help="show Excel while the writers run")
    args = parser.parse_args(argv)

    original = args.original.resolve()
    test_copy = args.test_copy.resolve()
    if not original.exists():
        print(f"Starting workbook not found: {original}")
        return 2
    if original == test_copy:
        print("The starting workbook and the test copy are the same file -- refusing to overwrite it.")
        return 2

    log = Log(LOG_DIR / f"end_to_end_{datetime.now():%Y-%m-%d_%H%M%S}.log")
    log(f"End-to-end run started {datetime.now():%Y-%m-%d %H:%M:%S}")
    log(f"Starting workbook : {original}")
    log(f"Test copy         : {test_copy}")
    log(f"Website folder    : {args.web_out}")
    log(f"Python            : {sys.version.split()[0]} ({sys.executable})")

    shutil.copyfile(original, test_copy)
    log(f"Copied the starting workbook to {test_copy.name} (fresh copy).")

    results = []
    for name, command in build_steps(args, test_copy):
        result = run_step(name, command, log)
        results.append(result)
        if "xlsx_update" in " ".join(command):
            time.sleep(PAUSE_BETWEEN_EXCEL_STEPS_S)

    log("")
    log("=" * 78)
    log("END-TO-END SUMMARY")
    log("=" * 78)
    for r in results:
        status = "OK    " if r["ok"] else "FAILED"
        log(f"{status} {r['name']:<34} {r['seconds']:6.0f}s  {r['summary']}")
    failed = [r for r in results if not r["ok"]]
    log("")
    log(f"{len(results) - len(failed)} of {len(results)} steps finished without an error code.")
    log("('OK' means the step ran to the end -- read its summary for what was written or flagged.)")
    log(f"Log saved to {log.path}")
    log.close()
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
