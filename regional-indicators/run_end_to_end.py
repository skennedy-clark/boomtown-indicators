"""
regional-indicators/run_end_to_end.py

Runs the complete annual update as a single command.

Steps, in order:
  1. Copy the starting workbook (last year's delivered workbook) to the
     working copy, replacing any existing working copy. The starting
     workbook is never modified.
  2. Copy any manual downloads from the data folder ("1 Data") into the
     cache, where the fetchers look for them (transform/data_folder.py).
  3. Run every fetcher (run_update.py). Skipped with --skip-fetch, which
     reuses the existing cache.
  4. Run every writer in transform/xlsx_update/ against the working
     copy, one sheet at a time, each in its own Excel session.
  5. Bring the charts up to date: show the new year in every chart
     whose data has it (transform/xlsx_update/update_charts.py).
  6. Build the website folder from the working copy
     (transform/web_export.py), compared with an earlier export if
     --previous-web is given.
  7. Build the data folder: every source file used, by topic, with
     where it came from (transform/data_folder.py).
  8. Print one line per step with its status and summary, and write a
     plain-English run report listing everything flagged for checking.

The year to add is the year after the starting workbook's latest year,
but never a year that has not finished; --year overrides it. It is
passed to every step through the BOOMTOWN_YEAR_END environment variable
(see config.py). With --output-folder, the year's results are gathered
in one folder: "1 Data", "2 Workbook" (the finished workbook, named for
the year), "3 Web Content" and "Run report.txt". "Run annual update.bat"
in the repository root runs this with sensible settings.

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
        [--output-folder <folder> | --output-root <folder>] [--open-report]
        [--web-out "3 Web Content"] [--data-out "1 Data"] [--year <year>]
        [--business-year 2025]      (default: the previous calendar year)
        [--last-year 2025]          (default: the previous calendar year)
        [--skip-fetch] [--visible]
"""

from __future__ import annotations

import argparse
import re
import os
import shutil
import subprocess
import threading
import sys
import time
from datetime import datetime
from pathlib import Path

HERE      = Path(__file__).resolve().parent            # regional-indicators/
REPO_ROOT = HERE.parent
WRITERS   = HERE / "transform" / "xlsx_update"
CACHE     = HERE / "cache"
LOG_DIR   = HERE / "logs"
RETRY_AFTER_EXCEL_ERROR_S = 15
PAUSE_BETWEEN_EXCEL_STEPS_S = 3     # pause for Excel and file-sync clients to release the file
SETTLE_AFTER_COPY_S = 10            # let a file-sync client finish with the new copy before Excel opens it
STEP_TIMEOUT_MIN = {"fetch": 30, "excel": 15, "other": 10}

# Output lines worth repeating in the run report: values not written,
# things to do by hand, and errors.
NOTE_MARKERS = ("FLAGGED", "SKIPPED", "MANUAL DOWNLOAD NEEDED", "NOT FOUND", "FAILED",
                "no category for", "STILL OUT OF DATE", "Error", "not found")
MAX_NOTES_PER_STEP = 200

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
        try:
            print(text, flush=True)
        except UnicodeEncodeError:      # console or redirect that is not UTF-8
            encoding = sys.stdout.encoding or "ascii"
            print(text.encode(encoding, "replace").decode(encoding), flush=True)
        self._file.write(text + "\n")
        self._file.flush()

    def close(self) -> None:
        self._file.close()


def run_step(name: str, command: list[str], log: Log, timeout_s: float | None = None) -> dict:
    """Run one command, echoing its output as it is produced.

    A step still running after `timeout_s` seconds is stopped; an Excel
    step usually hangs only when Excel is waiting on a dialog that no
    one can see. Never raises: a failure is reported in the returned
    result.
    """
    log("")
    log("=" * 78)
    log(f"STEP: {name}")
    log("  " + " ".join(f'"{c}"' if " " in c else c for c in command))
    log("=" * 78)
    started = time.time()
    summary = ""
    crashed = False
    excel_error = False
    timed_out = threading.Event()
    notes: list[str] = []
    # Child processes write to a pipe, where Python on Windows defaults to
    # the ANSI code page and fails on characters outside it. Force UTF-8 to
    # match the decoding below.
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    try:
        process = subprocess.Popen(
            command, cwd=REPO_ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", env=env,
        )
        timer = None
        if timeout_s:
            def stop():
                timed_out.set()
                process.kill()
            timer = threading.Timer(timeout_s, stop)
            timer.daemon = True
            timer.start()
        for line in process.stdout:
            line = line.rstrip("\n")
            log(line)
            stripped = line.strip()
            if stripped.startswith(("Summary:", "Wrote", "Would write")):
                summary = stripped                      # summary line printed by a writer or by the export
            elif stripped.startswith(("✓ Passed :", "✗ Failed :")):
                # pass/fail totals printed by run_update.py
                summary = (summary + ", " if summary else "") + stripped[2:].replace(" :", ":").lower()
            if any(marker in stripped for marker in NOTE_MARKERS) and len(notes) < MAX_NOTES_PER_STEP:
                notes.append(stripped)
            if "Traceback (most recent call last)" in line:
                crashed = True
            if "pywintypes.com_error" in line:
                excel_error = True
        code = process.wait()
        if timer is not None:
            timer.cancel()
        if crashed and code != 0:
            summary = "crashed (see the traceback above)"
        if timed_out.is_set():
            minutes = timeout_s / 60
            log(f"Stopped after {minutes:.0f} minutes without finishing.")
            code = code if code else -1
            summary = f"stopped after {minutes:.0f} minutes (see the note at the end of the log)"
    except OSError as exc:
        log(f"Could not start: {exc}")
        code, summary = -1, f"could not start: {exc}"
    return {
        "name": name, "ok": code == 0, "code": code,
        "summary": summary, "seconds": time.time() - started,
        "excel_error": excel_error and code != 0,
        "timed_out": timed_out.is_set(),
        "notes": notes,
    }


YEAR_HEADER_ROWS = (("Crime", 1), ("Exogenous", 1), ("Housing", 1), ("Employment", 2))


def latest_year_in_workbook(path: Path) -> int | None:
    """Return the latest year in the year headers of the starting
    workbook's data sheets (Crime, Exogenous, Housing, Employment), or
    None if it cannot be read. Projection years further ahead than the
    current year are ignored."""
    try:
        import openpyxl
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return None
    this_year = datetime.now().year
    found = []
    try:
        for sheet, row in YEAR_HEADER_ROWS:
            if sheet not in book.sheetnames:
                continue
            cells = next(book[sheet].iter_rows(min_row=row, max_row=row, values_only=True), ())
            years = [int(v) for v in cells
                     if isinstance(v, (int, float)) and not isinstance(v, bool)
                     and float(v).is_integer() and 1990 <= v <= this_year]
            if years:
                found.append(max(years))
    finally:
        book.close()
    return max(found) if found else None


def choose_data_year(latest_in_workbook: int | None, requested: int | None = None) -> tuple[int, str]:
    """The year to add: the one after the starting workbook's latest year,
    but never a year that has not finished. Returns (year, explanation)."""
    last_complete = datetime.now().year - 1
    if requested:
        return requested, "as requested (--year)"
    if latest_in_workbook is None:
        return last_complete, "the last complete calendar year (the workbook's years could not be read)"
    year = min(latest_in_workbook + 1, last_complete)
    if year == latest_in_workbook:
        return year, (f"the starting workbook already runs to {latest_in_workbook}, and {latest_in_workbook + 1} "
                      f"has not finished: the run refreshes {year}")
    return year, f"the year after the starting workbook's latest year, {latest_in_workbook}"


def in_synced_folder(path: Path) -> bool:
    """True if `path` is inside a OneDrive folder (or one whose name
    says it is synchronised by OneDrive)."""
    roots = [os.environ.get(name) for name in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer")]
    resolved = str(path.resolve()).lower()
    if any(root and resolved.startswith(str(Path(root).resolve()).lower()) for root in roots):
        return True
    return any(part.lower().startswith("onedrive") for part in path.resolve().parts)


def copy_starting_workbook(original: Path, test_copy: Path) -> str | None:
    """Copy the starting workbook over the working copy.

    Returns None on success, or a message explaining why the working
    copy could not be created. Nothing else runs in that case: every
    later step depends on being able to write this file.
    """
    lock_file = test_copy.with_name("~$" + test_copy.name)     # created by Excel while a workbook is open
    try:
        shutil.copyfile(original, test_copy)
    except PermissionError:
        reason = (
            "It is open in Excel; close it and run again."
            if lock_file.exists() else
            "It is probably open in Excel, or held by a file-sync client "
            "(OneDrive); close it, or pause syncing, and run again."
        )
        return f"Cannot write the working copy {test_copy}.\n{reason}"
    except OSError as exc:
        return f"Cannot create the working copy {test_copy}: {exc}"
    return None


PLAIN_NAMES = {
    "Manual downloads": "Manual downloads copied in from the data folder",
    "Fetch everything": "Downloading the latest figures",
    "Charts": "Charts brought up to date",
    "Website folder": "Website files (3 Web Content)",
    "Data folder": "Source files gathered (1 Data)",
}


ANSI = re.compile(r"\x1b?\[[0-9;]{1,5}m")          # colour codes, with or without the escape character
LOG_PREFIX = re.compile(r"^\d{2}:\d{2}:\d{2}\s+(INFO|WARNING|ERROR|DEBUG)\s+")
MISSING_BLOCK = re.compile(r"^(.+?)(?: \(Table 6\))?: SKIPPED \(row-finding\) .*Could not find town '([^']+)'")
NOTE_LENGTH = 240


def report_notes(notes: list[str]) -> list[str]:
    """Turn a step's flagged output lines into short, readable report lines.

    Colour codes and log prefixes are removed, a line's technical detail
    (after the first ' | ') is left to the log, routine status lines are
    dropped, and towns that have no block in a sheet are listed together.
    """
    lines: list[str] = []
    no_block: list[str] = []
    for raw in notes:
        cleaned = ANSI.sub("", raw).strip()
        if re.match(r"^\d{2}:\d{2}:\d{2}\s+WARNING", cleaned):
            continue
        note = LOG_PREFIX.sub("", cleaned).strip()
        if not note or note.startswith(("Summary:", "Status :", "Errors :")):
            continue
        missing = MISSING_BLOCK.match(note)
        if missing:
            if missing.group(2) not in no_block:
                no_block.append(missing.group(2))
            continue
        note = note.split(" | ")[0]
        if note.startswith("[LGA] "):
            note = note[len("[LGA] "):]
        if len(note) > NOTE_LENGTH:
            note = note[:NOTE_LENGTH].rstrip() + "... (full text in the log)"
        if note not in lines:
            lines.append(note)
    if no_block:
        lines.append("Not in this sheet, so nothing written (expected for towns the sheet does not cover): "
                     + ", ".join(no_block))
    return lines


def write_report(path: Path, results: list[dict], context: dict) -> None:
    """Write a plain-English summary of the run for the person who owns
    the workbook: what was produced, what needs checking, and where."""
    failed = [r for r in results if not r["ok"] and r["name"] != "Fetch everything"]
    fetch = next((r for r in results if r["name"] == "Fetch everything"), None)
    lines = [
        f"Indicators annual update - run report",
        f"Run on {datetime.now():%d %B %Y at %H:%M}",
        "",
        f"Starting workbook : {context['original']}",
        f"Year added        : {context['year']} ({context['year_reason']})",
        "",
        "RESULT",
        "------",
        ("Every step finished." if not failed else
         f"{len(failed)} step(s) did not finish: " + ", ".join(r["name"] for r in failed) + ". "
         "See the log for details; the other steps' results are still usable."),
    ]
    if fetch is not None and not fetch["ok"]:
        lines.append("Some data sources could not be downloaded (see 'Downloading the latest figures' "
                     "under PLEASE CHECK); the figures that come from them were not updated.")
    lines += [
        "",
    ]
    if context.get("outputs"):
        lines += ["WHERE THE RESULTS ARE", "---------------------"]
        lines += [f"  {label:<14} {where}" for label, where in context["outputs"]]
        lines.append("")
    lines += ["WHAT EACH STEP DID", "------------------"]
    for r in results:
        name = PLAIN_NAMES.get(r["name"], f"Workbook: {r['name']}")
        if r["ok"]:
            status = "done"
        elif r["name"] == "Fetch everything":
            status = "some sources unavailable"
        else:
            status = "DID NOT FINISH"
        lines.append(f"  [{status}] {name}: {r['summary'] or '-'}")
    lines.append("")

    to_check = [(PLAIN_NAMES.get(r["name"], r["name"]), note)
                for r in results for note in report_notes(r.get("notes", []))]
    lines += ["PLEASE CHECK", "------------"]
    if not to_check:
        lines.append("  Nothing was flagged.")
    else:
        lines += [
            "  These values were not written, or were written and marked for review (bold red in",
            "  the workbook), because they did not pass the automatic checks, or something has to",
            "  be done by hand. Check each against its source. A value that is right can be",
            "  typed into the workbook, or recorded in verified_overrides.toml so it is written",
            "  next time.",
            "",
        ]
        current = None
        for step, note in to_check:
            if step != current:
                lines.append(f"  {step}")
                current = step
            lines.append(f"    - {note}")
    lines += [
        "",
        "BEFORE THE NEXT RUN",
        "-------------------",
        "  Close Excel. If a step did not finish, end any EXCEL.EXE processes in Task Manager.",
        "  Manual downloads go in the data folder; each topic folder has a 'read me' saying",
        "  what to download and from where. Runs can be repeated: each starts again from the",
        "  starting workbook.",
        "",
        f"Full technical log: {context['log']}",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def build_steps(args, test_copy: Path) -> list[tuple[str, list[str]]]:
    py = sys.executable
    steps: list[tuple[str, list[str]]] = []
    data_folder = [py, str(HERE / "transform" / "data_folder.py"), "--out", str(args.data_out)]
    steps.append(("Manual downloads", data_folder + ["--import-manual"]))
    if not args.skip_fetch:
        steps.append(("Fetch everything", [py, str(HERE / "run_update.py")]))
    for name, script, cache in WRITER_STEPS:
        command = [py, str(WRITERS / script), str(test_copy), str(CACHE / cache)]
        if script == "update_business.py":
            command.append(str(args.business_year))
        if args.visible:
            command.append("--visible")
        steps.append((name, command))
    charts = [py, str(WRITERS / "update_charts.py"), str(test_copy), "--last-year", str(args.last_year)]
    if args.visible:
        charts.append("--visible")
    steps.append(("Charts", charts))
    web = [py, str(HERE / "transform" / "web_export.py"), str(test_copy), "--out", str(args.web_out)]
    if args.previous_web:
        web += ["--compare-with", str(args.previous_web)]
    steps.append(("Website folder", web))
    steps.append(("Data folder", data_folder))
    return steps


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run the whole update end to end, for testing.")
    parser.add_argument("original", type=Path, help="last year's delivered workbook (the starting point)")
    parser.add_argument("--test-copy", type=Path, default=REPO_ROOT / "test-copy.xlsx",
                        help="working copy to create and write into (default: test-copy.xlsx)")
    parser.add_argument("--previous-web", type=Path, default=None,
                        help="an earlier '3 Web Content' folder to compare the new one with")
    parser.add_argument("--output-folder", type=Path, default=None,
                        help="put everything for the year in this folder: '1 Data', '2 Workbook' "
                             "(the finished workbook), '3 Web Content' and 'Run report.txt'")
    parser.add_argument("--output-root", type=Path, default=None,
                        help="create the output folder here, named 'Indicators update <year>' "
                             "after the workbook it produces")
    parser.add_argument("--open-report", action="store_true",
                        help="open the run report when the run has finished (Windows)")
    parser.add_argument("--web-out", type=Path, default=None,
                        help="where to build the website folder (default: '3 Web Content' in the "
                             "output folder, or in the repository)")
    parser.add_argument("--data-out", type=Path, default=None,
                        help="the data folder: source files, and manual downloads to use "
                             "(default: '1 Data' in the output folder, or in the repository)")
    parser.add_argument("--year", type=int, default=None,
                        help="the data year to add (default: the year after the starting "
                             "workbook's latest year, but not a year that has not finished)")
    parser.add_argument("--business-year", type=int, default=None,
                        help="year for the Business sheet (financial year ending in it); "
                             "default: the data year")
    parser.add_argument("--last-year", type=int, default=None,
                        help="latest data year, used when updating the charts: data after it "
                             "is treated as projections (default: the data year)")
    parser.add_argument("--skip-fetch", action="store_true", help="reuse what is already in cache/")
    parser.add_argument("--visible", action="store_true", help="show Excel while the writers run")
    args = parser.parse_args(argv)

    year, year_reason = choose_data_year(latest_year_in_workbook(args.original), args.year)
    if args.output_root and not args.output_folder:
        args.output_folder = args.output_root / f"Indicators update {year + 1}"
    base = args.output_folder.resolve() if args.output_folder else REPO_ROOT
    args.web_out = args.web_out or base / "3 Web Content"
    args.data_out = args.data_out or base / "1 Data"
    args.business_year = args.business_year or year
    args.last_year = args.last_year or year
    os.environ["BOOMTOWN_YEAR_END"] = str(year)             # read by config.py in every step

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
    log(f"Data folder       : {args.data_out}")
    log(f"Year added        : {year} ({year_reason})")
    log(f"Python            : {sys.version.split()[0]} ({sys.executable})")

    problem = copy_starting_workbook(original, test_copy)
    if problem:
        log("")
        log(problem)
        log("Nothing was run.")
        log.close()
        return 2
    log(f"Copied the starting workbook to {test_copy.name} (fresh copy).")
    if in_synced_folder(test_copy):
        log("")
        log("NOTE: the working copy is in a OneDrive folder. OneDrive uploads the new copy while")
        log("Excel is opening it, which can make Excel lose the workbook part-way through a step")
        log("('The object invoked has disconnected from its clients'). If that happens, put the")
        log("working copy in a local folder, for example:  --test-copy C:\\Temp\\test-copy.xlsx")
    time.sleep(SETTLE_AFTER_COPY_S)

    results = []
    for name, command in build_steps(args, test_copy):
        is_writer = "xlsx_update" in " ".join(command)
        kind = "excel" if is_writer else ("fetch" if name == "Fetch everything" else "other")
        timeout_s = STEP_TIMEOUT_MIN[kind] * 60
        result = run_step(name, command, log, timeout_s)
        if is_writer and result["excel_error"]:
            # Excel automation errors are usually transient (the workbook
            # is still held by a file-sync client or a closing Excel
            # process). A writer saves only after all of its writes, so a
            # failed attempt leaves the workbook unchanged.
            log(f"Excel reported an automation error; retrying '{name}' once "
                f"in {RETRY_AFTER_EXCEL_ERROR_S} seconds.")
            time.sleep(RETRY_AFTER_EXCEL_ERROR_S)
            first_seconds = result["seconds"]
            result = run_step(name + " (retry)", command, log, timeout_s)
            result["name"] = name
            result["seconds"] += first_seconds
            if result["ok"]:
                result["summary"] += "  [succeeded on retry]"
        results.append(result)
        if is_writer:
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
    if any(r.get("timed_out") or r.get("excel_error") for r in results):
        log("")
        log("Excel did not finish one or more steps. Before running again, close Excel and end any")
        log("EXCEL.EXE processes left in Task Manager (Details tab): a hidden Excel from a failed")
        log("step keeps the working copy open, and the next step then waits on it.")
    outputs = [("Data", args.data_out), ("Website", args.web_out)]
    if args.output_folder:
        finished = base / "2 Workbook" / f"Indicators Data-Charts {year + 1}.xlsx"
        try:
            finished.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(test_copy, finished)
            outputs.insert(1, ("Workbook", finished))
            log(f"Finished workbook copied to {finished}")
        except OSError as exc:
            outputs.insert(1, ("Workbook", test_copy))
            log(f"Could not copy the finished workbook to {finished} ({exc}); it is at {test_copy}")
        report = base / "Run report.txt"
    else:
        outputs.insert(1, ("Workbook", test_copy))
        report = log.path.with_name(log.path.stem.replace("end_to_end", "run_report") + ".txt")
    write_report(report, results, {
        "original": original, "year": year, "year_reason": year_reason,
        "outputs": outputs, "log": log.path,
    })
    log(f"Run report saved to {report}")
    if args.open_report and hasattr(os, "startfile"):
        try:
            os.startfile(report)                              # Windows: opens in Notepad
        except OSError:
            pass
    log(f"Log saved to {log.path}")
    log.close()
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
