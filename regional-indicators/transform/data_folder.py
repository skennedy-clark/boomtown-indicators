"""
regional-indicators/transform/data_folder.py

Builds the "1 Data" folder: every source file used for the year's
update, sorted into one subfolder per discipline, with a list of where
each file came from. It also holds the files that have to be downloaded
by hand, so that the person running the update deals with one folder.

Layout:
    1 Data/
        README.txt              what the folder is and what to do by hand
        Population/
            <downloaded source files>
            extracted/          figures taken from them, one JSON file per town or region
            Sources.csv         file, source URL, date downloaded, size
        Employment/  Housing/  Income/  Business/  Crime/
        Rainfall/  Fuel/  Education/
        Other/                  any cached file not matched to a discipline

Source files are copied from the pipeline's cache (regional-indicators/
cache/). Files already in the folder are replaced by newer copies; other
files there (for example the manual downloads) are left in place.

Manual downloads: a file the pipeline cannot download itself is listed
in MANUAL_FILES with the subfolder and name it is expected under. With
--import-manual, any such file found in the data folder is copied into
the cache before the fetchers run, where the fetcher finds it. Each
discipline with a manual file gets a "MANUAL DOWNLOAD - read me.txt"
with instructions.

Usage:
    python data_folder.py --out "1 Data"                  # build or refresh the folder
    python data_folder.py --out "1 Data" --import-manual  # copy manual downloads into the cache
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import shutil
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_CACHE = HERE.parent / "cache"

DISCIPLINES = ("Population", "Employment", "Housing", "Income", "Business",
               "Crime", "Rainfall", "Fuel", "Education")

# Subfolders of the cache that hold extracted figures (JSON), by discipline.
EXTRACTED_FOLDERS = {
    "population": "Population",
    "unemployment": "Employment",
    "housing": "Housing",
    "ato": "Income",
    "business": "Business",
    "crime": "Crime",
    "rainfall": "Rainfall",
    "fuel": "Fuel",
    "schools": "Education",
}

# Downloaded source files, matched on file name (case-insensitive), first match wins.
SOURCE_FILE_RULES = [
    (r"^qgso_ucl_erp|^surat_basin_|^bowen_basin_", "Population"),
    (r"^abs_lga_population", "Crime"),               # populations used to turn NSW offence counts into rates
    (r"^salm|^abs_nsw_labour", "Employment"),
    (r"^dcj_(rent|sales)|^vgv_|^dffh_", "Housing"),
    (r"^ato_", "Income"),
    (r"^abs_dc9|business", "Business"),
    (r"^qps_|^bocsar", "Crime"),
    (r"^silo_rainfall|^bom_|rainfall", "Rainfall"),
    (r"racq|fuel", "Fuel"),
    (r"school profile|acara", "Education"),
]

SKIP_FILES = {"index.json"}


@dataclass(frozen=True)
class ManualFile:
    discipline: str
    file_name: str          # name expected in the data folder and in the cache
    title: str
    instructions: str


MANUAL_FILES = [
    ManualFile(
        discipline="Employment",
        file_name="salm_smoothed_lga.csv",
        title="Small Area Labour Markets - smoothed LGA data (CSV)",
        instructions=(
            "1. Open https://www.dewr.gov.au/employment-research/small-area-labour-markets\n"
            "2. Download the latest 'SALM Smoothed LGA Datafiles' in CSV format.\n"
            "3. Save it in this folder as:  salm_smoothed_lga.csv\n"
            "The update tries to download this file itself; this is needed only if the run\n"
            "report says the LGA file was not found. Without it the LGA unemployment rows are\n"
            "not updated; everything else still runs."
        ),
    ),
    ManualFile(
        discipline="Housing",
        file_name="vgv_houses_by_suburb.xlsx",
        title="Valuer-General Victoria - Houses by suburb (time series, Excel)",
        instructions=(
            "Needed only for Victorian towns, and only if the run report says the file was not obtained.\n"
            "1. Open https://discover.data.vic.gov.au/dataset/victorian-property-sales-report-median-house-by-suburb-time-series\n"
            "2. Download the latest 'Houses by suburb <year>-<year>' file.\n"
            "3. Save it in this folder as:  vgv_houses_by_suburb.xlsx\n"
            "Without it the Victorian towns' Median Sale Price rows are not updated."
        ),
    ),
    ManualFile(
        discipline="Housing",
        file_name="dffh_rents_by_lga.xlsx",
        title="DFFH Rental Report - Quarterly median rents by LGA (Excel)",
        instructions=(
            "Needed only for Victorian towns, and only if the run report says the file was not obtained.\n"
            "1. Open https://discover.data.vic.gov.au/dataset/rental-report-quarterly-quarterly-median-rents-by-lga\n"
            "2. Open the latest quarter and download its Excel workbook.\n"
            "3. Save it in this folder as:  dffh_rents_by_lga.xlsx\n"
            "Without it the Victorian towns' House 3-Bed rent rows are not updated."
        ),
    ),
]

MANUAL_NOTES = {
    "Rainfall": (
        "Rainfall for stations that are not in the SILO dataset (currently Goondiwindi,\n"
        "station 041507, and Moranbah, station 034035) is entered by hand.\n\n"
        "1. Open the station's monthly rainfall page on BOM Climate Data Online; the run\n"
        "   log gives the exact link for each station.\n"
        "2. Enter each month's figure in regional-indicators/manual_rainfall_data.toml\n"
        "   (the file explains the format).\n"
        "Without these entries those two towns' rainfall rows are not updated."
    ),
}


def classify(file_name: str) -> str:
    name = file_name.lower()
    for pattern, discipline in SOURCE_FILE_RULES:
        if re.search(pattern, name):
            return discipline
    return "Other"


def _index_by_name(cache_dir: Path) -> dict[str, dict]:
    """Download records from the cache index, keyed by file name."""
    index_path = cache_dir / "index.json"
    if not index_path.exists():
        return {}
    try:
        index = json.loads(index_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    records = {}
    for record in index.values():
        if isinstance(record, dict) and record.get("path"):
            records[Path(record["path"]).name] = record
    return records


def _source_files(cache_dir: Path):
    """Yield (path, discipline) for every downloaded source file in the cache."""
    for path in sorted(cache_dir.iterdir()):
        if path.is_file() and path.name not in SKIP_FILES and not path.name.startswith("."):
            yield path, classify(path.name)
    for folder, discipline in EXTRACTED_FOLDERS.items():
        area = cache_dir / folder
        if area.is_dir():
            for path in sorted(area.iterdir()):
                if path.is_file() and path.suffix.lower() != ".json":
                    yield path, discipline


def _extracted_files(cache_dir: Path):
    """Yield (path, discipline, relative folder) for every extracted JSON file."""
    for folder, discipline in EXTRACTED_FOLDERS.items():
        area = cache_dir / folder
        if not area.is_dir():
            continue
        for path in sorted(area.rglob("*.json")):
            yield path, discipline, path.parent.relative_to(area)


def build_data_folder(cache_dir: Path, out_dir: Path) -> list[str]:
    """Copy the cache's source files and extracted figures into `out_dir`.
    Returns report lines."""
    records = _index_by_name(cache_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    sources: dict[str, list[dict]] = {}
    counts: dict[str, list[int]] = {}

    for path, discipline in _source_files(cache_dir):
        target = out_dir / discipline / path.name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, target)
        record = records.get(path.name, {})
        downloaded = record.get("downloaded_at") or datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds")
        sources.setdefault(discipline, []).append({
            "file": path.name,
            "source_url": record.get("url", ""),
            "downloaded": downloaded,
            "size_kb": round(path.stat().st_size / 1024),
        })
        counts.setdefault(discipline, [0, 0])[0] += 1

    for path, discipline, relative in _extracted_files(cache_dir):
        target = out_dir / discipline / "extracted" / relative / path.name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, target)
        counts.setdefault(discipline, [0, 0])[1] += 1

    for manual in MANUAL_FILES:
        folder = out_dir / manual.discipline
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "MANUAL DOWNLOAD - read me.txt").write_text(
            f"{manual.title}\n\n{manual.instructions}\n", encoding="utf-8")
        present = (folder / manual.file_name).exists()
        if present and not any(s["file"] == manual.file_name for s in sources.get(manual.discipline, [])):
            sources.setdefault(manual.discipline, []).append({
                "file": manual.file_name, "source_url": "downloaded by hand",
                "downloaded": datetime.fromtimestamp((folder / manual.file_name).stat().st_mtime).isoformat(timespec="seconds"),
                "size_kb": round((folder / manual.file_name).stat().st_size / 1024),
            })
    for discipline, note in MANUAL_NOTES.items():
        folder = out_dir / discipline
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "MANUAL ENTRY - read me.txt").write_text(note + "\n", encoding="utf-8")

    for discipline, rows in sources.items():
        with open(out_dir / discipline / "Sources.csv", "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=["file", "source_url", "downloaded", "size_kb"])
            writer.writeheader()
            writer.writerows(sorted(rows, key=lambda r: r["file"].lower()))

    _write_readme(out_dir)

    lines = [f"Data folder: {out_dir}"]
    for discipline in list(DISCIPLINES) + ["Other"]:
        if discipline in counts:
            raw, extracted = counts[discipline]
            lines.append(f"  {discipline:<11} {raw:3d} source files, {extracted:3d} extracted")
    missing = [m for m in MANUAL_FILES if not (out_dir / m.discipline / m.file_name).exists()]
    for manual in missing:
        lines.append(f"  MANUAL DOWNLOAD NEEDED: {manual.discipline}/{manual.file_name} ({manual.title})")
    raw_total = sum(c[0] for c in counts.values())
    lines.append(f"Summary: {raw_total} source files in {len(counts)} folders, "
                 f"{len(missing)} manual download(s) missing.")
    return lines


def _write_readme(out_dir: Path) -> None:
    manual = "\n".join(f"  - {m.discipline}/{m.file_name}: {m.title}" for m in MANUAL_FILES)
    entries = "\n".join(f"  - {d}: see '{d}/MANUAL ENTRY - read me.txt'" for d in MANUAL_NOTES)
    (out_dir / "README.txt").write_text(
        "1 Data\n"
        "======\n\n"
        "The source files used for this year's update of the indicators workbook, one folder\n"
        "per topic. Each folder's Sources.csv lists where every file came from and when it\n"
        "was downloaded. The 'extracted' subfolders hold the figures the pipeline took from\n"
        "those files, one file per town or region.\n\n"
        "Files to download by hand before the update is run (instructions in each folder):\n"
        f"{manual}\n\n"
        "Figures to enter by hand:\n"
        f"{entries}\n\n"
        f"Folder last refreshed {datetime.now():%Y-%m-%d %H:%M}.\n",
        encoding="utf-8",
    )


def import_manual_files(out_dir: Path, cache_dir: Path) -> list[str]:
    """Copy manual downloads found in the data folder into the cache."""
    lines = []
    cache_dir.mkdir(parents=True, exist_ok=True)
    found = 0
    for manual in MANUAL_FILES:
        source = out_dir / manual.discipline / manual.file_name
        if source.exists():
            shutil.copy2(source, cache_dir / manual.file_name)
            found += 1
            lines.append(f"Imported {manual.discipline}/{manual.file_name} into the cache.")
        else:
            lines.append(f"Not found (download it by hand to update these rows): {source}")
    lines.append(f"Summary: {found} of {len(MANUAL_FILES)} manual download(s) imported.")
    return lines


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build the '1 Data' folder of source files.")
    parser.add_argument("--out", type=Path, default=Path("1 Data"), help="the data folder (default: '1 Data')")
    parser.add_argument("--cache", type=Path, default=DEFAULT_CACHE, help="the pipeline's cache folder")
    parser.add_argument("--import-manual", action="store_true",
                        help="copy manual downloads from the data folder into the cache, then stop")
    args = parser.parse_args(argv)

    if args.import_manual:
        lines = import_manual_files(args.out, args.cache)
    else:
        if not args.cache.is_dir():
            print(f"Cache folder not found: {args.cache}")
            return 2
        lines = build_data_folder(args.cache, args.out)
    for line in lines:
        print(line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
