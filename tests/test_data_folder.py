"""
tests/test_data_folder.py

Tests for building the "1 Data" folder
(regional-indicators/transform/data_folder.py) from a small cache built
here. No network and no Excel are needed.
"""

import csv
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators" / "transform"))


@pytest.fixture
def cache(tmp_path):
    cache = tmp_path / "cache"
    (cache / "population").mkdir(parents=True)
    (cache / "fuel").mkdir()
    (cache / "qgso_ucl_erp.csv").write_text("ucl")
    (cache / "silo_rainfall_42078.csv").write_text("rain")
    (cache / "something_new.xlsx").write_text("?")
    (cache / "population" / "roma_population_ucl.json").write_text("{}")
    (cache / "fuel" / "racq-annual-fuel-price-report-2025.pdf").write_text("pdf")
    (cache / "fuel" / "racq_rulp_annual.json").write_text("{}")
    (cache / "index.json").write_text(json.dumps({
        "qgso_ucl_erp": {"path": str(cache / "qgso_ucl_erp.csv"), "url": "https://example.org/ucl.csv",
                         "downloaded_at": "2026-10-06T13:29:38"},
    }))
    return cache


@pytest.mark.parametrize("name, discipline", [
    ("qgso_ucl_erp.csv", "Population"),
    ("surat_basin_nrw_lga.xlsx", "Population"),
    ("abs_lga_population_2001_2025.xlsx", "Crime"),
    ("dcj_rent_2025q1.xlsx", "Housing"),
    ("ato_table8_2023-24.xlsx", "Income"),
    ("abs_nsw_labour_table002.xlsx", "Employment"),
    ("bocsar_lga.xlsm", "Crime"),
    ("School Profile 2008-2025.xlsx", "Education"),
    ("unknown.bin", "Other"),
])
def test_source_files_are_sorted_by_topic(name, discipline):
    from data_folder import classify

    assert classify(name) == discipline


def test_folder_holds_sources_extracted_figures_and_instructions(cache, tmp_path):
    from data_folder import build_data_folder

    out = tmp_path / "1 Data"
    lines = build_data_folder(cache, out)

    assert (out / "Population" / "qgso_ucl_erp.csv").exists()
    assert (out / "Population" / "extracted" / "roma_population_ucl.json").exists()
    assert (out / "Fuel" / "racq-annual-fuel-price-report-2025.pdf").exists()        # source file kept in a cache subfolder
    assert (out / "Fuel" / "extracted" / "racq_rulp_annual.json").exists()
    assert (out / "Rainfall" / "silo_rainfall_42078.csv").exists()
    assert (out / "Other" / "something_new.xlsx").exists()                          # nothing is lost
    assert not (out / "Other" / "index.json").exists()
    assert (out / "Employment" / "MANUAL DOWNLOAD - read me.txt").exists()
    assert (out / "Rainfall" / "MANUAL ENTRY - read me.txt").exists()
    assert "salm_smoothed_lga.csv" in (out / "README.txt").read_text()

    with open(out / "Population" / "Sources.csv", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    assert rows == [{"file": "qgso_ucl_erp.csv", "source_url": "https://example.org/ucl.csv",
                     "downloaded": "2026-10-06T13:29:38", "size_kb": "0"}]
    assert any("MANUAL DOWNLOAD NEEDED: Employment/salm_smoothed_lga.csv" in line for line in lines)


def test_manual_downloads_are_kept_and_imported_into_the_cache(cache, tmp_path):
    from data_folder import build_data_folder, import_manual_files

    out = tmp_path / "1 Data"
    (out / "Employment").mkdir(parents=True)
    (out / "Employment" / "salm_smoothed_lga.csv").write_text("lga")

    lines = import_manual_files(out, cache)
    assert (cache / "salm_smoothed_lga.csv").read_text() == "lga"
    assert lines[-1] == "Summary: 1 of 3 manual download(s) imported."

    build_data_folder(cache, out)                                       # rebuilding keeps the manual file
    assert (out / "Employment" / "salm_smoothed_lga.csv").read_text() == "lga"


def test_a_missing_manual_download_is_reported_not_fatal(cache, tmp_path):
    from data_folder import main

    assert main(["--out", str(tmp_path / "1 Data"), "--cache", str(cache), "--import-manual"]) == 0
