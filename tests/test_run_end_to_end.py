"""
tests/test_run_end_to_end.py -- the end-to-end runner
(regional-indicators/run_end_to_end.py).

Tests the runner itself, with tiny stand-in scripts: that it copies the
starting workbook, runs the steps in order with the right arguments,
keeps going when a step fails, and reports each step's own summary.
No fetching, no Excel.
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "regional-indicators"))


@pytest.fixture
def log(tmp_path):
    import run_end_to_end as e2e

    log = e2e.Log(tmp_path / "run.log")
    yield log
    log.close()


def _script(tmp_path, name, body):
    path = tmp_path / name
    path.write_text(body, encoding="utf-8")
    return str(path)


def test_step_reports_the_writers_own_summary_line(tmp_path, log):
    import run_end_to_end as e2e

    script = _script(tmp_path, "ok.py", 'print("Dalby: WRITTEN 2025 = 1 -> $Z$9")\nprint("Summary: 8 written, 0 flagged for review.")\n')
    result = e2e.run_step("Fuel", [sys.executable, script], log)
    assert result["ok"] and result["summary"] == "Summary: 8 written, 0 flagged for review."


def test_failed_step_is_recorded_not_raised(tmp_path, log):
    import run_end_to_end as e2e

    script = _script(tmp_path, "boom.py", 'raise RuntimeError("no Excel here")\n')
    result = e2e.run_step("Crime", [sys.executable, script], log)
    assert not result["ok"] and "crashed" in result["summary"]
    log.close()
    assert "no Excel here" in (tmp_path / "run.log").read_text(encoding="utf-8")


def test_fetch_totals_are_combined(tmp_path, log):
    import run_end_to_end as e2e

    script = _script(tmp_path, "fetch.py", 'import sys\nprint("  ✓ Passed : 16")\nprint("  ✗ Failed : 2")\nsys.exit(1)\n')
    result = e2e.run_step("Fetch everything", [sys.executable, script], log)
    assert result["summary"] == "passed: 16, failed: 2"


def test_steps_are_fetch_then_every_writer_then_the_website(tmp_path):
    import argparse
    import run_end_to_end as e2e

    args = argparse.Namespace(skip_fetch=False, business_year=2025, visible=True,
                              web_out=tmp_path / "3 Web Content", previous_web=tmp_path / "old")
    steps = e2e.build_steps(args, tmp_path / "test-copy.xlsx")
    names = [name for name, _ in steps]
    assert names[0] == "Fetch everything" and names[-1] == "Website folder"
    assert len(names) == len(e2e.WRITER_STEPS) + 2

    commands = dict(steps)
    business = commands["Business"]
    assert business[-2:] == ["2025", "--visible"]                      # the one writer that takes a year
    assert Path(commands["Income"][3]).name == "ato"                   # Income reads cache/ato, not cache/income
    assert Path(commands["Population: LGA"][3]).name == "population"
    assert "--compare-with" in commands["Website folder"]

    # every writer named here exists
    for _, script, _ in e2e.WRITER_STEPS:
        assert (e2e.WRITERS / script).exists(), script


def test_skip_fetch_leaves_the_fetch_step_out(tmp_path):
    import argparse
    import run_end_to_end as e2e

    args = argparse.Namespace(skip_fetch=True, business_year=2025, visible=False,
                              web_out=tmp_path / "w", previous_web=None)
    steps = e2e.build_steps(args, tmp_path / "t.xlsx")
    assert steps[0][0] != "Fetch everything"
    assert "--compare-with" not in steps[-1][1]


def test_refuses_to_overwrite_the_starting_workbook(tmp_path, capsys):
    import run_end_to_end as e2e

    book = tmp_path / "original.xlsx"
    book.write_bytes(b"PK")
    assert e2e.main([str(book), "--test-copy", str(book)]) == 2
    assert book.read_bytes() == b"PK"
