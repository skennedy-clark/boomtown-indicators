"""
tests/conftest.py

Shared fixtures for the test suite.

The fixtures build small, self-contained towns.toml files instead of
loading the project's own. Tests therefore do not depend on the current
town list and are unaffected when a town is added or a field corrected.
"""

import sys
from pathlib import Path

import pytest

# Adds regional-indicators/ to sys.path so that config.py, fetchers/ and
# the other packages import as they do for the fetchers themselves,
# whichever directory pytest is run from. This duplicates the pythonpath
# setting under [tool.pytest.ini_options] in pyproject.toml.
ROOT = Path(__file__).parent.parent / "regional-indicators"
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture
def sample_toml(tmp_path) -> Path:
    """A minimal, valid towns.toml: one study town and one benchmark town."""
    content = """
[settings]
output_dir = "output"

[towns.testtown]
name       = "Testtown"
state      = "QLD"
postcode   = "4000"
postcodes  = ["4000"]
sa2_code   = "300000000"
sa2_name   = "Testtown"
sa3_code   = "30000"
lga        = "Test LGA"
benchmark  = false

[towns.benchmarkcity]
name       = "Benchmark City"
state      = "QLD"
postcode   = "4001"
postcodes  = ["4001"]
benchmark  = true
"""
    path = tmp_path / "towns.toml"
    path.write_text(content)
    return path


@pytest.fixture
def broken_toml(tmp_path) -> Path:
    """A towns.toml with two validation problems: a duplicate town name,
    and a non-benchmark town with no sa2_code.
    """
    content = """
[towns.a]
name       = "Duplicate"
state      = "QLD"
postcode   = "4000"
postcodes  = ["4000"]
sa2_code   = "300000000"
benchmark  = false

[towns.b]
name       = "Duplicate"
state      = "QLD"
postcode   = "4001"
postcodes  = ["4001"]
benchmark  = false
"""
    path = tmp_path / "towns.toml"
    path.write_text(content)
    return path
