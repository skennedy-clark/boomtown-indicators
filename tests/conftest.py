"""
conftest.py -- shared fixtures for the boomtown-indicators test suite.

Deliberately builds small, self-contained towns.toml fixtures rather than
loading the real one: real fetchers/config tests should never depend on the
current state of the actual town list, or they'll break every time a town
is added or a field is corrected.
"""

import sys
from pathlib import Path

import pytest

# Belt-and-braces alongside pyproject.toml's [tool.pytest.ini_options]
# pythonpath setting, so `pytest` run from anywhere still finds config.py,
# fetchers/, etc. the same way the fetchers themselves import them.
ROOT = Path(__file__).parent.parent / "regional-indicators"
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture
def sample_toml(tmp_path) -> Path:
    """A minimal, valid towns.toml: one study town, one benchmark town."""
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
    """A towns.toml with two validation problems at once: a duplicate town
    name, and a non-benchmark town missing its sa2_code."""
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