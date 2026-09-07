"""
test_config.py -- Town/Config parsing and validation.

These exercise config.py directly against small fixture towns.toml files
(see conftest.py), never the real project towns.toml -- so adding or
correcting a real town can never break this suite.
"""

import pytest

from config import Config


def test_loads_towns_and_slug(sample_toml):
    config = Config(sample_toml)
    assert len(config.towns) == 2

    testtown = config.town_by_name("Testtown")
    assert testtown is not None
    assert testtown.slug == "testtown"
    assert testtown.is_qld
    assert not testtown.is_nsw
    assert not testtown.is_vic


def test_study_towns_excludes_benchmarks(sample_toml):
    config = Config(sample_toml)
    study = config.study_towns()
    assert [t.name for t in study] == ["Testtown"]


def test_towns_by_state(sample_toml):
    config = Config(sample_toml)
    assert len(config.towns_by_state("QLD")) == 2
    assert config.towns_by_state("NSW") == []


def test_town_by_name_missing_returns_none(sample_toml):
    config = Config(sample_toml)
    assert config.town_by_name("Nowhere") is None


def test_qld_study_towns(sample_toml):
    config = Config(sample_toml)
    assert [t.name for t in config.qld_study_towns()] == ["Testtown"]


def test_repr_reports_counts(sample_toml):
    config = Config(sample_toml)
    text = repr(config)
    assert "2 towns" in text
    assert "1 study" in text
    assert "1 benchmark" in text


def test_duplicate_name_raises(broken_toml):
    with pytest.raises(ValueError, match="Duplicate town name"):
        Config(broken_toml)


def test_missing_sa2_raises(broken_toml):
    # The same fixture's second town also has no sa2_code, so this error
    # is raised alongside the duplicate-name error above -- both land in
    # the same ValueError message.
    with pytest.raises(ValueError, match="missing sa2_code"):
        Config(broken_toml)