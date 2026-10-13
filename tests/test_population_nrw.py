"""
tests/test_population_nrw.py

Tests for the LGA name matching in the non-resident worker fetcher
(regional-indicators/fetchers/fetch_population_nrw.py). towns.toml uses
the ABS LGA names, which carry a state suffix such as "(Qld)"; the QGSO
tables use the bare name.
"""

from fetchers.fetch_population_nrw import REGIONS, in_lgas, lga_key


def test_lga_key_drops_state_suffix_and_case():
    assert lga_key("Central Highlands (Qld)") == "central highlands"
    assert lga_key("Isaac (R)") == "isaac"
    assert lga_key("  Maranoa ") == "maranoa"


def test_lga_key_keeps_other_brackets():
    assert lga_key("Whitsunday (Bowen only)") == "whitsunday (bowen only)"


def test_in_lgas_matches_abs_name_to_qgso_name():
    assert in_lgas("Central Highlands (Qld)", ["Isaac", "Central Highlands"])
    assert not in_lgas("Murweh", ["Isaac", "Central Highlands"])


def test_emerald_lga_is_in_the_bowen_basin():
    assert in_lgas("Central Highlands (Qld)", REGIONS["bowen_basin"]["lgas"])
