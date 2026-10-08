"""
tests/test_cache_index.py

Tests for CacheIndex, using an index file under tmp_path in place of
cache/index.json.
"""

from config import CacheIndex


def test_register_and_has(tmp_path):
    cache = CacheIndex(tmp_path / "index.json")
    data_file = tmp_path / "data.csv"
    data_file.write_text("a,b\n1,2\n")

    assert not cache.has("mykey")
    cache.register("mykey", data_file, url="http://example.com/data.csv")
    assert cache.has("mykey")
    assert cache.get_path("mykey") == data_file


def test_checksum_recorded(tmp_path):
    cache = CacheIndex(tmp_path / "index.json")
    data_file = tmp_path / "data.csv"
    data_file.write_text("hello")

    cache.register("k", data_file)
    assert cache.list_entries()["k"]["checksum"]  # non-empty MD5 hex string


def test_invalidate_removes_entry(tmp_path):
    cache = CacheIndex(tmp_path / "index.json")
    data_file = tmp_path / "data.csv"
    data_file.write_text("hello")

    cache.register("k", data_file)
    assert cache.has("k")
    cache.invalidate("k")
    assert not cache.has("k")


def test_has_false_if_file_deleted_after_registration(tmp_path):
    """An entry whose file has since been deleted (for example, cache/
    cleared by hand instead of with --force) reads as not cached and
    does not raise.
    """
    cache = CacheIndex(tmp_path / "index.json")
    data_file = tmp_path / "data.csv"
    data_file.write_text("hello")

    cache.register("k", data_file)
    data_file.unlink()
    assert not cache.has("k")


def test_persists_across_instances(tmp_path):
    """The index is a JSON file on disk: a new CacheIndex opened on the
    same path sees entries registered by an earlier instance, as needed
    for caching across separate `python run_update.py` runs.
    """
    index_path = tmp_path / "index.json"
    data_file = tmp_path / "data.csv"
    data_file.write_text("hello")

    first = CacheIndex(index_path)
    first.register("k", data_file)

    second = CacheIndex(index_path)
    assert second.has("k")