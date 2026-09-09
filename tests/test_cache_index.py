"""
test_cache_index.py -- CacheIndex behaviour, isolated from the real
cache/index.json via tmp_path.
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
    assert cache.list_entries()["k"]["checksum"]  # non-empty md5 hex string


def test_invalidate_removes_entry(tmp_path):
    cache = CacheIndex(tmp_path / "index.json")
    data_file = tmp_path / "data.csv"
    data_file.write_text("hello")

    cache.register("k", data_file)
    assert cache.has("k")
    cache.invalidate("k")
    assert not cache.has("k")


def test_has_false_if_file_deleted_after_registration(tmp_path):
    """A cache entry pointing at a file that's since been deleted (e.g. the
    user manually cleared cache/ without going through --force) should read
    as not-cached, not crash."""
    cache = CacheIndex(tmp_path / "index.json")
    data_file = tmp_path / "data.csv"
    data_file.write_text("hello")

    cache.register("k", data_file)
    data_file.unlink()
    assert not cache.has("k")


def test_persists_across_instances(tmp_path):
    """The index is a JSON file on disk -- a fresh CacheIndex pointed at the
    same path should see what an earlier instance registered, which is the
    whole point of caching across separate `python run_update.py` runs."""
    index_path = tmp_path / "index.json"
    data_file = tmp_path / "data.csv"
    data_file.write_text("hello")

    first = CacheIndex(index_path)
    first.register("k", data_file)

    second = CacheIndex(index_path)
    assert second.has("k")