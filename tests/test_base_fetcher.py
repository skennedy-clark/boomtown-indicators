"""
tests/test_base_fetcher.py

Tests for BaseFetcher.download(): the cache-hit skip, retry then
success, and giving up once the retries are exhausted.

requests.get and time.sleep are replaced throughout, so the suite makes
no network calls and does not wait through backoff delays.
"""

import pytest

from config import CacheIndex, Config
from fetchers.base import BaseFetcher


class DummyFetcher(BaseFetcher):
    """Minimal concrete subclass, sufficient to instantiate BaseFetcher.

    fetch_all() is never called; the tests exercise download() directly.
    """
    SOURCE_NAME = "dummy"

    def fetch_all(self):
        raise NotImplementedError("not exercised by these tests")


@pytest.fixture
def fetcher(monkeypatch, sample_toml, tmp_path):
    """A DummyFetcher with an isolated Config and CacheIndex and a
    temporary cache directory, so that the project configuration and
    cache/ folder are not touched.
    """
    config = Config(sample_toml)
    cache = CacheIndex(tmp_path / "index.json")

    monkeypatch.setattr("fetchers.base.get_config", lambda: config)
    monkeypatch.setattr("fetchers.base.get_cache", lambda: cache)
    monkeypatch.setattr("fetchers.base.CACHE_DIR", tmp_path)

    return DummyFetcher()


def test_download_skips_network_when_cached(fetcher, tmp_path):
    cached_file = tmp_path / "already_here.csv"
    cached_file.write_text("cached content")
    fetcher.cache.register("existing_key", cached_file, url="http://example.com")

    # requests.get is not replaced in this test. If download() did not
    # return from the cache before making the request, the test would fail
    # with a connection error.
    result = fetcher.download(
        "http://example.com/should-not-be-fetched", cache_key="existing_key"
    )
    assert result == cached_file


def test_download_retries_then_succeeds(fetcher, monkeypatch):
    calls = {"n": 0}

    class FakeResponse:
        def raise_for_status(self):
            pass

        def iter_content(self, chunk_size):
            yield b"hello world"

    def fake_get(url, headers=None, params=None, timeout=None, stream=None):
        calls["n"] += 1
        if calls["n"] < 3:
            raise ConnectionError("simulated network failure")
        return FakeResponse()

    monkeypatch.setattr("requests.get", fake_get)
    monkeypatch.setattr("time.sleep", lambda s: None)  # no backoff delay

    result = fetcher.download(
        "http://example.com/data.csv", cache_key="new_key", retries=3
    )

    assert calls["n"] == 3
    assert result is not None
    assert result.read_text() == "hello world"
    assert fetcher.cache.has("new_key")


def test_download_gives_up_after_retries_exhausted(fetcher, monkeypatch):
    def always_fails(*args, **kwargs):
        raise ConnectionError("simulated network failure")

    monkeypatch.setattr("requests.get", always_fails)
    monkeypatch.setattr("time.sleep", lambda s: None)

    result = fetcher.download(
        "http://example.com/data.csv", cache_key="doomed_key", retries=2
    )

    assert result is None
    assert not fetcher.cache.has("doomed_key")
    assert fetcher.result.errors == []  # download() logs the failure; it does not add to result.errors