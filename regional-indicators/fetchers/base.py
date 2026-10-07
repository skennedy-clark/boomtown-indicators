"""
regional-indicators/fetchers/base.py

Abstract base class for the data fetchers.

Provides configuration, caching, logging, downloading with retry and a
uniform result object, so that run_update.py can run and report on every
fetcher in the same way.
"""

from __future__ import annotations

import time
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from config import Config, CacheIndex, Town, get_config, get_cache, CACHE_DIR
from logger import get_child_logger


# ── Result ─────────────────────────────────────────────────────────────────────

@dataclass
class FetchResult:
    """Result of a fetcher run, returned by BaseFetcher.run() and used by
    run_update.py to build its summary.
    """
    source:        str
    success:       bool
    towns_ok:      list[str]  = field(default_factory=list)
    towns_skipped: list[str]  = field(default_factory=list)   # served from the cache
    towns_failed:  list[str]  = field(default_factory=list)
    cached_files:  list[Path] = field(default_factory=list)
    new_files:     list[Path] = field(default_factory=list)
    errors:        list[str]  = field(default_factory=list)
    warnings:      list[str]  = field(default_factory=list)
    elapsed_s:     float      = 0.0

    def add_error(self, town: str, msg: str):
        self.towns_failed.append(town)
        self.errors.append(f"[{town}] {msg}")

    def add_warning(self, town: str, msg: str):
        self.warnings.append(f"[{town}] {msg}")

    def summary(self) -> str:
        lines = [
            f"Source : {self.source}",
            f"Status : {'✓ OK' if self.success else '✗ FAILED'}",
            f"Towns  : {len(self.towns_ok)} ok, "
            f"{len(self.towns_skipped)} cached, "
            f"{len(self.towns_failed)} failed",
            f"Time   : {self.elapsed_s:.1f}s",
        ]
        if self.errors:
            lines.append("Errors :")
            lines.extend(f"  {e}" for e in self.errors)
        if self.warnings:
            lines.append("Warnings:")
            lines.extend(f"  {w}" for w in self.warnings)
        return "\n".join(lines)


# ── Base fetcher ───────────────────────────────────────────────────────────────

class BaseFetcher(ABC):
    """Base class for fetchers.

    A subclass sets:
        SOURCE_NAME        short identifier, used in logs and summaries
        SUPPORTED_STATES   states the source covers; empty means all
    and implements fetch_all(), which records its outcome in
    self.result.

    Available to subclasses:
        self.log            logger for this fetcher
        self.config         the loaded Config
        self.cache          the CacheIndex
        self.cache_path()   path for a cache file
        self.is_cached()    whether a key is in the cache
        self.download()     download a URL into the cache, with retry
    """

    SOURCE_NAME:      str       = "base"
    SUPPORTED_STATES: list[str] = []   # empty = all states

    def __init__(self, force: bool = False):
        self.config: Config     = get_config()
        self.cache:  CacheIndex = get_cache()
        self.force:  bool       = force
        self.log = get_child_logger("regional-indicators", self.SOURCE_NAME)
        self.result = FetchResult(source=self.SOURCE_NAME, success=False)

    # ── Public entry point ─────────────────────────────────────────────────────

    def run(self) -> FetchResult:
        start = time.time()
        self.log.info(f"Starting {self.SOURCE_NAME}")
        try:
            self.fetch_all()
            self.result.success = len(self.result.towns_failed) == 0
        except Exception as exc:
            self.log.error(f"Unhandled exception in {self.SOURCE_NAME}: {exc}",
                           exc_info=True)
            self.result.success = False
            self.result.errors.append(f"Unhandled: {exc}")
        finally:
            self.result.elapsed_s = time.time() - start
            self.log.info(self.result.summary())
        return self.result

    # ── Abstract ───────────────────────────────────────────────────────────────

    @abstractmethod
    def fetch_all(self):
        """Fetch the source and write the cache files.

        Implementations add each town or region to self.result.towns_ok,
        towns_failed or towns_skipped.
        """

    # ── Helpers ────────────────────────────────────────────────────────────────

    def applicable_towns(self) -> list[Town]:
        """Return the study towns in SUPPORTED_STATES."""
        towns = self.config.study_towns()
        if self.SUPPORTED_STATES:
            towns = [t for t in towns if t.state in self.SUPPORTED_STATES]
        return towns

    def cache_path(self, key: str, suffix: str = "") -> Path:
        """Return the path for a cache file, creating the cache directory if needed."""
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        fname = key if not suffix else f"{key}{suffix}"
        return CACHE_DIR / fname

    def is_cached(self, key: str) -> bool:
        return self.cache.has(key)

    def download(
        self,
        url:        str,
        cache_key:  str,
        suffix:     str      = "",
        retries:    int      = 3,
        backoff_s:  float    = 2.0,
        force:      bool     = False,
        headers:    dict     = None,
        params:     dict     = None,
    ) -> Optional[Path]:
        """Download `url` into the cache and return the local path, or None on
        failure. An existing cached copy is reused unless force is set.
        """
        import requests

        if not force and not self.force and self.is_cached(cache_key):
            path = self.cache.get_path(cache_key)
            self.log.debug(f"Cache hit: {cache_key} → {path}")
            return path

        dest = self.cache_path(cache_key, suffix)
        self.log.info(f"Downloading {url} → {dest.name}")

        for attempt in range(1, retries + 1):
            try:
                resp = requests.get(
                    url,
                    headers=headers or {},
                    params=params or {},
                    timeout=60,
                    stream=True,
                )
                resp.raise_for_status()

                with open(dest, "wb") as f:
                    for chunk in resp.iter_content(chunk_size=8192):
                        f.write(chunk)

                self.cache.register(cache_key, dest, url=url)
                self.result.new_files.append(dest)
                self.log.info(f"  Saved {dest.stat().st_size / 1024:.1f} KB")
                return dest

            except Exception as exc:
                self.log.warning(
                    f"  Attempt {attempt}/{retries} failed: {exc}"
                )
                if attempt < retries:
                    time.sleep(backoff_s * attempt)

        self.log.error(f"All {retries} attempts failed for {url}")
        return None
