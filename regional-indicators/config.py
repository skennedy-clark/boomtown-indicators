"""
config.py
---------
Loads and validates towns.toml, exposes typed Town objects and source settings.
Also manages the cache index so fetchers can check whether a file is already
downloaded before hitting a remote source.
"""

from __future__ import annotations

# Year range for output CSVs
YEAR_START = 2000
YEAR_END   = 2025   # update each cycle

import hashlib
import json
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

try:
    import tomllib          # Python 3.11+
except ImportError:
    try:
        import tomli as tomllib   # pip install tomli
    except ImportError:
        raise ImportError(
            "tomllib not found. On Python < 3.11 install with: pip install tomli"
        )

SILO_EMAIL = "uqsken12@uq.edu.au"

# ── Paths ──────────────────────────────────────────────────────────────────────

PROJECT_ROOT = Path(__file__).parent
TOML_PATH    = PROJECT_ROOT / "towns.toml"
CACHE_DIR    = PROJECT_ROOT / "cache"
CACHE_INDEX  = CACHE_DIR / "index.json"
LOG_DIR      = PROJECT_ROOT / "logs"
OUTPUT_DIR   = PROJECT_ROOT / "output"

for _d in (CACHE_DIR, LOG_DIR, OUTPUT_DIR):
    _d.mkdir(parents=True, exist_ok=True)


# ── Dataclasses ────────────────────────────────────────────────────────────────

@dataclass
class Town:
    name:             str
    state:            str
    postcode:         str
    postcodes:        list[str]
    sa2_code:         str
    sa2_name:         str
    sa3_code:         str
    lga:              str
    qps_division:     str        = ""
    qgso_sa2:         str        = ""   # ASGS Ed 2 SA2 code used by QGSO files
    qgso_lga:         str        = ""   # QGSO LGA identifier e.g. "LGA/34860"
    bom_station:      str        = ""   # BOM rainfall station number
    csg_notice_year:  int        = 0
    benchmark:        bool       = False
    notes:            str        = ""

    # Derived
    @property
    def slug(self) -> str:
        """URL/filesystem safe name, e.g. 'Toowoomba (West)' → 'toowoomba_west'"""
        return (
            self.name.lower()
            .replace(" ", "_")
            .replace("(", "")
            .replace(")", "")
            .replace("-", "_")
        )

    @property
    def output_dir(self) -> Path:
        return OUTPUT_DIR / self.name

    @property
    def is_vic(self) -> bool:
        return self.state == "VIC"

    @property
    def is_qld(self) -> bool:
        return self.state == "QLD"

    @property
    def is_nsw(self) -> bool:
        return self.state == "NSW"


@dataclass
class SourceConfig:
    name:      str
    base_url:  str = ""
    api_url:   str = ""
    states:    list[str] = field(default_factory=list)
    notes:     str = ""
    extra:     dict = field(default_factory=dict)


# ── Loader ─────────────────────────────────────────────────────────────────────

class Config:
    def __init__(self, toml_path: Path = TOML_PATH):
        with open(toml_path, "rb") as f:
            raw = tomllib.load(f)

        self.settings: dict         = raw.get("settings", {})
        self.towns:    list[Town]   = self._load_towns(raw)
        self.sources:  dict[str, SourceConfig] = self._load_sources(raw)

        self._validate()

    # ── Loading ────────────────────────────────────────────────────────────────

    def _load_towns(self, raw: dict) -> list[Town]:
        towns = []
        # towns.toml uses [towns.slug] dict-of-tables syntax, so raw["towns"]
        # parses to {"roma": {...}, "chinchilla": {...}, ...} -- iterate
        # .values() for the per-town dicts, not the dict itself (which
        # would just yield the slug strings, e.g. "roma", "chinchilla").
        for t in raw.get("towns", {}).values():
            towns.append(Town(
                name            = t["name"],
                state           = t["state"],
                postcode        = t.get("postcode", ""),
                postcodes       = t.get("postcodes", []),
                sa2_code        = t.get("sa2_code", ""),
                sa2_name        = t.get("sa2_name", ""),
                sa3_code        = t.get("sa3_code", ""),
                lga             = t.get("lga", ""),
                qps_division    = t.get("qps_division", ""),
                qgso_sa2        = t.get("qgso_sa2", ""),
                qgso_lga        = t.get("qgso_lga", ""),
                bom_station     = t.get("bom_station", ""),
                csg_notice_year = t.get("csg_notice_year", 0),
                benchmark       = t.get("benchmark", False),
                notes           = t.get("notes", ""),
            ))
        return towns

    def _load_sources(self, raw: dict) -> dict[str, SourceConfig]:
        sources = {}
        for key, val in raw.get("sources", {}).items():
            sources[key] = SourceConfig(
                name     = key,
                base_url = val.get("base_url", ""),
                api_url  = val.get("api_url", ""),
                states   = val.get("states", []),
                notes    = val.get("notes", ""),
                extra    = {k: v for k, v in val.items()
                            if k not in ("base_url", "api_url", "states", "notes")},
            )
        return sources

    # ── Validation ─────────────────────────────────────────────────────────────

    def _validate(self):
        errors = []
        names  = [t.name for t in self.towns]

        # Duplicate names
        seen = set()
        for n in names:
            if n in seen:
                errors.append(f"Duplicate town name: '{n}'")
            seen.add(n)

        # Missing critical fields for non-benchmark towns
        for t in self.towns:
            if t.benchmark:
                continue
            if not t.sa2_code:
                errors.append(f"[{t.name}] missing sa2_code")
            if not t.postcodes:
                errors.append(f"[{t.name}] missing postcodes")
            if t.state not in ("QLD", "NSW", "VIC", "WA", "SA", "TAS", "NT", "ACT"):
                errors.append(f"[{t.name}] unknown state '{t.state}'")

        if errors:
            raise ValueError("towns.toml validation errors:\n  " + "\n  ".join(errors))

    # ── Convenience queries ────────────────────────────────────────────────────

    def study_towns(self) -> list[Town]:
        """All non-benchmark towns."""
        return [t for t in self.towns if not t.benchmark]

    def towns_by_state(self, state: str) -> list[Town]:
        return [t for t in self.towns if t.state == state]

    def town_by_name(self, name: str) -> Optional[Town]:
        for t in self.towns:
            if t.name == name:
                return t
        return None

    def qld_study_towns(self) -> list[Town]:
        return [t for t in self.study_towns() if t.is_qld]

    def __repr__(self) -> str:
        return (
            f"<Config: {len(self.towns)} towns "
            f"({len(self.study_towns())} study, "
            f"{len(self.towns) - len(self.study_towns())} benchmark)>"
        )


# ── Cache management ───────────────────────────────────────────────────────────

class CacheIndex:
    """
    Tracks which raw files have been downloaded and when.
    Stored as a simple JSON file: cache/index.json

    Structure:
        {
            "ato_table8_2022-23": {
                "path": "cache/ato_table8_2022-23.xlsx",
                "downloaded_at": "2026-03-26T14:30:00",
                "url": "https://...",
                "checksum": "abc123"
            },
            ...
        }
    """

    def __init__(self, index_path: Path = CACHE_INDEX):
        self.path = index_path
        self._data: dict = self._load()

    def _load(self) -> dict:
        if self.path.exists():
            with open(self.path) as f:
                return json.load(f)
        return {}

    def _save(self):
        with open(self.path, "w") as f:
            json.dump(self._data, f, indent=2, default=str)

    def has(self, key: str) -> bool:
        """Check if a cached file exists and is still on disk."""
        if key not in self._data:
            return False
        cached_path = Path(self._data[key]["path"])
        return cached_path.exists()

    def get_path(self, key: str) -> Optional[Path]:
        if not self.has(key):
            return None
        return Path(self._data[key]["path"])

    def register(self, key: str, path: Path, url: str = "", meta: dict = None):
        """Record a newly downloaded file in the index."""
        from datetime import datetime
        checksum = ""
        if path.exists():
            checksum = hashlib.md5(path.read_bytes()).hexdigest()

        self._data[key] = {
            "path":          str(path),
            "downloaded_at": datetime.now().isoformat(timespec="seconds"),
            "url":           url,
            "checksum":      checksum,
            **(meta or {}),
        }
        self._save()

    def invalidate(self, key: str):
        """Force re-download on next run."""
        if key in self._data:
            del self._data[key]
            self._save()

    def list_entries(self) -> dict:
        return dict(self._data)


# ── Module-level singletons (imported by fetchers) ────────────────────────────

_config_instance: Optional[Config]     = None
_cache_instance:  Optional[CacheIndex] = None


def get_config() -> Config:
    global _config_instance
    if _config_instance is None:
        _config_instance = Config()
    return _config_instance


def get_cache() -> CacheIndex:
    global _cache_instance
    if _cache_instance is None:
        _cache_instance = CacheIndex()
    return _cache_instance