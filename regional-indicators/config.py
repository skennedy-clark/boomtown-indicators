"""
regional-indicators/config.py

Project configuration.

Loads and validates towns.toml, exposing towns as typed Town objects, and
maintains the index of downloaded files (cache/index.json) that lets
fetchers reuse an earlier download.
"""

from __future__ import annotations

# Range of calendar years handled by the pipeline.
YEAR_START = 2000
YEAR_END   = 2025   # last data year; update at the start of each annual cycle.
                     # transform/to_csv.py defines the same constant and must be kept in step.

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
    qgso_sa2:         str        = ""   # SA2 code used in QGSO queries
    qgso_lga:         str        = ""   # QGSO LGA identifier, e.g. "LGA/34860"
    lga_code:         str        = ""   # ABS LGA code, digits only, e.g. "15750";
                                        # needed only where qgso_lga is absent
                                        # (towns outside Queensland). See abs_lga_code.
    bom_station:      str        = ""   # Bureau of Meteorology station number
    csg_notice_year:  int        = 0
    benchmark:        bool       = False
    notes:            str        = ""

    # Derived values
    @property
    def slug(self) -> str:
        """File-system-safe form of the name: 'Toowoomba (West)' -> 'toowoomba_west'."""
        return (
            self.name.lower()
            .replace(" ", "_")
            .replace("(", "")
            .replace(")", "")
            .replace("-", "_")
        )

    @property
    def abs_lga_code(self) -> str:
        """ABS LGA code (digits only) of this town's LGA, or "".

        Taken from lga_code if set, otherwise from the digits of qgso_lga
        ("LGA/37310" -> "37310"); QGSO uses ABS LGA codes.
        """
        if self.lga_code:
            return self.lga_code
        if self.qgso_lga.startswith("LGA/"):
            return self.qgso_lga[4:]
        return ""

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
        raw_towns = raw.get("towns", {})

        if isinstance(raw_towns, dict):
            town_records = raw_towns.items()
        elif isinstance(raw_towns, list):
            town_records = (
                (f"entry_{index}", town)
                for index, town in enumerate(raw_towns)
            )
        else:
            raise TypeError(
                "'towns' in towns.toml must be either a table of named towns "
                "or an array of town tables"
            )

        towns: list[Town] = []

        for config_key, t in town_records:
            if not isinstance(t, dict):
                raise TypeError(
                    f"Town '{config_key}' must be a TOML table, "
                    f"but found {type(t).__name__}"
                )

            try:
                town = Town(
                    name=t["name"],
                    state=t["state"],
                    postcode=t.get("postcode", ""),
                    postcodes=t.get("postcodes", []),
                    sa2_code=t.get("sa2_code", ""),
                    sa2_name=t.get("sa2_name", ""),
                    sa3_code=t.get("sa3_code", ""),
                    lga=t.get("lga", ""),
                    qps_division=t.get("qps_division", ""),
                    qgso_sa2=t.get("qgso_sa2", ""),
                    qgso_lga=t.get("qgso_lga", ""),
                    lga_code=str(t.get("lga_code", "")),
                    bom_station=t.get("bom_station", ""),
                    csg_notice_year=t.get("csg_notice_year", 0),
                    benchmark=t.get("benchmark", False),
                    notes=t.get("notes", ""),
                )
            except KeyError as exc:
                missing_field = exc.args[0]
                raise ValueError(
                    f"Town '{config_key}' is missing required field "
                    f"'{missing_field}'"
                ) from exc

            towns.append(town)

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

        # Town names must be unique.
        seen = set()
        for n in names:
            if n in seen:
                errors.append(f"Duplicate town name: '{n}'")
            seen.add(n)

        # Required fields for study towns.
        for t in self.towns:
            if t.benchmark:
                continue
            if not t.sa2_code:
                errors.append(f"[{t.name}] missing sa2_code")
            if not t.postcodes:
                errors.append(f"[{t.name}] missing postcodes")
            if t.state not in ("QLD", "NSW", "VIC", "WA", "SA", "TAS", "NT", "ACT"):
                errors.append(f"[{t.name}] unknown state '{t.state}'")

        # LGA codes must be numeric, and a code must not appear under two names.
        lga_name_by_code: dict[str, str] = {}
        for t in self.towns:
            code = t.abs_lga_code
            if not code:
                continue
            if not code.isdigit():
                errors.append(
                    f"[{t.name}] LGA code must be digits only (lga_code = \"15750\" "
                    f"or qgso_lga = \"LGA/37310\"), got '{code}'"
                )
            previous = lga_name_by_code.setdefault(code, t.lga)
            if previous != t.lga:
                errors.append(
                    f"[{t.name}] LGA code {code} is named '{t.lga}' here but "
                    f"'{previous}' on another town"
                )

        if errors:
            raise ValueError("towns.toml validation errors:\n  " + "\n  ".join(errors))

    # ── Queries ────────────────────────────────────────────────────────────────

    def study_towns(self) -> list[Town]:
        """Return all towns that are not benchmarks."""
        return [t for t in self.towns if not t.benchmark]

    def towns_by_state(self, state: str) -> list[Town]:
        return [t for t in self.towns if t.state == state]

    def town_by_name(self, name: str) -> Optional[Town]:
        for t in self.towns:
            if t.name == name:
                return t
        return None

    def lgas(self) -> list[tuple[str, str, str]]:
        """Return the distinct LGAs of all towns, benchmarks included.

        Each is (LGA name, ABS LGA code, state), in towns.toml order. An
        LGA shared by several towns is returned once. Towns without an
        LGA code are skipped.
        """
        seen: set[str] = set()
        result: list[tuple[str, str, str]] = []
        for t in self.towns:
            code = t.abs_lga_code
            if not code or not t.lga or code in seen:
                continue
            seen.add(code)
            result.append((t.lga, code, t.state))
        return result

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
    """Index of downloaded source files, stored in cache/index.json.

    Each entry records the local path, download time, source URL and an
    MD5 checksum:
        {
            "<key>": {
                "path": "cache/<file>",
                "downloaded_at": "2026-03-26T14:30:00",
                "url": "https://...",
                "checksum": "..."
            }
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
        """True if `key` is indexed and its file still exists."""
        if key not in self._data:
            return False
        cached_path = Path(self._data[key]["path"])
        return cached_path.exists()

    def get_path(self, key: str) -> Optional[Path]:
        if not self.has(key):
            return None
        return Path(self._data[key]["path"])

    def register(self, key: str, path: Path, url: str = "", meta: dict = None):
        """Add a downloaded file to the index."""
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
        """Remove `key` from the index so that it is downloaded again."""
        if key in self._data:
            del self._data[key]
            self._save()

    def list_entries(self) -> dict:
        return dict(self._data)


# ── Shared instances ───────────────────────────────────────────────────────────

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