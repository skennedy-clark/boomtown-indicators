"""
fetchers/ato_release.py
-----------------------
Discovers the newest ATO Taxation Statistics release on data.gov.au that
contains both Individuals Table 6 and Individuals Table 8.

The data.gov.au catalogue search index can lag behind newly published datasets,
so this module probes likely package slugs directly using CKAN package_show.

Example package slug:
    taxation-statistics-2023-24
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from datetime import date
from functools import lru_cache

import requests


PACKAGE_SHOW_URL = "https://data.gov.au/data/api/3/action/package_show"

# Oldest release worth probing. This can be moved further back if required.
EARLIEST_START_YEAR = 2015

REQUEST_TIMEOUT_S = 30

log = logging.getLogger("regional-indicators.ato_release")


@dataclass(frozen=True)
class ATORelease:
    financial_year: str
    slug: str
    title: str
    modified: str
    resources: tuple[dict, ...]


def _financial_year(start_year: int) -> str:
    """
    Convert 2023 to '2023-24'.
    """
    return f"{start_year}-{(start_year + 1) % 100:02d}"


def _latest_possible_start_year() -> int:
    """
    Return the start year of the latest completed Australian financial year.

    Before 1 July 2026, the latest completed financial year is 2024-25.
    From 1 July 2026, it becomes 2025-26.

    The latest published ATO release will normally lag behind this.
    """
    today = date.today()

    if today.month >= 7:
        return today.year - 1

    return today.year - 2


def _candidate_financial_years() -> list[str]:
    latest_start = _latest_possible_start_year()

    return [
        _financial_year(start_year)
        for start_year in range(
            latest_start,
            EARLIEST_START_YEAR - 1,
            -1,
        )
    ]


def _resource_text(resource: dict) -> str:
    """
    Combine resource metadata into searchable lowercase text.
    """
    fields = [
        resource.get("name", ""),
        resource.get("description", ""),
        resource.get("url", ""),
        resource.get("format", ""),
    ]

    return " ".join(str(value) for value in fields).lower()


def _resource_is_table(resource: dict, table_number: int) -> bool:
    """
    Detect an ATO Individuals table using both resource name and filename.

    Examples:
        Individuals Table 6
        Table 6A
        ts24individual06taxablestatusstatesa4postcode.xlsx
        ts24individual08medianaveragetaxableincomestatepostcode.xlsx
    """
    text = _resource_text(resource)
    padded = f"{table_number:02d}"

    explicit_patterns = [
        rf"\btable\s*{table_number}\b",
        rf"\btable\s*{padded}\b",
        rf"individual{padded}\b",
        rf"individual[_\-\s]*{padded}\b",
    ]

    if not any(re.search(pattern, text) for pattern in explicit_patterns):
        return False

    # The project specifically requires postcode-level Individuals tables.
    return "postcode" in text and "individual" in text


def _contains_required_tables(resources: list[dict]) -> bool:
    has_table6 = any(_resource_is_table(r, 6) for r in resources)
    has_table8 = any(_resource_is_table(r, 8) for r in resources)

    return has_table6 and has_table8


def _fetch_package(financial_year: str) -> ATORelease | None:
    slug = f"taxation-statistics-{financial_year}"

    try:
        response = requests.get(
            PACKAGE_SHOW_URL,
            params={"id": slug},
            timeout=REQUEST_TIMEOUT_S,
            headers={
                "User-Agent": (
                    "boomtown-indicators/1.0 "
                    "(UQ regional indicators research pipeline)"
                )
            },
        )
    except requests.RequestException as exc:
        log.warning(
            "ATO package request failed for %s: %s",
            financial_year,
            exc,
        )
        return None

    try:
        payload = response.json()
    except requests.JSONDecodeError:
        # data.gov.au can return HTML or an empty response for a missing slug.
        log.debug(
            "ATO package %s not found or returned non-JSON content",
            slug,
        )
        return None

    if not payload.get("success"):
        return None

    dataset = payload.get("result", {})
    resources = dataset.get("resources", [])

    if not isinstance(resources, list):
        log.warning(
            "ATO package %s returned an invalid resources structure",
            slug,
        )
        return None

    if not _contains_required_tables(resources):
        log.info(
            "ATO package %s exists but does not contain both Tables 6 and 8",
            slug,
        )
        return None

    return ATORelease(
        financial_year=financial_year,
        slug=slug,
        title=str(dataset.get("title", "")),
        modified=str(dataset.get("metadata_modified", "")),
        resources=tuple(resources),
    )


@lru_cache(maxsize=1)
def discover_latest_ato_release() -> ATORelease:
    """
    Return the newest Taxation Statistics package containing Tables 6 and 8.

    The result is cached in memory so both ATO fetchers use the same release
    without repeating the full sequence of package probes.
    """
    candidates = _candidate_financial_years()

    log.info(
        "Checking ATO Taxation Statistics releases: %s",
        ", ".join(candidates),
    )

    for financial_year in candidates:
        release = _fetch_package(financial_year)

        if release is not None:
            log.info(
                "Newest usable ATO release: %s "
                "(slug=%s, modified=%s)",
                release.financial_year,
                release.slug,
                release.modified or "unknown",
            )
            return release

        log.debug("ATO release %s not available", financial_year)

    raise RuntimeError(
        "Could not find an ATO Taxation Statistics package containing "
        "both Individuals Table 6 and Individuals Table 8. "
        f"Checked financial years: {', '.join(candidates)}"
    )