from __future__ import annotations

from datetime import date

import requests


PACKAGE_SHOW_URL = (
    "https://data.gov.au/data/api/3/action/package_show"
)


def candidate_financial_years() -> list[str]:
    """
    Generate plausible ATO release years, newest first.

    In June 2026, candidates include:
      2025-26, 2024-25, 2023-24, ...
    """
    current_year = date.today().year

    candidates = []

    for start_year in range(current_year - 1, 2019, -1):
        end_year = (start_year + 1) % 100
        candidates.append(f"{start_year}-{end_year:02d}")

    return candidates


def inspect_package(financial_year: str) -> dict | None:
    slug = f"taxation-statistics-{financial_year}"

    try:
        response = requests.get(
            PACKAGE_SHOW_URL,
            params={"id": slug},
            timeout=30,
        )
    except requests.RequestException as exc:
        print(f"{financial_year}: request failed: {exc}")
        return None

    try:
        payload = response.json()
    except requests.JSONDecodeError:
        # data.gov.au occasionally returns HTML or an empty response
        # for package slugs that do not exist.
        return None

    if not payload.get("success"):
        return None

    dataset = payload.get("result", {})
    resources = dataset.get("resources", [])

    resource_names = [
        str(resource.get("name", ""))
        for resource in resources
    ]

    has_table6 = any(
        "individual" in name.lower()
        and "table 6" in name.lower()
        for name in resource_names
    )

    has_table8 = any(
        "individual" in name.lower()
        and "table 8" in name.lower()
        for name in resource_names
    )

    return {
        "financial_year": financial_year,
        "slug": slug,
        "title": dataset.get("title", ""),
        "modified": dataset.get("metadata_modified", ""),
        "resource_count": len(resources),
        "has_table6": has_table6,
        "has_table8": has_table8,
    }


def main() -> int:
    print("Probing ATO Taxation Statistics packages...\n")

    found = []

    for financial_year in candidate_financial_years():
        package = inspect_package(financial_year)

        if package is None:
            print(f"{financial_year}: not found")
            continue

        found.append(package)

        print(f"{financial_year}: FOUND")
        print(f"  title:      {package['title']}")
        print(f"  slug:       {package['slug']}")
        print(f"  modified:   {package['modified']}")
        print(f"  resources:  {package['resource_count']}")
        print(f"  Table 6:    {package['has_table6']}")
        print(f"  Table 8:    {package['has_table8']}")
        print()

    usable = [
        package
        for package in found
        if package["has_table6"] and package["has_table8"]
    ]

    if not usable:
        print("No package containing both Tables 6 and 8 was found.")
        return 1

    newest = usable[0]

    print("=" * 60)
    print("Newest usable release")
    print(f"Financial year: {newest['financial_year']}")
    print(f"Package slug:   {newest['slug']}")
    print("=" * 60)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())