"""
transform/booklet/make_booklet.py
----------------------------------
Entry point for booklet generation.

Page orientation:
    Pages 1-2  (cover, title)   : Portrait  A4
    Pages 3-10 (map, data pages): Landscape A4
    Final page (back)           : Portrait  A4  [TODO]

Usage:
    python transform/booklet/make_booklet.py --town Chinchilla
    python transform/booklet/make_booklet.py --town Roma --date "June 2026"

Output:
    booklets/{slug}/{TownName}_Indicators_{year}.docx
"""
from __future__ import annotations

import argparse
import sys
import tomllib
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.shared import Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.section import WD_ORIENTATION

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT  = SCRIPT_DIR.parent.parent

# Ensure the booklet directory is on sys.path so 'common' and 'pages'
# are importable regardless of which directory the script is run from.
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from common import BOOKLETS, MARGIN_MM
from pages.cover      import build_cover_page
from pages.title      import build_title_page
from pages.map        import build_map_page
from pages.population import build_population_page
from pages.data_page  import (build_unemployment_page, build_house_price_page,
                               build_rent_page, build_approvals_page,
                               build_rainfall_page, build_crime_page)

A4_SHORT = int(210 / 25.4 * 914400)
A4_LONG  = int(297 / 25.4 * 914400)
MARGIN   = Cm(MARGIN_MM / 10)


def load_town_cfg(town_name: str) -> dict:
    toml_path = REPO_ROOT / "towns.toml"
    with open(toml_path, "rb") as f:
        cfg = tomllib.load(f)
    for slug, town in cfg.get("towns", {}).items():
        if town.get("name", "").lower() == town_name.lower() or slug.lower() == town_name.lower():
            return {
                "slug":         slug,
                "name":         town.get("name", town_name),
                "sa2_code":     str(town.get("sa2_code", "")),
                "sa2_name":     town.get("sa2_name", town.get("name", town_name)),
                "district":     town.get("district", f"{town.get('name', town_name)} and District"),
                "cover_photo":  town.get("cover_photo", "cover.jpg"),
                "bom_station":  str(town.get("bom_station", "")),
                "qps_division": town.get("qps_division", ""),
            }
    raise ValueError(f"Town '{town_name}' not found in towns.toml")


def _set_section(section, portrait: bool):
    if portrait:
        section.page_width  = A4_SHORT
        section.page_height = A4_LONG
        section.orientation = WD_ORIENTATION.PORTRAIT
    else:
        section.page_width  = A4_LONG
        section.page_height = A4_SHORT
        section.orientation = WD_ORIENTATION.LANDSCAPE
    section.left_margin   = MARGIN
    section.right_margin  = MARGIN
    section.top_margin    = MARGIN
    section.bottom_margin = MARGIN


def _add_section_break(doc: Document, portrait: bool) -> None:
    """Insert a paragraph whose sectPr switches the page orientation."""
    p   = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    sectPr = OxmlElement("w:sectPr")

    pgSz = OxmlElement("w:pgSz")
    if portrait:
        pgSz.set(qn("w:w"), str(A4_SHORT))
        pgSz.set(qn("w:h"), str(A4_LONG))
    else:
        pgSz.set(qn("w:w"),      str(A4_LONG))
        pgSz.set(qn("w:h"),      str(A4_SHORT))
        pgSz.set(qn("w:orient"), "landscape")
    sectPr.append(pgSz)

    m = str(int(MARGIN_MM / 25.4 * 1440))
    pgMar = OxmlElement("w:pgMar")
    pgMar.set(qn("w:top"),    m)
    pgMar.set(qn("w:right"),  m)
    pgMar.set(qn("w:bottom"), m)
    pgMar.set(qn("w:left"),   m)
    pgMar.set(qn("w:header"), "720")
    pgMar.set(qn("w:footer"), "720")
    pgMar.set(qn("w:gutter"), "0")
    sectPr.append(pgMar)

    pPr.append(sectPr)
    p.paragraph_format.space_before = None
    p.paragraph_format.space_after  = None


def _plain_page_break(doc: Document):
    p   = doc.add_paragraph()
    p.paragraph_format.space_before = None
    p.paragraph_format.space_after  = None
    run = p.add_run()
    br  = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)


def new_document() -> Document:
    doc = Document()
    _set_section(doc.sections[0], portrait=True)
    for p in list(doc.paragraphs):
        p._element.getparent().remove(p._element)
    return doc


def main():
    parser = argparse.ArgumentParser(description="Generate town indicator booklet")
    parser.add_argument("--town",  required=True)
    parser.add_argument("--date",  default=None)
    parser.add_argument("--pages", nargs="+", default=["all"])
    args = parser.parse_args()

    date_str  = args.date or datetime.now().strftime("%B %Y")
    town_cfg  = load_town_cfg(args.town)
    slug      = town_cfg["slug"]
    pages     = args.pages
    build_all = "all" in pages

    print(f"Generating booklet for: {town_cfg['name']} ({slug})")
    print(f"Date: {date_str}")

    doc = new_document()   # Section 1: Portrait

    # ── PORTRAIT: pages 1 & 2 ────────────────────────────────────────────────
    if build_all or "cover" in pages:
        print("  [portrait]  Page 1: cover...")
        build_cover_page(doc, town_cfg, date_str)

    if build_all or "title" in pages:
        _plain_page_break(doc)
        print("  [portrait]  Page 2: title...")
        build_title_page(doc, town_cfg, date_str)

    # ── Switch to LANDSCAPE ───────────────────────────────────────────────────
    _add_section_break(doc, portrait=False)

    if build_all or "map" in pages:
        print("  [landscape] Page 3: map...")
        build_map_page(doc, town_cfg)

    if build_all or "population" in pages:
        _plain_page_break(doc)
        print("  [landscape] Page 4: population...")
        build_population_page(doc, town_cfg)

    if build_all or "unemployment" in pages:
        _plain_page_break(doc)
        print("  [landscape] Page 5: unemployment...")
        build_unemployment_page(doc, town_cfg)

    if build_all or "housing" in pages:
        _plain_page_break(doc)
        print("  [landscape] Page 6: housing...")
        build_house_price_page(doc, town_cfg)

    if build_all or "rent" in pages:
        _plain_page_break(doc)
        print("  [landscape] Page 7: rent...")
        build_rent_page(doc, town_cfg)

    if build_all or "approvals" in pages:
        _plain_page_break(doc)
        print("  [landscape] Page 8: approvals...")
        build_approvals_page(doc, town_cfg)

    if build_all or "rainfall" in pages:
        _plain_page_break(doc)
        print("  [landscape] Page 9: rainfall...")
        build_rainfall_page(doc, town_cfg)

    if build_all or "crime" in pages:
        _plain_page_break(doc)
        print("  [landscape] Page 10: crime...")
        build_crime_page(doc, town_cfg)

    # ── Switch back to PORTRAIT for back page (TODO: build_back_page) ─────────
    # _add_section_break(doc, portrait=True)
    # build_back_page(doc, town_cfg)

    # ── Save ──────────────────────────────────────────────────────────────────
    year     = datetime.now().year
    out_dir  = BOOKLETS / slug
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{town_cfg['name']}_Indicators_{year}.docx"
    doc.save(str(out_path))
    print(f"\nSaved: {out_path}")


if __name__ == "__main__":
    main()