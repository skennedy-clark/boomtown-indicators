"""
regional-indicators/fetchers/pdf_text.py

The text of a PDF report as lines, page by page, read with pypdfium2
(a project dependency, as in fetch_fuel.py). Table rows come out as one
line each ("Series A 4,100 4,200 4,230 4,230 4,120 4,120"), which is
what the QGSO table parsers rely on.
"""

from __future__ import annotations

from pathlib import Path


def page_lines(path: Path) -> list[list[str]]:
    """[[line, ...] for each page], lines stripped, empty lines dropped."""
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(str(path))
    try:
        pages = []
        for index in range(len(pdf)):
            text = pdf[index].get_textpage().get_text_range()
            pages.append([line.strip() for line in text.replace("\r", "").split("\n") if line.strip()])
        return pages
    finally:
        pdf.close()
