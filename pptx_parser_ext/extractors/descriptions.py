"""
descrpition.py — OOP wrapper for extracting image descriptions from PPTX slides.

Behavior preserved from original `extract_picture_descriptions` function:
- Scans each slide XML for <p:cNvPr> and reads the 'descr' attribute.
- Produces: [{"slide": <int>, "descriptions": [<str>, ...]}, ...]
"""

from __future__ import annotations

import io
import zipfile
import logging
from lxml import etree
from typing import List, Dict

logger = logging.getLogger(__name__)

# Namespaces used in XPath queries
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


class DescriptionExtractor:
    """Extract image descriptions (p:cNvPr/@descr) per slide."""

    def extract(self, pptx_bytes: bytes) -> List[Dict]:
        """
        Extracts image descriptions from a .pptx file's slide XML (p:cNvPr tags).

        Args:
            pptx_bytes: Binary content of uploaded PPTX file.

        Returns:
            list[dict]: Each dict contains:
                - "slide": slide index (1-based, sorted by slide number)
                - "descriptions": list[str] of found descriptions (or "(No description)")
        Raises:
            Exception: If parsing fails or pptx structure is invalid.
        """
        slides_output: List[Dict] = []

        try:
            with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as pptx_zip:
                slide_files = sorted(
                    [f for f in pptx_zip.namelist() if f.startswith("ppt/slides/slide") and f.endswith(".xml")],
                    key=lambda x: int("".join(filter(str.isdigit, x))),
                )
                logger.info(f"Found {len(slide_files)} slide(s) to scan for descriptions")

                for index, slide_file in enumerate(slide_files, start=1):
                    slide_descriptions = []
                    with pptx_zip.open(slide_file) as file:
                        tree = etree.parse(file)
                        for pic in tree.xpath("//p:cNvPr", namespaces=NS):
                            descr = pic.get("descr")
                            desc = descr if descr else "(No description)"
                            slide_descriptions.append(desc)

                    slides_output.append({"slide": index, "descriptions": slide_descriptions})

            return slides_output

        except Exception:
            logger.exception("Error occurred while extracting descriptions")
            raise
