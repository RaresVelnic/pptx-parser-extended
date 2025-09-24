# descriptions.py — OOP wrapper for extracting image/shape descriptions from PPTX slides.
#
# Behavior (default, cleaner output):
# - Scans each slide XML for <p:cNvPr> and reads the 'descr' attribute.
# - Suppresses empty/whitespace-only descriptions.
# - Skips slides that contain no valid (non-empty) descriptions.
# - Produces: [{"slide": <int>, "descriptions": [<str>, ...]}, ...]
#
# To restore the old behavior (show "(No description)" entries and include every slide),
# instantiate with: DescriptionExtractor(include_empty=True)

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
    """Extract descriptions (p:cNvPr/@descr) per slide."""

    def __init__(self, include_empty: bool = False) -> None:
        """
        Args:
            include_empty: If True, include empty descriptions as "(No description)"
                           and include slides even if they only contain empty entries.
                           Defaults to False (suppress empty entries and omit such slides).
        """
        self.include_empty = include_empty

    def extract(self, pptx_bytes: bytes) -> List[Dict]:
        """
        Extracts descriptions from a .pptx file's slide XML.

        Notes:
            We intentionally scan all <p:cNvPr> nodes, which cover pictures and many
            shape types. If you only want pictures, switch the XPath to:
            //p:pic/p:nvPicPr/p:cNvPr

        Args:
            pptx_bytes: Binary content of uploaded PPTX file.

        Returns:
            list[dict]: Each dict contains:
                - "slide": slide index (1-based, sorted by slide number)
                - "descriptions": list[str] of found descriptions (possibly empty
                  entries mapped to "(No description)" if include_empty=True)

        Raises:
            Exception: If parsing fails or pptx structure is invalid.
        """
        slides_output: List[Dict] = []

        try:
            with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as pptx_zip:
                # Sort slide parts by their numeric index
                slide_files = sorted(
                    [f for f in pptx_zip.namelist() if f.startswith("ppt/slides/slide") and f.endswith(".xml")],
                    key=lambda x: int("".join(ch for ch in x if ch.isdigit())),
                )
                logger.info(f"Found {len(slide_files)} slide(s) to scan for descriptions")

                for index, slide_file in enumerate(slide_files, start=1):
                    with pptx_zip.open(slide_file) as file:
                        tree = etree.parse(file)

                    # Scan all cNvPr nodes for @descr
                    nodes = tree.xpath("//p:cNvPr", namespaces=NS)

                    if self.include_empty:
                        # Old behavior: include empty entries as "(No description)"
                        slide_descriptions = [
                            (node.get("descr") or "").strip() or "(No description)" for node in nodes
                        ]
                        # Always include slide (even if all are "(No description)")
                        slides_output.append({"slide": index, "descriptions": slide_descriptions})
                    else:
                        # New behavior: filter out empties; include slide only if something remains
                        slide_descriptions = [
                            d for d in ((node.get("descr") or "").strip() for node in nodes) if d
                        ]
                        if slide_descriptions:
                            slides_output.append({"slide": index, "descriptions": slide_descriptions})

            return slides_output

        except Exception:
            logger.exception("Error occurred while extracting descriptions")
            raise
