# xlsx_ext/extractors/descriptions.py
"""
XlsxDescriptionExtractor — collect Alt Text (descriptions) from XLSX drawings.

What it finds:
- For each worksheet, follow its drawing rel(s) to xl/drawings/drawingN.xml.
- In those drawings, read <xdr:cNvPr descr="..."> (non-empty only).
- Try to attach the top-left anchor cell (A1 notation) for context.

Output schema per sheet (only if any descriptions exist):
[
  {"sheet_index": 1, "sheet": "Sheet 1", "descriptions": ["E3 — Figure caption", "B7 — Logo"]},
  ...
]
"""

from __future__ import annotations

import io
import zipfile
import logging
from typing import Dict, List, Optional

from lxml import etree

from .base import XlsxBaseExtractor, NS, REL_TYPES

logger = logging.getLogger(__name__)


def _col_letters(zero_based_col: int) -> str:
    """Convert 0-based column index to Excel letters (0 -> A)."""
    n = zero_based_col + 1
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _anchor_cell(anchor_el: etree._Element) -> Optional[str]:
    """
    For twoCellAnchor/oneCellAnchor, return the top-left cell like 'E3'.
    For absolute anchors, return None.
    """
    fr = anchor_el.find("xdr:from", NS)
    if fr is None:
        return None
    try:
        col = int(fr.findtext("xdr:col", default="0", namespaces=NS))
        row = int(fr.findtext("xdr:row", default="0", namespaces=NS))
        return f"{_col_letters(col)}{row + 1}"
    except Exception:
        return None


class XlsxDescriptionExtractor(XlsxBaseExtractor):
    """Extract per-sheet Alt Text descriptions from XLSX drawings."""

    def extract(self, xlsx_bytes: bytes) -> List[Dict]:
        out: List[Dict] = []
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as z:
            sheet_names = self._sheet_name_map(z)
            sheet_parts = self._sheet_parts_in_workbook_order(z)
            all_files = set(z.namelist())

            for sp in sheet_parts:
                idx = self._sheet_index(sp) or 0
                friendly = sheet_names.get(sp, f"Sheet {idx}")
                descs_for_sheet: List[str] = []

                # sheet rels -> drawing part(s)
                sheet_rels = self._read_rels(z, sp)
                drawing_targets = [
                    r["Target"] for r in sheet_rels
                    if r["Type"] == REL_TYPES["drawing"]
                ]

                for drawing_part in drawing_targets:
                    if drawing_part not in all_files:
                        continue

                    try:
                        root = self._read_xml(z, drawing_part)
                    except Exception:
                        logger.exception("Failed to parse drawing %s", drawing_part)
                        continue

                    # all anchors that can host shapes/pictures
                    for anchor in root.findall(".//xdr:twoCellAnchor", NS) + \
                                   root.findall(".//xdr:oneCellAnchor", NS) + \
                                   root.findall(".//xdr:absoluteAnchor", NS):

                        where = _anchor_cell(anchor)

                        # a single anchor can contain multiple shapes/pictures; find all cNvPr
                        for cNvPr in anchor.findall(".//xdr:cNvPr", NS):
                            descr = (cNvPr.get("descr") or "").strip()
                            if not descr:
                                continue
                            label = f"{where} — {descr}" if where else descr
                            descs_for_sheet.append(label)

                if descs_for_sheet:
                    out.append({
                        "sheet_index": idx,
                        "sheet": friendly,
                        "descriptions": descs_for_sheet
                    })

        return out
