"""
xlsx_ext.base — helpers & constants for Excel (.xlsx) parsing
"""

from __future__ import annotations

import re
import posixpath
from typing import Dict, Optional, List, Tuple
from zipfile import ZipFile
from lxml import etree

# Namespaces used across xlsx parts
NS: Dict[str, str] = {
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "ws":  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

REL_TYPES: Dict[str, str] = {
    "worksheet":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
    "drawing":    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    "hyperlink":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    "chart":      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    "extLink":    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
    "extLinkPath":"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
}


class XlsxBaseExtractor:
    """
    Base class with ZIP/XML helpers for .xlsx:
    - read XML parts
    - read .rels
    - resolve relative targets
    - sheet sorting and friendly sheet names
    """

    # ---------- low-level IO ----------
    def _rels_path(self, part_path: str) -> str:
        d, b = posixpath.split(part_path)
        return posixpath.join(d, "_rels", b + ".rels")

    def _norm_join(self, base_part: str, target: str) -> str:
        base_dir = posixpath.dirname(base_part)
        p = posixpath.normpath(posixpath.join(base_dir, target))
        return p.lstrip("/")

    def _read_xml(self, zf: ZipFile, path: str):
        with zf.open(path) as f:
            return etree.fromstring(f.read())

    def _read_rels(self, zf: ZipFile, part_path: str) -> List[Dict[str, str]]:
        rp = self._rels_path(part_path)
        if rp not in zf.namelist():
            return []
        rels = self._read_xml(zf, rp)
        out = []
        for rel in rels.findall(f".//{{{REL_NS}}}Relationship"):
            rid = rel.get("Id")
            rtype = rel.get("Type")
            target_mode = rel.get("TargetMode")  # "External" or None
            raw_target = rel.get("Target")

            # IMPORTANT: do not path-join external targets like https://...
            if target_mode == "External":
                resolved_target = raw_target
            else:
                resolved_target = self._norm_join(part_path, raw_target)

            out.append({
                "Id": rid,
                "Type": rtype,
                "Target": resolved_target,
                "TargetMode": target_mode,
            })
        return out

    # ---------- sheets ----------
    def _sorted_sheet_parts(self, zf: ZipFile) -> List[str]:
        """Return xl/worksheets/sheetN.xml (sorted by N)."""
        sheet_parts = [
            n for n in zf.namelist()
            if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")
        ]

        def sheet_no(p: str) -> int:
            m = re.search(r"sheet(\d+)\.xml$", p)
            return int(m.group(1)) if m else 10**9

        sheet_parts.sort(key=sheet_no)
        return sheet_parts

    def _sheet_index(self, part_path: str) -> Optional[int]:
        m = re.search(r"sheet(\d+)\.xml$", part_path)
        return int(m.group(1)) if m else None

    def _sheet_name_map(self, zf: ZipFile) -> Dict[str, str]:
        """
        Build mapping from 'xl/worksheets/sheetN.xml' -> user-visible sheet name
        using workbook.xml and workbook.xml.rels.
        """
        mapping: Dict[str, str] = {}
        wb = "xl/workbook.xml"
        if wb not in zf.namelist():
            return mapping

        root = self._read_xml(zf, wb)
        wb_rels = self._read_rels(zf, wb)

        # map rId -> target part
        rid_to_target = {r["Id"]: r["Target"] for r in wb_rels}

        for sheet in root.findall(".//ws:sheets/ws:sheet", NS):
            name = sheet.get("name")
            rid = sheet.get(f"{{{NS['r']}}}id")
            if not rid:
                continue
            target = rid_to_target.get(rid)
            if not target:
                continue
            # normalize to 'xl/worksheets/sheetN.xml'
            part = self._norm_join(wb, target)
            mapping[part] = name or posixpath.basename(part)
        return mapping
