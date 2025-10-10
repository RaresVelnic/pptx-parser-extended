"""
xlsx_ext.base — helpers & constants for Excel (.xlsx) parsing
"""

from __future__ import annotations

import re
import posixpath
from typing import Dict, Optional, List
from zipfile import ZipFile
from lxml import etree

# Namespaces used across xlsx parts
NS: Dict[str, str] = {
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "ws":  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    # NEW (for legacy VML drawings e.g., comments/shapes)
    "v":   "urn:schemas-microsoft-com:vml",
    "o":   "urn:schemas-microsoft-com:office:office",
    "x":   "urn:schemas-microsoft-com:office:excel",
}

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

REL_TYPES: Dict[str, str] = {
    "worksheet":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
    "drawing":     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    "hyperlink":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    "chart":       "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    "extLink":     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
    "extLinkPath": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
    # NEW
    "vmlDrawing":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
}


class XlsxBaseExtractor:
    """
    Base class with ZIP/XML helpers for .xlsx:
    - read XML parts
    - read .rels
    - resolve relative targets
    - sheet ordering & friendly sheet names
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
        """
        Read relationships for 'part_path'. For external targets (TargetMode='External'
        or a scheme like 'http://', 'https://', 'mailto:' etc.), DO NOT join with the
        base path—return the target as-is.
        """
        rp = self._rels_path(part_path)
        if rp not in zf.namelist():
            return []
        rels = self._read_xml(zf, rp)
        out: List[Dict[str, str]] = []
        for rel in rels.findall(f".//{{{REL_NS}}}Relationship"):
            rid = rel.get("Id")
            rtype = rel.get("Type")
            raw_target = rel.get("Target") or ""
            tmode = rel.get("TargetMode")
            # Detect external: TargetMode="External" OR obviously absolute/schemed Target
            is_external = (
                (tmode and tmode.lower() == "external")
                or raw_target.startswith(("http://", "https://", "mailto:", "ftp:", "file:", "tel:", "news:"))
                or "://" in raw_target
            )
            target = raw_target if is_external else self._norm_join(part_path, raw_target)
            out.append({
                "Id": rid,
                "Type": rtype,
                "Target": target,
                "TargetMode": tmode,
            })
        return out

    # ---------- sheets ----------
    def _sorted_sheet_parts(self, zf: ZipFile) -> List[str]:
        """Fallback: return xl/worksheets/sheetN.xml (sorted by N)."""
        sheet_parts = [
            n for n in zf.namelist()
            if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")
        ]

        def sheet_no(p: str) -> int:
            m = re.search(r"sheet(\d+)\.xml$", p)
            return int(m.group(1)) if m else 10**9

        sheet_parts.sort(key=sheet_no)
        return sheet_parts

    def _sheet_parts_in_workbook_order(self, zf: ZipFile) -> List[str]:
        """
        Preferred: return sheet parts as ordered in workbook.xml (<sheets><sheet> sequence).
        """
        wb = "xl/workbook.xml"
        if wb not in zf.namelist():
            return self._sorted_sheet_parts(zf)

        root = self._read_xml(zf, wb)
        wb_rels = self._read_rels(zf, wb)
        rid_to_target = {r["Id"]: r["Target"] for r in wb_rels}

        parts: List[str] = []
        for sheet in root.findall(".//ws:sheets/ws:sheet", NS):
            rid = sheet.get(f"{{{NS['r']}}}id")
            if not rid:
                continue
            target = rid_to_target.get(rid)
            if not target:
                continue
            parts.append(self._norm_join(wb, target))
        # Ensure they actually exist; otherwise fall back
        parts = [p for p in parts if p in zf.namelist()]
        return parts or self._sorted_sheet_parts(zf)

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
        rid_to_target = {r["Id"]: r["Target"] for r in wb_rels}

        for sheet in root.findall(".//ws:sheets/ws:sheet", NS):
            name = sheet.get("name")
            rid = sheet.get(f"{{{NS['r']}}}id")
            if not rid:
                continue
            target = rid_to_target.get(rid)
            if not target:
                continue
            part = self._norm_join(wb, target)  # normalize to 'xl/worksheets/sheetN.xml'
            mapping[part] = name or posixpath.basename(part)
        return mapping
