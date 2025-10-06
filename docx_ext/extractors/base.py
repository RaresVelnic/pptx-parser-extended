# docx_ext/extractors/base.py
"""
Common helpers for DOCX extractors.

- Works directly on the .docx ZIP (OOXML WordprocessingML)
- Provides XML read, relationships read, part enumeration
"""

from __future__ import annotations

from io import BytesIO
from zipfile import ZipFile
import posixpath
from lxml import etree
from typing import Dict, List, Optional

# Namespaces
NS = {
    "w":  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "v":  "urn:schemas-microsoft-com:vml",
    "o":  "urn:schemas-microsoft-com:office:office",
}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_TYPES = {
    "hyperlink": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    "theme":     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
}

class DocxBaseExtractor:
    """Base with handy ZIP/XML utilities."""

    # ---------- ZIP helpers ----------
    def _zip(self, data: bytes) -> ZipFile:
        return ZipFile(BytesIO(data))

    def _exists(self, zf: ZipFile, path: str) -> bool:
        return path in zf.namelist()

    # ---------- Path helpers ----------
    def _rels_path(self, part_path: str) -> str:
        d, b = posixpath.split(part_path)
        return posixpath.join(d, "_rels", b + ".rels")

    def _norm_join(self, base_part: str, target: str) -> str:
        base_dir = posixpath.dirname(base_part)
        p = posixpath.normpath(posixpath.join(base_dir, target))
        return p.lstrip("/")

    # ---------- XML helpers ----------
    def _read_xml(self, zf: ZipFile, path: str):
        with zf.open(path) as f:
            return etree.fromstring(f.read())

    def _read_rels(self, zf: ZipFile, part_path: str) -> List[Dict[str, str]]:
        rp = self._rels_path(part_path)
        if not self._exists(zf, rp):
            return []
        rels = self._read_xml(zf, rp)
        out = []
        for rel in rels.findall(f".//{{{REL_NS}}}Relationship"):
            out.append({
                "Id": rel.get("Id"),
                "Type": rel.get("Type"),
                "Target": self._norm_join(part_path, rel.get("Target")),
                "TargetMode": rel.get("TargetMode"),  # External or None
            })
        return out

    # ---------- Part lists ----------
    def _doc_part(self) -> str:
        return "word/document.xml"

    def _header_parts(self, zf: ZipFile) -> List[str]:
        return sorted([n for n in zf.namelist() if n.startswith("word/header") and n.endswith(".xml")])

    def _footer_parts(self, zf: ZipFile) -> List[str]:
        return sorted([n for n in zf.namelist() if n.startswith("word/footer") and n.endswith(".xml")])

    def _footnotes_part(self, zf: ZipFile) -> Optional[str]:
        p = "word/footnotes.xml"
        return p if self._exists(zf, p) else None

    def _endnotes_part(self, zf: ZipFile) -> Optional[str]:
        p = "word/endnotes.xml"
        return p if self._exists(zf, p) else None

    # ---------- Theme ----------
    def _theme_part(self, zf: ZipFile) -> Optional[str]:
        """Find theme via document.xml.rels."""
        doc = self._doc_part()
        rels = self._read_rels(zf, doc)
        return next((r["Target"] for r in rels if r["Type"] == REL_TYPES["theme"]), None)
