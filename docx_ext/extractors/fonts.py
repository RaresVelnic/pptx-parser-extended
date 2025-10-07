# docx_ext/extractors/fonts.py
"""
DocxFontExtractor — collect font families used in a DOCX.

Coverage:
- Run-level fonts: <w:rPr><w:rFonts ...> across:
    * word/document.xml
    * word/header*.xml
    * word/footer*.xml
    * word/footnotes.xml / word/endnotes.xml (if present)
- Document defaults & style-level fonts: word/styles.xml
- Theme resolution for @*Theme attributes.

Output:
[
  {"part": "styles/defaults", "fonts": [...]},  # if any found in styles
  {"part": "document", "fonts": [...]},
  {"part": "header1", "fonts": [...]},
  ...
]
"""

from __future__ import annotations

import logging
import posixpath
from typing import Dict, List, Optional, Set

from .base import DocxBaseExtractor, NS

logger = logging.getLogger(__name__)


class DocxFontExtractor(DocxBaseExtractor):
    """Collect per-part font families for DOCX (with styles/defaults surfaced)."""

    # -------- Theme mapping --------
    def _theme_map(self, z, theme_part: Optional[str]) -> Dict[str, Dict[str, Optional[str]]]:
        default = {
            "major": {"latin": None, "ea": None, "cs": None},
            "minor": {"latin": None, "ea": None, "cs": None},
        }
        if not theme_part or not self._exists(z, theme_part):
            return default
        root = self._read_xml(z, theme_part)
        fs = root.find(".//a:themeElements/a:fontScheme", NS)
        if fs is None:
            return default

        def pick(tag: str) -> Dict[str, Optional[str]]:
            bucket = {"latin": None, "ea": None, "cs": None}
            node = fs.find(f"a:{tag}", NS)
            if node is not None:
                lat = node.find("a:latin", NS)
                if lat is not None:
                    bucket["latin"] = lat.get("typeface")
                ea = node.find("a:ea", NS)
                if ea is not None:
                    bucket["ea"] = ea.get("typeface")
                cs = node.find("a:cs", NS)
                if cs is not None:
                    bucket["cs"] = cs.get("typeface")
            return bucket

        return {"major": pick("majorFont"), "minor": pick("minorFont")}

    def _resolve_theme_face(self, theme_map: Dict, theme_token: Optional[str]) -> Optional[str]:
        """Resolve @*Theme tokens via theme_map."""
        if not theme_token:
            return None
        token = theme_token.strip()
        if token.startswith("major"):
            bucket = "major"
        elif token.startswith("minor"):
            bucket = "minor"
        else:
            return None

        if token.endswith("Ascii"):
            key = "latin"
        elif token.endswith("EastAsia"):
            key = "ea"
        elif token.endswith("Bidi"):
            key = "cs"
        else:
            key = "latin"
        return (theme_map.get(bucket) or {}).get(key)

    # -------- Fonts from styles.xml --------
    def _styles_rFonts(self, z) -> Set[str]:
        """Collect fonts from docDefaults and style-level rPr."""
        fonts: Set[str] = set()
        styles_part = "word/styles.xml"
        if not self._exists(z, styles_part):
            return fonts
        root = self._read_xml(z, styles_part)

        # docDefaults / rPrDefault
        for rFonts in root.findall(".//w:docDefaults/w:rPrDefault/w:rPr/w:rFonts", NS):
            fonts |= self._rFonts_to_names(z, rFonts)

        # style-level rPr (paragraph/character styles)
        for rFonts in root.findall(".//w:style/w:rPr/w:rFonts", NS):
            fonts |= self._rFonts_to_names(z, rFonts)

        return fonts

    # -------- Fonts from a part --------
    def _part_rFonts(self, z, part: str) -> Set[str]:
        fonts: Set[str] = set()
        if not self._exists(z, part):
            return fonts
        root = self._read_xml(z, part)
        for rFonts in root.findall(".//w:rPr/w:rFonts", NS):
            fonts |= self._rFonts_to_names(z, rFonts)
        return fonts

    # -------- Convert <w:rFonts ...> to concrete face names --------
    def _rFonts_to_names(self, z, rFonts) -> Set[str]:
        theme_part = self._theme_part(z)
        theme_map = self._theme_map(z, theme_part)
        names: Set[str] = set()

        # Direct names
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            val = rFonts.get(f"{{{NS['w']}}}{attr}")
            if val:
                names.add(val)

        # Theme tokens (correct attribute names, including csTheme)
        token_map = {
            "asciiTheme": "Ascii",
            "hAnsiTheme": "Ascii",
            "eastAsiaTheme": "EastAsia",
            "csTheme": "Bidi",
        }
        for attr, _suffix in token_map.items():
            token = rFonts.get(f"{{{NS['w']}}}{attr}")
            if token:
                face = self._resolve_theme_face(theme_map, token)
                if face:
                    names.add(face)

        return names

    def _label_for_part(self, part: str) -> str:
        if part == "word/document.xml":
            return "document"
        if part.startswith("word/header"):
            base = posixpath.basename(part)  # header1.xml
            n = "".join(ch for ch in base if ch.isdigit())
            return f"header{n or ''}".rstrip()
        if part.startswith("word/footer"):
            base = posixpath.basename(part)
            n = "".join(ch for ch in base if ch.isdigit())
            return f"footer{n or ''}".rstrip()
        if part.endswith("footnotes.xml"):
            return "footnotes"
        if part.endswith("endnotes.xml"):
            return "endnotes"
        return part

    # -------- Public API --------
    def extract(self, docx_bytes: bytes) -> List[Dict]:
        rows: List[Dict] = []
        with self._zip(docx_bytes) as z:
            # Styles/defaults (surface as its own row so users see something even if runs rely on styles)
            styles_fonts = self._styles_rFonts(z)
            if styles_fonts:
                rows.append({"part": "styles/defaults", "fonts": sorted(styles_fonts)})

            # Per-part runs
            parts = [self._doc_part()]
            parts += self._header_parts(z)
            parts += self._footer_parts(z)
            fn = self._footnotes_part(z)
            en = self._endnotes_part(z)
            if fn: parts.append(fn)
            if en: parts.append(en)

            for p in parts:
                fset = self._part_rFonts(z, p)
                if fset:
                    rows.append({"part": self._label_for_part(p), "fonts": sorted(fset)})

        return rows
