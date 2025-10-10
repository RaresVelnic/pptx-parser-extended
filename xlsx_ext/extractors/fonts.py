# xlsx_ext/extractors/fonts.py
"""
XlsxFontExtractor — collect font names used in an .xlsx workbook, per sheet.

Sources scanned per sheet:
- Cell styles (styles.xml: cellXfs[xf].@fontId -> fonts[fontId]/name@val)
- Inline rich text inside the cell (<is><r><rPr><rFont val="...">)
- Shared strings referenced by that sheet's cells (sharedStrings.xml runs' <rPr><rFont val="...">)
- Drawing text in drawings attached to the sheet (xl/drawings/drawingN.xml: a:latin/@typeface, a:rFonts@*)

Also resolves theme placeholders (+mj, +mn) against theme1.xml when present.
"""

from __future__ import annotations

import io
import zipfile
import logging
from typing import Dict, List, Optional, Set, Tuple

from lxml import etree

from .base import XlsxBaseExtractor, NS, REL_TYPES

logger = logging.getLogger(__name__)


# ---------------- helpers: theme resolution ----------------

def _read_theme_fonts(zf: zipfile.ZipFile) -> Dict[str, str]:
    """
    Return {'major': 'Typeface', 'minor': 'Typeface'} if theme1.xml exists.
    """
    out = {}
    theme_part = "xl/theme/theme1.xml"
    if theme_part not in zf.namelist():
        return out
    try:
        root = etree.fromstring(zf.read(theme_part))
        latin_major = root.find(".//a:themeElements/a:fontScheme/a:majorFont/a:latin", NS)
        latin_minor = root.find(".//a:themeElements/a:fontScheme/a:minorFont/a:latin", NS)
        if latin_major is not None and latin_major.get("typeface"):
            out["major"] = latin_major.get("typeface")
        if latin_minor is not None and latin_minor.get("typeface"):
            out["minor"] = latin_minor.get("typeface")
    except Exception:
        logger.exception("Failed to read theme fonts from %s", theme_part)
    return out


def _resolve_theme_face(face: str, theme_fonts: Dict[str, str]) -> str:
    """
    If face looks like a theme placeholder (+mj or +mn), map to theme latin typeface.
    Otherwise return as-is.
    """
    if not face:
        return face
    f = face.strip()
    # common tokens like "+mj-lt", "+mn-lt", "+mj-ea" etc.
    if f.startswith("+mj") and "major" in theme_fonts:
        return theme_fonts["major"]
    if f.startswith("+mn") and "minor" in theme_fonts:
        return theme_fonts["minor"]
    return f


# ---------------- helpers: styles / sharedStrings ----------------

def _read_styles(zf: zipfile.ZipFile) -> Tuple[Dict[int, str], Dict[int, int]]:
    """
    Parse styles.xml:
      - fonts_map: {fontId -> name}
      - xf_to_font: {xf_idx -> fontId}
    """
    fonts_map: Dict[int, str] = {}
    xf_to_font: Dict[int, int] = {}
    part = "xl/styles.xml"
    if part not in zf.namelist():
        return fonts_map, xf_to_font

    try:
        root = etree.fromstring(zf.read(part))
    except Exception:
        logger.exception("Failed to read %s", part)
        return fonts_map, xf_to_font

    # fonts
    for i, font in enumerate(root.findall(".//ws:fonts/ws:font", NS)):
        # prefer <name val="...">
        name_el = font.find("ws:name", NS)
        if name_el is not None and name_el.get("val"):
            fonts_map[i] = name_el.get("val")
            continue
        # fallback: sometimes drawing font tags exist; ignore for styles
        fonts_map[i] = "Unknown"

    # cellXfs -> fontId
    for i, xf in enumerate(root.findall(".//ws:cellXfs/ws:xf", NS)):
        try:
            fid = int(xf.get("fontId", "0"))
        except Exception:
            fid = 0
        xf_to_font[i] = fid

    return fonts_map, xf_to_font


def _read_shared_strings_fonts(zf: zipfile.ZipFile) -> Dict[int, Set[str]]:
    """
    Build map: {si_index -> set(font names used in that shared string item)}.
    Looks at <si><r><rPr><rFont val="...">.
    """
    out: Dict[int, Set[str]] = {}
    part = "xl/sharedStrings.xml"
    if part not in zf.namelist():
        return out
    try:
        root = etree.fromstring(zf.read(part))
    except Exception:
        logger.exception("Failed to read %s", part)
        return out

    for i, si in enumerate(root.findall(".//ws:si", NS)):
        fonts: Set[str] = set()
        for rpr in si.findall(".//ws:r/ws:rPr", NS):
            rfont = rpr.find("ws:rFont", NS)
            if rfont is not None:
                val = (rfont.get("val") or "").strip()
                if val:
                    fonts.add(val)
        if fonts:
            out[i] = fonts
    return out


# ---------------- helpers: drawings ----------------

def _fonts_in_drawing(root: etree._Element) -> Set[str]:
    """
    From a drawing part, collect font typefaces in text properties:
      - a:latin/@typeface
      - a:rFonts/@ascii|@hAnsi|@ea|@cs
    """
    faces: Set[str] = set()
    # a:latin typefaces
    for el in root.findall(".//a:latin", NS):
        val = (el.get("typeface") or "").strip()
        if val:
            faces.add(val)
    # a:rFonts attributes
    for el in root.findall(".//a:rFonts", NS):
        for attr in ("ascii", "hAnsi", "ea", "cs"):
            v = (el.get(attr) or "").strip()
            if v:
                faces.add(v)
    return faces


# ---------------- main extractor ----------------

class XlsxFontExtractor(XlsxBaseExtractor):
    """
    Returns: List[{"sheet_index": int, "sheet": str, "fonts": [str, ...]}]
    """

    def extract(self, xlsx_bytes: bytes) -> List[Dict]:
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as z:
            theme_fonts = _read_theme_fonts(z)
            fonts_map, xf_to_font = _read_styles(z)
            shared_si_fonts = _read_shared_strings_fonts(z)

            sheet_names = self._sheet_name_map(z)
            sheet_parts = self._sheet_parts_in_workbook_order(z)
            all_files = set(z.namelist())

            results: List[Dict] = []

            for order_idx, sp in enumerate(sheet_parts, start=1):
                idx = self._sheet_index(sp) or order_idx
                friendly = sheet_names.get(sp, f"Sheet {order_idx}")

                used: Set[str] = set()

                # 1) cell styles + inline rich text + shared string fonts used by this sheet
                sheet_root = None
                try:
                    sheet_root = self._read_xml(z, sp)
                except Exception:
                    logger.exception("Failed to parse sheet %s", sp)

                if sheet_root is not None:
                    # a) styles via s="@xf"
                    for c in sheet_root.findall(".//ws:c", NS):
                        # style
                        s_idx = c.get("s")
                        if s_idx is not None:
                            try:
                                xf_idx = int(s_idx)
                                font_id = xf_to_font.get(xf_idx)
                                if font_id is not None:
                                    fname = fonts_map.get(font_id)
                                    if fname:
                                        used.add(_resolve_theme_face(fname, theme_fonts))
                            except Exception:
                                pass

                        # b) inline rich text font (<is><r><rPr><rFont val="...">)
                        for rfont in c.findall(".//ws:is/ws:r/ws:rPr/ws:rFont", NS):
                            val = (rfont.get("val") or "").strip()
                            if val:
                                used.add(_resolve_theme_face(val, theme_fonts))

                        # c) sharedStrings usage in this sheet
                        #    if t="s", <v>index</v> points to sharedStrings.xml[si]
                        t = (c.get("t") or "").strip()
                        if t == "s":
                            v_el = c.find("ws:v", NS)
                            if v_el is not None and v_el.text and v_el.text.strip().isdigit():
                                si_idx = int(v_el.text.strip())
                                for fname in shared_si_fonts.get(si_idx, set()):
                                    used.add(_resolve_theme_face(fname, theme_fonts))

                # 2) drawing fonts attached to this sheet
                sheet_rels = self._read_rels(z, sp)
                drawing_targets = [
                    r["Target"] for r in sheet_rels
                    if r["Type"] == REL_TYPES["drawing"]
                ]
                for drawing_part in drawing_targets:
                    if drawing_part not in all_files:
                        continue
                    try:
                        droot = self._read_xml(z, drawing_part)
                    except Exception:
                        logger.exception("Failed to parse drawing %s", drawing_part)
                        continue
                    for face in _fonts_in_drawing(droot):
                        used.add(_resolve_theme_face(face, theme_fonts))

                # finalize
                if used:
                    results.append({
                        "sheet_index": idx,
                        "sheet": friendly,
                        "fonts": sorted(used, key=lambda s: s.lower()),
                    })

            return results
