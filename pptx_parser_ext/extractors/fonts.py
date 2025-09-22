"""
FontUsageExtractor — OOP wrapper that extracts fonts-per-slide from a PPTX.

- Slide text (runs, paragraph defaults, list-level defaults)
- Master-style fallback (p:txStyles major/minor/other by paragraph level)
- Tables (a:tbl text)
- Charts (ppt/charts/chart*.xml)
- SmartArt / Diagrams (ppt/diagrams/data*.xml)
- Notes pages (ppt/notesSlides/notesSlide*.xml) via Notes Master (with fallback)
- Friendly names:
    * Layout name from <p:sldLayout><p:cSld name="..."> (or @type → humanized)
    * Master name prefers <p:sldMaster><p:cSld @name>, else theme <a:theme @name>, else "Master N"
"""
from __future__ import annotations
from typing import List, Dict, Optional, Set, Any
from zipfile import ZipFile
from io import BytesIO
import re, posixpath
from lxml import etree
from .base import BaseExtractor

NS = {
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p":   "http://schemas.openxmlformats.org/presentationml/2006/main",
    "c":   "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_TYPES = {
    "slideLayout":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
    "slideMaster":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
    "theme":         "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    "chart":         "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    "diagramData":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
    "diagramLayout": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
    "diagramQS":     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
    "diagramColors": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
    "notesSlide":    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
    "notesMaster":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
}

class FontUsageExtractor(BaseExtractor):
    """Extracts per-slide font usage with master/layout names."""
    # ----------- Public API -----------
    def extract(self, pptx_bytes: bytes) -> List[Dict[str, Any]]:
        with ZipFile(BytesIO(pptx_bytes)) as z:
            return self._analyze_fonts_and_masters(z)

    # ----------- Low-level helpers (instance methods to keep OOP) -----------
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

    def _read_rels(self, zf: ZipFile, part_path: str):
        rp = self._rels_path(part_path)
        if rp not in zf.namelist():
            return []
        rels = self._read_xml(zf, rp)
        out = []
        for rel in rels.findall(f".//{{{REL_NS}}}Relationship"):
            out.append({
                "Id": rel.get("Id"),
                "Type": rel.get("Type"),
                "Target": self._norm_join(part_path, rel.get("Target")),
            })
        return out

    # ---- Theme & alias resolution
    def _theme_map(self, zf: ZipFile, theme_part: Optional[str]):
        default = {
            "major": {"latin": None, "ea": None, "cs": None},
            "minor": {"latin": None, "ea": None, "cs": None},
        }
        if not theme_part or theme_part not in zf.namelist():
            return default
        theme = self._read_xml(zf, theme_part)
        fs = theme.find(".//a:themeElements/a:fontScheme", NS)
        if fs is None:
            return default

        def pick(bucket_tag):
            b = {"latin": None, "ea": None, "cs": None}
            node = fs.find(f"a:{bucket_tag}", NS)
            if node is not None:
                for tag in ("latin", "ea", "cs"):
                    el = node.find(f"a:{tag}", NS)
                    if el is not None:
                        b[tag] = el.get("typeface")
            return b

        return {"major": pick("majorFont"), "minor": pick("minorFont")}

    def _resolve_alias(self, face: Optional[str], theme_map: dict, placeholder_style: Optional[str]):
        if not face:
            return None
        if not face.startswith("+"):
            return face
        bucket = theme_map["major"] if placeholder_style == "titleStyle" else theme_map["minor"]
        if face.endswith("-lt"):
            return bucket.get("latin") or bucket.get("ea") or bucket.get("cs")
        if face.endswith("-ea"):
            return bucket.get("ea") or bucket.get("latin") or bucket.get("cs")
        if face.endswith("-cs"):
            return bucket.get("cs") or bucket.get("latin") or bucket.get("ea")
        return None

    # ---- Slide/master helpers
    def _placeholder_type(self, shape):
        ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
        return ph.get("type") if ph is not None and ph.get("type") else None

    def _style_bucket_for_placeholder(self, ph_type: Optional[str]) -> str:
        if ph_type in ("title", "ctrTitle", "subTitle"):
            return "titleStyle"
        if ph_type in ("ftr", "hdr", "dt", "sldNum"):
            return "otherStyle"
        return "bodyStyle"

    def _first_font_from_rPr(self, rPr):
        if rPr is None:
            return None
        for tag in ("latin", "ea", "cs"):
            el = rPr.find(f"a:{tag}", NS)
            if el is not None and el.get("typeface"):
                return el.get("typeface")
        return None

    def _collect_fonts_from_root(self, root, theme_map, bucket: str) -> Set[str]:
        fonts: Set[str] = set()
        for rPr in root.findall(".//a:rPr", NS):
            face = self._first_font_from_rPr(rPr)
            face = self._resolve_alias(face, theme_map, bucket)
            if face:
                fonts.add(face)
        for defRPr in root.findall(".//a:pPr/a:defRPr", NS):
            face = self._first_font_from_rPr(defRPr)
            face = self._resolve_alias(face, theme_map, bucket)
            if face:
                fonts.add(face)
        return fonts

    def _get_tx_styles(self, zf: ZipFile, master_part: Optional[str]):
        if not master_part:
            return None
        root = self._read_xml(zf, master_part)
        return root.find("p:txStyles", NS)

    def _defRPr_from_styles(self, txStyles, style_bucket: str, level: int):
        if txStyles is None:
            return None
        level = max(1, min(9, level))
        return txStyles.find(f"p:{style_bucket}/a:lvl{level}pPr/a:defRPr", NS)

    def _para_level(self, p) -> int:
        pPr = p.find("a:pPr", NS)
        if pPr is None or pPr.get("lvl") is None:
            return 1
        try:
            return int(pPr.get("lvl")) + 1
        except Exception:
            return 1

    # ---- Charts / SmartArt / Notes
    def _scan_chart_fonts(self, zf: ZipFile, chart_part: str, theme_map):
        root = self._read_xml(zf, chart_part)
        return self._collect_fonts_from_root(root, theme_map, bucket="bodyStyle")

    def _scan_diagram_fonts(self, zf: ZipFile, data_part: str, theme_map):
        root = self._read_xml(zf, data_part)
        return self._collect_fonts_from_root(root, theme_map, bucket="bodyStyle")

    def _scan_notes_fonts(self, zf: ZipFile, notes_part: str,
                          notes_master: Optional[str], fallback_theme_map: Optional[dict],
                          fallback_txStyles: Any):
        fonts = set()
        if notes_master:
            nrels = self._read_rels(zf, notes_master)
            theme_part = next((r["Target"] for r in nrels if r["Type"] == REL_TYPES["theme"]), None)
            theme_map = self._theme_map(zf, theme_part)
            txStyles = self._get_tx_styles(zf, notes_master)
        else:
            theme_map = fallback_theme_map or {"major": {"latin": None, "ea": None, "cs": None},
                                               "minor": {"latin": None, "ea": None, "cs": None}}
            txStyles = fallback_txStyles

        notes = self._read_xml(zf, notes_part)
        for shape in notes.findall(".//p:sp", NS):
            ph_type = self._placeholder_type(shape)
            bucket = self._style_bucket_for_placeholder(ph_type)
            txBody = shape.find("p:txBody", NS)
            if txBody is None:
                continue
            for p in txBody.findall("a:p", NS):
                lvl = self._para_level(p)
                for rPr in p.findall("a:r/a:rPr", NS):
                    face = self._first_font_from_rPr(rPr)
                    face = self._resolve_alias(face, theme_map, bucket)
                    if face:
                        fonts.add(face)
                defRPr = p.find("a:pPr/a:defRPr", NS)
                face = self._first_font_from_rPr(defRPr)
                face = self._resolve_alias(face, theme_map, bucket)
                if face:
                    fonts.add(face)
                for defRPr in txBody.findall(f"a:lstStyle/a:lvl{lvl}pPr/a:defRPr", NS):
                    face = self._first_font_from_rPr(defRPr)
                    face = self._resolve_alias(face, theme_map, bucket)
                    if face:
                        fonts.add(face)
                m_defRPr = self._defRPr_from_styles(txStyles, bucket, lvl) if txStyles is not None else None
                face = self._first_font_from_rPr(m_defRPr)
                face = self._resolve_alias(face, theme_map, bucket)
                if face:
                    fonts.add(face)
        for tbl_tx in notes.findall(".//a:tbl//a:txBody", NS):
            fonts |= self._collect_fonts_from_root(tbl_tx, theme_map, bucket="bodyStyle")
        return fonts

    # ---- Master & layout names
    def _master_id_from_path(self, master_part: Optional[str]) -> Optional[int]:
        if not master_part:
            return None
        m = re.search(r"slideMaster(\d+)\.xml$", master_part)
        return int(m.group(1)) if m else None

    def _master_friendly_name(self, zf: ZipFile, master_part: Optional[str]) -> Optional[str]:
        if not master_part or master_part not in zf.namelist():
            return None
        root = self._read_xml(zf, master_part)
        cSld = root.find("p:cSld", NS)
        if cSld is not None:
            n = cSld.get("name")
            if n:
                return n
        return None

    def _humanize_layout_type(self, t: Optional[str]) -> Optional[str]:
        if not t:
            return None
        MAP = {
            "title": "Title",
            "titleOnly": "Title Only",
            "titleAndContent": "Title and Content",
            "twoContent": "Two Content",
            "sectionHeader": "Section Header",
            "comparison": "Comparison",
            "contentWithCaption": "Content with Caption",
            "pictureWithCaption": "Picture with Caption",
            "blank": "Blank",
        }
        return MAP.get(t, t)

    def _layout_name(self, zf: ZipFile, layout_part: Optional[str]) -> Optional[str]:
        if not layout_part or layout_part not in zf.namelist():
            return None
        root = self._read_xml(zf, layout_part)
        cSld = root.find("p:cSld", NS)
        if cSld is not None:
            n = cSld.get("name")
            if n:
                return n
        for attr in ("name", "matchingName"):
            n = root.get(attr)
            if n:
                return n
        t = self._humanize_layout_type(root.get("type"))
        if t:
            return t
        return posixpath.basename(layout_part)

    def _theme_name(self, zf: ZipFile, theme_part: Optional[str]) -> Optional[str]:
        if not theme_part or theme_part not in zf.namelist():
            return None
        root = self._read_xml(zf, theme_part)  # <a:theme name="...">
        return root.get("name")

    # ----------- Core analyzer -----------
    def _analyze_fonts_and_masters(self, z: ZipFile) -> List[Dict[str, Any]]:
        slide_parts = [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]

        def slide_no(p):
            m = re.search(r"slide(\d+)\.xml$", p)
            return int(m.group(1)) if m else 10**9

        slide_parts.sort(key=slide_no)

        theme_cache: dict[str, dict] = {}
        styles_cache: dict[str, object] = {}
        rows: List[Dict[str, Any]] = []

        for sp in slide_parts:
            s_num = slide_no(sp)
            rels = self._read_rels(z, sp)

            layout = next((r["Target"] for r in rels if r["Type"] == REL_TYPES["slideLayout"]), None)
            direct_master = next((r["Target"] for r in rels if r["Type"] == REL_TYPES["slideMaster"]), None)

            master = None
            if layout:
                lrels = self._read_rels(z, layout)
                master = next((r["Target"] for r in lrels if r["Type"] == REL_TYPES["slideMaster"]), None)
            if not master and direct_master:
                master = direct_master

            theme_part = None
            if master:
                mrels = self._read_rels(z, master)
                theme_part = next((r["Target"] for r in mrels if r["Type"] == REL_TYPES["theme"]), None)

            if master and master not in theme_cache:
                theme_cache[master] = self._theme_map(z, theme_part)
            theme_map = theme_cache.get(
                master,
                {"major": {"latin": None, "ea": None, "cs": None},
                 "minor": {"latin": None, "ea": None, "cs": None}}
            )

            if master and master not in styles_cache:
                styles_cache[master] = self._get_tx_styles(z, master)
            txStyles = styles_cache.get(master)

            fonts_here: Set[str] = set()
            slide = self._read_xml(z, sp)

            for shape in slide.findall(".//p:sp", NS):
                ph_type = self._placeholder_type(shape)
                bucket = self._style_bucket_for_placeholder(ph_type)
                txBody = shape.find("p:txBody", NS)
                if txBody is None:
                    continue
                for p in txBody.findall("a:p", NS):
                    lvl = self._para_level(p)
                    for rPr in p.findall("a:r/a:rPr", NS):
                        face = self._first_font_from_rPr(rPr)
                        face = self._resolve_alias(face, theme_map, bucket)
                        if face:
                            fonts_here.add(face)
                    defRPr = p.find("a:pPr/a:defRPr", NS)
                    face = self._first_font_from_rPr(defRPr)
                    face = self._resolve_alias(face, theme_map, bucket)
                    if face:
                        fonts_here.add(face)
                    for defRPr in txBody.findall(f"a:lstStyle/a:lvl{lvl}pPr/a:defRPr", NS):
                        face = self._first_font_from_rPr(defRPr)
                        face = self._resolve_alias(face, theme_map, bucket)
                        if face:
                            fonts_here.add(face)
                    m_defRPr = self._defRPr_from_styles(txStyles, bucket, lvl)
                    face = self._first_font_from_rPr(m_defRPr)
                    face = self._resolve_alias(face, theme_map, bucket)
                    if face:
                        fonts_here.add(face)

            for tbl_tx in slide.findall(".//a:tbl//a:txBody", NS):
                fonts_here |= self._collect_fonts_from_root(tbl_tx, theme_map, bucket="bodyStyle")

            for rel in rels:
                if rel["Type"] == REL_TYPES["chart"] and rel["Target"] in z.namelist():
                    try:
                        fonts_here |= self._scan_chart_fonts(z, rel["Target"], theme_map)
                    except Exception:
                        pass

            for rel in rels:
                if rel["Type"] == REL_TYPES["diagramData"] and rel["Target"] in z.namelist():
                    try:
                        fonts_here |= self._scan_diagram_fonts(z, rel["Target"], theme_map)
                    except Exception:
                        pass

            notes_part = next((r["Target"] for r in rels if r["Type"] == REL_TYPES["notesSlide"]), None)
            if notes_part:
                nrels = self._read_rels(z, notes_part)
                notes_master = next((r["Target"] for r in nrels if r["Type"] == REL_TYPES["notesMaster"]), None)
                try:
                    fonts_here |= self._scan_notes_fonts(
                        z, notes_part, notes_master,
                        fallback_theme_map=theme_map, fallback_txStyles=txStyles
                    )
                except Exception:
                    pass

            mid = self._master_id_from_path(master)
            layout_label = self._layout_name(z, layout) or (posixpath.basename(layout) if layout else None)
            mfriendly = self._master_friendly_name(z, master)
            tname = self._theme_name(z, theme_part)
            if mfriendly:
                master_label = mfriendly
            elif tname:
                master_label = tname
            elif mid is not None:
                master_label = f"Master {mid}"
            else:
                master_label = "—"

            rows.append({
                "slide": s_num,
                "master_id": mid,
                "master": master_label,
                "layout": layout_label or "—",
                "fonts": sorted(fonts_here),
            })

        rows.sort(key=lambda r: r["slide"])
        return rows
