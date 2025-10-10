# xlsx_ext/extractors/descriptions.py
"""
XlsxDescriptionExtractor — collect Alt Text (descriptions) from XLSX drawings.

Primary path:
- For each worksheet, follow its drawing rel(s) to xl/drawings/drawingN.xml.
- In those drawings, read <xdr:cNvPr descr="..."> (non-empty only).
- Try to attach the top-left anchor cell (A1 notation) for context.

Additional mapping (so results look like Excel placement):
- If a cell uses IMAGE(...) (any locale), take the second quoted arg as alt and
  attach to that cell (e.g., "E3 — ...").
- If richData parts expose ids (rvId) and the sheet references them, map the
  richData alt to the nearest <c r="..."> cell.

Fallback (if nothing found per-sheet or via mappings):
- Look into xl/richData/rdrichvalue.xml and collect any text nodes.
  Return them under a synthetic "sheet" named "Rich Data (workbook)".

Output schema per sheet:
[
  {"sheet_index": 1, "sheet": "Sheet 1", "descriptions": ["E3 — Figure caption", "B7 — Logo"]},
  ...
]
"""

from __future__ import annotations

import io
import re
import zipfile
import logging
from typing import Dict, List, Optional, Tuple, Set

from lxml import etree

from .base import XlsxBaseExtractor, NS, REL_TYPES

logger = logging.getLogger(__name__)

# --------- small helpers: columns / A1 / anchors ----------

def _col_letters(zero_based_col: int) -> str:
    n = zero_based_col + 1
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters

def _anchor_pos(anchor_el: etree._Element) -> Tuple[Optional[int], Optional[int]]:
    fr = anchor_el.find("xdr:from", NS)
    if fr is None:
        return (None, None)
    try:
        col = int(fr.findtext("xdr:col", default="0", namespaces=NS))
        row = int(fr.findtext("xdr:row", default="0", namespaces=NS))
        return (row, col)
    except Exception:
        return (None, None)

def _anchor_cell(anchor_el: etree._Element) -> Optional[str]:
    row, col = _anchor_pos(anchor_el)
    if row is None or col is None:
        return None
    return f"{_col_letters(col)}{row + 1}"

_CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")

def _a1_from_cell_elem(c_el: etree._Element) -> Optional[str]:
    if c_el is None:
        return None
    r = c_el.get("r")
    if r and _CELL_RE.match(r):
        return r
    return None

def _rowcol_from_a1(a1: str) -> Optional[Tuple[int, int]]:
    m = _CELL_RE.match(a1 or "")
    if not m:
        return None
    col_letters, row_str = m.groups()
    col = 0
    for ch in col_letters:
        col = col * 26 + (ord(ch) - 64)  # A->1
    col0 = col - 1
    try:
        row0 = int(row_str) - 1
        return (row0, col0)
    except Exception:
        return None

# --------- IMAGE() formula parsing (locale-agnostic) ----------

_QUOTED_STR_RE = re.compile(r'"((?:[^"]|"")*)"')

def _parse_image_formula_alt_text(formula: str) -> Optional[str]:
    """
    Works for any localized IMAGE function name.
    We collect the quoted strings; if there are >=2, take the second as 'alt'.
    """
    if "(" not in (formula or ""):
        return None
    quoted = _QUOTED_STR_RE.findall(formula or "")
    if len(quoted) >= 2:
        alt = quoted[1].replace('""', '"').strip()
        return alt or None
    return None

# --------- richData helpers (ids + text + sheet references) ----------

def _closest_cell_ancestor(node: etree._Element) -> Optional[etree._Element]:
    cur = node
    while cur is not None:
        # tag endswith '}c' is <ws:c> in any ns
        if isinstance(cur.tag, str) and cur.tag.endswith("}c"):
            return cur
        cur = cur.getparent()
    return None

def _iter_nodes_with_rvid(root: etree._Element) -> List[Tuple[str, etree._Element]]:
    results: List[Tuple[str, etree._Element]] = []
    for el in root.iter():
        for attr_name, attr_val in el.items():
            if (attr_name.endswith("}rvId") or attr_name.lower().endswith("rvid")) and attr_val:
                results.append((attr_val, el))
    return results

def _collect_richdata_items(zf: zipfile.ZipFile) -> Dict[str, str]:
    """
    Try to build {rvId -> text} from any xl/richData/*.xml.
    We grab all text content and also inspect attributes (JSON-ish) if needed.
    If no obvious id found, skip (ids are required for mapping).
    """
    out: Dict[str, str] = {}
    for name in zf.namelist():
        if not name.startswith("xl/richData/") or not name.endswith(".xml"):
            continue
        try:
            root = etree.fromstring(zf.read(name))
        except Exception:
            continue

        # Gather human text
        text_bits = [t for t in root.xpath(".//text()") if isinstance(t, str) and t.strip()]
        text = " ".join(t.strip() for t in text_bits).strip()
        if not text:
            # as a last resort, try attribute values (sometimes JSON-ish)
            for _, val in root.items():
                if val and isinstance(val, str) and val.strip():
                    text = val.strip()
                    break
        if not text:
            continue

        # Try find an id attribute that looks like a rich value id
        rv_id = None
        for attr_name, attr_val in root.items():
            if attr_val and (attr_name.endswith("}rvId") or attr_name.lower().endswith("rvid")):
                rv_id = attr_val.strip()
                break

        if rv_id:
            out[rv_id] = text
    return out

# --------- main extractor ----------

class XlsxDescriptionExtractor(XlsxBaseExtractor):
    """Extract per-sheet Alt Text from drawings; add IMAGE()/richData→cell mapping; fallback to workbook-level richData."""

    # ---------- Fallback: richData (workbook) ----------
    def _fallback_richdata_texts(self, z: zipfile.ZipFile) -> List[str]:
        """
        Collect human-meaningful text from xl/richData/rdrichvalue.xml (if present).
        Prefer DrawingML runs (<a:t>), else all text nodes. Light filtering only.
        """
        part = "xl/richData/rdrichvalue.xml"
        if part not in z.namelist():
            return []

        try:
            root = self._read_xml(z, part)
        except Exception:
            logger.exception("Failed to parse richData part %s", part)
            return []

        def is_meaningful(s: str) -> bool:
            s = s.strip()
            if not s:
                return False
            # keep typical human text (letters), plus URLs/emails
            if any(ch.isalpha() for ch in s):
                return True
            lowered = s.lower()
            if "http://" in lowered or "https://" in lowered or "www." in lowered or "@" in s:
                return True
            return False

        try:
            a_t_nodes = root.findall(".//a:t", NS)
            candidates = [ (n.text or "").strip() for n in a_t_nodes if (n.text or "").strip() ]
            if not candidates:
                candidates = [
                    t.strip() for t in root.xpath(".//text()")
                    if isinstance(t, str) and t.strip()
                ]

            seen = set()
            out: List[str] = []
            for s in candidates:
                if not is_meaningful(s):
                    continue
                if s in seen:
                    continue
                seen.add(s)
                out.append(s)
            return out
        except Exception:
            logger.exception("Error extracting text from %s", part)
            return []

    # ---------- Primary: DrawingML + mappings ----------
    def extract(self, xlsx_bytes: bytes) -> List[Dict]:
        out: List[Dict] = []
        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as z:
            sheet_names = self._sheet_name_map(z)
            sheet_parts = self._sheet_parts_in_workbook_order(z)
            all_files = set(z.namelist())

            # Pre-scan richData (ids -> text) for later sheet mapping
            rd_id_to_text = _collect_richdata_items(z)

            # Iterate in workbook (tab) order
            for order_idx, sp in enumerate(sheet_parts, start=1):
                idx = self._sheet_index(sp) or order_idx
                friendly = sheet_names.get(sp, f"Sheet {order_idx}")

                # Collect raw entries as (row_sort, col_sort, label)
                entries: List[Tuple[int, int, str]] = []

                # ---- 1) Drawing-based (xdr) ----
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

                    anchors = (
                        root.findall(".//xdr:twoCellAnchor", NS) +
                        root.findall(".//xdr:oneCellAnchor", NS) +
                        root.findall(".//xdr:absoluteAnchor", NS)
                    )

                    for anchor in anchors:
                        row, col = _anchor_pos(anchor)
                        where = _anchor_cell(anchor)

                        for cNvPr in anchor.findall(".//xdr:cNvPr", NS):
                            descr = (cNvPr.get("descr") or "").strip()
                            if not descr:
                                continue
                            label = f"{where} — {descr}" if where else descr
                            r_sort = row if row is not None else 10**9
                            c_sort = col if col is not None else 10**9
                            entries.append((r_sort, c_sort, label))

                # ---- 2) IMAGE() formula-based alts on this sheet ----
                sheet_root = None
                try:
                    sheet_root = self._read_xml(z, sp)
                except Exception:
                    logger.exception("Failed to parse sheet %s", sp)

                if sheet_root is not None:
                    for c in sheet_root.findall(".//ws:c", NS):
                        a1 = _a1_from_cell_elem(c)
                        if not a1:
                            continue
                        f_el = c.find("ws:f", NS)
                        if f_el is None or f_el.text is None:
                            continue
                        alt = _parse_image_formula_alt_text(f_el.text or "")
                        if not alt:
                            continue
                        rc = _rowcol_from_a1(a1) or (10**9, 10**9)
                        label = f"{a1} — {alt}"
                        entries.append((rc[0], rc[1], label))

                    # ---- 3) Map richData ids referenced in this sheet to cells ----
                    if rd_id_to_text:
                        for rvId, node in _iter_nodes_with_rvid(sheet_root):
                            text = rd_id_to_text.get(rvId)
                            if not text:
                                continue
                            host_cell = _closest_cell_ancestor(node)
                            a1 = _a1_from_cell_elem(host_cell) if host_cell is not None else None
                            if not a1:
                                continue
                            rc = _rowcol_from_a1(a1) or (10**9, 10**9)
                            label = f"{a1} — {text.strip()}"
                            entries.append((rc[0], rc[1], label))

                # Sort by row, then column so the list follows sheet reading order
                entries.sort(key=lambda t: (t[0], t[1]))

                # De-duplicate while preserving order
                seen: Set[str] = set()
                descs_for_sheet: List[str] = []
                for _, __, label in entries:
                    if label in seen:
                        continue
                    seen.add(label)
                    descs_for_sheet.append(label)

                if descs_for_sheet:
                    out.append({
                        "sheet_index": idx,
                        "sheet": friendly,
                        "descriptions": descs_for_sheet
                    })

            # ---------- Fallback: if nothing found anywhere, try richData ----------
            if not out:
                rd_texts = self._fallback_richdata_texts(z)
                if rd_texts:
                    logger.info("XLSX descriptions: using richData fallback with %d entries", len(rd_texts))
                    out.append({
                        "sheet_index": 0,
                        "sheet": "Rich Data (workbook)",
                        "descriptions": rd_texts
                    })

        return out
