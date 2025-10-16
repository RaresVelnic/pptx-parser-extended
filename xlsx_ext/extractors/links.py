"""
xlsx_ext.extractors.links — find & check links in XLSX (ordered by cell)
"""

from __future__ import annotations

import io
import re
import zipfile
import logging
from typing import List, Dict, Optional, Tuple
from lxml import etree

import requests

from .base import XlsxBaseExtractor, NS, REL_TYPES

logger = logging.getLogger(__name__)

# ---------- small helpers ----------

_CELL_RE = re.compile(r"\$?([A-Za-z]+)\$?(\d+)")

def _col_letters_to_num(col: str) -> int:
    """Convert column letters (A, Z, AA, AB, ...) to 1-based number."""
    col = col.upper()
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n

def _a1_first_to_rc(a1: str) -> Tuple[int, int]:
    """
    Parse first A1 token from a ref like 'E2' or 'E2:E10' -> (row, col).
    Returns very large values for non-parsable refs so they sort last.
    """
    if not a1:
        return (10**9, 10**9)
    token = a1.split(":", 1)[0]  # take top-left of a range
    m = _CELL_RE.fullmatch(token.strip())
    if not m:
        return (10**9, 10**9)
    col_letters, row_str = m.groups()
    return (int(row_str), _col_letters_to_num(col_letters))

def _http_code_meaning(code: Optional[int]) -> str:
    if code is None:
        return "No response"
    try:
        code = int(code)
    except Exception:
        return "Unknown"
    if 200 <= code < 300:
        return "OK"
    if 300 <= code < 400:
        return "Redirect"
    if 400 <= code < 500:
        return "Client Error"
    if 500 <= code < 600:
        return "Server Error"
    return "Other"

def _check_url(url: str) -> tuple[Optional[int], str, str]:
    """
    HEAD with fall-back to GET for servers that reject HEAD.
    """
    headers = {
        "User-Agent": "xlsx-link-checker/1.0",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
    }
    try:
        resp = requests.head(url, allow_redirects=True, timeout=6, headers=headers)
        if resp.status_code in (403, 405, 501):
            resp = requests.get(url, allow_redirects=True, timeout=6, stream=True, headers=headers)
            resp.close()
        return resp.status_code, _http_code_meaning(resp.status_code), resp.reason
    except Exception as e:
        return None, "Bad link", str(e)

def _classify_and_check_target(target: str) -> Tuple[str, str, Optional[int], str]:
    """
    Decide link 'type' and perform checks when appropriate.

    Returns: (type, status, code, description_note)
      - type: "External" | "Email"
      - status: "OK" | "Bad link" | "Error" | "Skipped"
      - code: HTTP status code or None
      - description_note: short note or reason (e.g., 'Non-HTTP scheme')
    """
    if not target:
        return ("External", "Error", None, "Empty target")

    t = target.lower()

    # Email links: mailto:
    if t.startswith("mailto:"):
        return ("Email", "OK", None, "Email link")

    # HTTP(S) -> run checker
    if t.startswith("http://") or t.startswith("https://"):
        code, status, desc = _check_url(target)
        return ("External", status, code, desc or "")

    # Any other scheme (tel:, ftp:, file:, etc.) -> don't HTTP check
    return ("External", "Skipped", None, "Non-HTTP scheme")


class XlsxLinkExtractor(XlsxBaseExtractor):
    """Extract + check hyperlinks in .xlsx workbooks (sorted by cell)."""

    def extract(self, xlsx_bytes: bytes) -> List[Dict]:
        results: List[Dict] = []

        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as z:
            sheet_names = self._sheet_name_map(z)
            # workbook-visible order
            sheet_parts = self._sheet_parts_in_workbook_order(z)
            all_files = set(z.namelist())

            for sp in sheet_parts:
                per_sheet: List[Dict] = []
                idx = self._sheet_index(sp) or 0
                friendly = sheet_names.get(sp, f"Sheet {idx}")
                rels = self._read_rels(z, sp)

                # rId->(target, mode) for hyperlinks on this sheet
                rid_to_link = {
                    r["Id"]: (r["Target"], r.get("TargetMode"))
                    for r in rels
                    if r["Type"] == REL_TYPES["hyperlink"]
                }

                # 1) Worksheet hyperlinks
                try:
                    root = self._read_xml(z, sp)
                except Exception:
                    logger.exception("Failed to parse worksheet %s", sp)
                    root = None

                if root is not None:
                    for h in root.findall(".//ws:hyperlink", NS):
                        cell = h.get("ref") or "(unknown)"
                        rid = h.get(f"{{{NS['r']}}}id")
                        location = h.get("location")

                        if rid and rid in rid_to_link:
                            # External hyperlink via relationship
                            target, mode = rid_to_link[rid]
                            link_type, status, code, note = _classify_and_check_target(target)
                            per_sheet.append({
                                "sheet_index": idx,
                                "sheet": friendly,
                                "where": cell,
                                "type": link_type,        # "External" or "Email"
                                "link": target,
                                "status": status,
                                "code": code if code is not None else "",
                                "description": note,
                            })
                        elif location:
                            # Internal cell reference (e.g., "Sheet2!A1")
                            target_sheet = location.split("!", 1)[0].strip("'\"")
                            status = "OK" if target_sheet in sheet_names.values() else "Unknown"
                            per_sheet.append({
                                "sheet_index": idx,
                                "sheet": friendly,
                                "where": cell,
                                "type": "Internal",
                                "link": location,
                                "status": status,
                                "code": "",
                                "description": "Internal cell reference",
                            })

                # 2) Drawing hyperlinks (pictures/shapes)
                for rel in rels:
                    if rel["Type"] != REL_TYPES["drawing"]:
                        continue
                    drawing_part = rel["Target"]
                    if drawing_part not in all_files:
                        continue

                    d_rels = self._read_rels(z, drawing_part)
                    d_rid_to_link = {
                        r["Id"]: (r["Target"], r.get("TargetMode"))
                        for r in d_rels
                        if r["Type"] == REL_TYPES["hyperlink"]
                    }

                    try:
                        droot = self._read_xml(z, drawing_part)
                    except Exception:
                        logger.exception("Failed to parse drawing %s", drawing_part)
                        droot = None

                    if droot is not None:
                        for cNvPr in droot.findall(".//xdr:cNvPr", NS):
                            h = cNvPr.find("a:hlinkClick", NS)
                            if h is None:
                                continue
                            rid = h.get(f"{{{NS['r']}}}id")
                            if not rid or rid not in d_rid_to_link:
                                continue
                            target, mode = d_rid_to_link[rid]
                            link_type, status, code, note = _classify_and_check_target(target)
                            per_sheet.append({
                                "sheet_index": idx,
                                "sheet": friendly,
                                "where": "drawing",   # not a cell; will sort after real cells
                                "type": link_type,     # "External" or "Email"
                                "link": target,
                                "status": status,
                                "code": code if code is not None else "",
                                "description": note,
                            })

                # ---- sort per-sheet by A1 (row, then col); drawings go last ----
                def _sort_key(item: Dict) -> Tuple[int, int, str]:
                    where = item.get("where") or ""
                    row, col = _a1_first_to_rc(where)
                    return (row, col, where)

                per_sheet.sort(key=_sort_key)
                results.extend(per_sheet)

        return results
