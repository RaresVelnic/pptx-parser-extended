# xlsx_ext/extractors/links.py
from __future__ import annotations

import io
import zipfile
import logging
from typing import List, Dict, Optional, Tuple

import requests
from lxml import etree

from .base import XlsxBaseExtractor, NS, REL_TYPES

logger = logging.getLogger(__name__)

def _http_code_meaning(code: Optional[int]) -> str:
    if code is None:
        return "No response"
    try:
        c = int(code)
    except Exception:
        return "Unknown"
    if 200 <= c < 300:
        return "OK"
    if 300 <= c < 400:
        return "Redirect"
    if 400 <= c < 500:
        return "Client Error"
    if 500 <= c < 600:
        return "Server Error"
    return "Other"

def _check_url(url: str) -> Tuple[Optional[int], str, str]:
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

class XlsxLinkExtractor(XlsxBaseExtractor):
    """Extract + check hyperlinks in .xlsx workbooks (cells, drawings, externalLinks)."""

    def extract(self, xlsx_bytes: bytes) -> List[Dict]:
        results: List[Dict] = []

        with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as z:
            sheet_name_by_part = self._sheet_name_map(z)
            sheet_parts = self._sorted_sheet_parts(z)
            all_files = set(z.namelist())
            sheet_name_set = set(sheet_name_by_part.values())

            # 1) Worksheet hyperlinks (<ws:hyperlink>)
            for sp in sheet_parts:
                idx = self._sheet_index(sp) or 0
                friendly = sheet_name_by_part.get(sp, f"Sheet {idx}")
                rels = self._read_rels(z, sp)
                rid_to_link = {
                    r["Id"]: (r["Target"], r.get("TargetMode"))
                    for r in rels
                    if r["Type"] == REL_TYPES["hyperlink"]
                }

                try:
                    root = self._read_xml(z, sp)
                except Exception:
                    logger.exception("Failed to parse worksheet %s", sp)
                    continue

                for h in root.findall(".//ws:hyperlink", NS):
                    cell = h.get("ref") or "(unknown)"
                    rid = h.get(f"{{{NS['r']}}}id")
                    location = h.get("location")

                    if rid and rid in rid_to_link:
                        target, _mode = rid_to_link[rid]
                        if target.lower().startswith(("http://", "https://")):
                            code, status, desc = _check_url(target)
                        else:
                            code, status, desc = None, "External", target
                        results.append({
                            "sheet_index": idx,
                            "sheet": friendly,
                            "where": cell,
                            "type": "External",
                            "link": target,
                            "status": status,
                            "code": code,
                            "description": desc,
                        })
                    elif location:
                        # Internal target (Sheet!A1)
                        target_sheet = location.split("!", 1)[0].strip("'\"")
                        status = "OK" if target_sheet in sheet_name_set else "Broken/Missing"
                        results.append({
                            "sheet_index": idx,
                            "sheet": friendly,
                            "where": cell,
                            "type": "Internal",
                            "link": location,
                            "status": status,
                            "code": "",
                            "description": f"Target sheet {'exists' if status=='OK' else 'missing'}: {target_sheet}",
                        })

            # 2) Drawing hyperlinks (pictures/shapes with click actions)
            for sp in sheet_parts:
                idx = self._sheet_index(sp) or 0
                friendly = sheet_name_by_part.get(sp, f"Sheet {idx}")
                sheet_rels = self._read_rels(z, sp)

                for rel in sheet_rels:
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
                        continue

                    for cNvPr in droot.findall(".//xdr:cNvPr", NS):
                        hlink = cNvPr.find("a:hlinkClick", NS)
                        if hlink is None:
                            continue
                        rid = hlink.get(f"{{{NS['r']}}}id")
                        if not rid or rid not in d_rid_to_link:
                            continue
                        target, _mode = d_rid_to_link[rid]
                        if target.lower().startswith(("http://", "https://")):
                            code, status, desc = _check_url(target)
                        else:
                            code, status, desc = None, "External", target
                        where = cNvPr.get("name") or "(drawing)"
                        results.append({
                            "sheet_index": idx,
                            "sheet": friendly,
                            "where": where,
                            "type": "External",
                            "link": target,
                            "status": status,
                            "code": code,
                            "description": desc,
                        })

            # 3) External workbook links (xl/externalLinks/externalLinkN.xml -> .rels extLinkPath)
            wb_part = "xl/workbook.xml"
            wb_rels = self._read_rels(z, wb_part)
            for rel in wb_rels:
                if rel["Type"] != REL_TYPES["extLink"]:
                    continue
                ext_part = rel["Target"]
                if ext_part not in all_files:
                    continue
                ext_rels = self._read_rels(z, ext_part)
                for xr in ext_rels:
                    if xr["Type"] != REL_TYPES["extLinkPath"]:
                        continue
                    target = xr["Target"]
                    if target.lower().startswith(("http://", "https://")):
                        code, status, desc = _check_url(target)
                    else:
                        code, status, desc = None, "External", target
                    results.append({
                        "sheet_index": 0,
                        "sheet": "(workbook)",
                        "where": ext_part.rsplit("/", 1)[-1],
                        "type": "External",
                        "link": target,
                        "status": status,
                        "code": code,
                        "description": "Workbook external link: " + (desc or ""),
                    })

        return results
