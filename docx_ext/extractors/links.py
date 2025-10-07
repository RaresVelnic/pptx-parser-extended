# docx_ext/extractors/links.py
"""
DOCX link extractor

Finds links in:
- w:hyperlink (external via r:id, internal via @anchor)
- Field codes: w:fldSimple/@w:instr and w:instrText (HYPERLINK "...")
- Drawing parts (best effort): any a:hlinkClick/@r:id mapped via drawing .rels

Returns list[dict]:
{
  "part": "document" | "header1" | "footer1" | "footnotes" | "endnotes",
  "type": "External" | "Internal",
  "link": str,
  "status": str,  # OK / Redirect / Client Error / Server Error / Bad link / No response / External
  "code": int|""|None,
  "description": str
}
"""

from __future__ import annotations

import logging
import re
from typing import Dict, List, Optional, Set, Tuple

import requests
from lxml import etree

from .base import DocxBaseExtractor, NS, REL_TYPES

logger = logging.getLogger(__name__)

_HTTP_SCHEMES = ("http://", "https://")
_EXT_SCHEMES  = _HTTP_SCHEMES + ("mailto:", "ftp://", "tel:", "file://")


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


def _check_url(url: str) -> Tuple[Optional[int], str, str]:
    """HEAD with fallback to GET for servers that reject HEAD."""
    headers = {
        "User-Agent": "docx-link-checker/1.0",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
    }
    try:
        if url.startswith(_HTTP_SCHEMES):
            resp = requests.head(url, allow_redirects=True, timeout=6, headers=headers)
            if resp.status_code in (403, 405, 501):
                resp = requests.get(url, allow_redirects=True, timeout=6, stream=True, headers=headers)
                resp.close()
            return resp.status_code, _http_code_meaning(resp.status_code), resp.reason
        # Non-HTTP schemes: don't fetch; just report as External
        return None, "External", url.split(":", 1)[0].upper()
    except Exception as e:
        return None, "Bad link", str(e)


class DocxLinkExtractor(DocxBaseExtractor):
    """Extract and validate hyperlinks from DOCX."""

    # ---------- helpers ----------
    def _label_for_part(self, part: str) -> str:
        # reuse the one in fonts extractor (same logic)
        if part == "word/document.xml":
            return "document"
        if part.startswith("word/header"):
            base = part.rsplit("/", 1)[-1]  # header1.xml
            n = "".join(ch for ch in base if ch.isdigit())
            return f"header{n or ''}".rstrip()
        if part.startswith("word/footer"):
            base = part.rsplit("/", 1)[-1]
            n = "".join(ch for ch in base if ch.isdigit())
            return f"footer{n or ''}".rstrip()
        if part.endswith("footnotes.xml"):
            return "footnotes"
        if part.endswith("endnotes.xml"):
            return "endnotes"
        return part

    _HYP_RE = re.compile(r'HYPERLINK\s+"([^"]+)"', re.I)

    def _urls_from_instr_text(self, root) -> List[str]:
        # Combine all instrText blocks; Word often splits field code across runs
        texts = [t for t in root.findall(".//w:instrText", NS) if t.text]
        if not texts:
            return []
        blob = " ".join(t.text for t in texts)
        return self._HYP_RE.findall(blob)

    def _urls_from_fldSimple(self, root) -> List[str]:
        out: List[str] = []
        for fs in root.findall(".//w:fldSimple", NS):
            instr = fs.get(f"{{{NS['w']}}}instr") or ""
            out += self._HYP_RE.findall(instr)
        return out

    # ---------- main ----------
    def extract(self, docx_bytes: bytes) -> List[Dict]:
        rows: List[Dict] = []
        with self._zip(docx_bytes) as z:
            parts = [self._doc_part()] + self._header_parts(z) + self._footer_parts(z)
            fn = self._footnotes_part(z)
            en = self._endnotes_part(z)
            if fn: parts.append(fn)
            if en: parts.append(en)

            for part in parts:
                label = self._label_for_part(part)
                if not self._exists(z, part):
                    continue

                # relationships for this part
                rels = self._read_rels(z, part)
                rid_to_link: Dict[str, Tuple[str, Optional[str]]] = {
                    r["Id"]: (r["Target"], r.get("TargetMode"))
                    for r in rels
                    if r["Type"] == REL_TYPES["hyperlink"]
                }

                # parse the XML
                try:
                    root = self._read_xml(z, part)
                except Exception:
                    logger.exception("Failed to parse %s", part)
                    continue

                # 1) Standard <w:hyperlink>
                for h in root.findall(".//w:hyperlink", NS):
                    rid = h.get(f"{{{NS['r']}}}id")
                    anchor = h.get("anchor")
                    if rid and rid in rid_to_link:
                        target, _mode = rid_to_link[rid]
                        code, status, desc = _check_url(target)
                        rows.append({
                            "part": label,
                            "type": "External",
                            "link": target,
                            "status": status,
                            "code": code,
                            "description": desc,
                        })
                    elif anchor:
                        rows.append({
                            "part": label,
                            "type": "Internal",
                            "link": f"#{anchor}",
                            "status": "OK",
                            "code": "",
                            "description": "Bookmark",
                        })

                # 2) Field codes (fldSimple / instrText)
                seen: Set[str] = set()
                for url in self._urls_from_fldSimple(root) + self._urls_from_instr_text(root):
                    if url in seen:
                        continue
                    seen.add(url)
                    code, status, desc = _check_url(url)
                    rows.append({
                        "part": label,
                        "type": "External",
                        "link": url,
                        "status": status,
                        "code": code,
                        "description": desc,
                    })

                # 3) Drawing parts: find any drawing rel and look for a:hlinkClick
                for r in rels:
                    if r["Type"] != REL_TYPES.get("drawing"):
                        continue
                    drawing_part = r["Target"]
                    if not self._exists(z, drawing_part):
                        continue

                    d_rels = self._read_rels(z, drawing_part)
                    d_rid_to_link: Dict[str, Tuple[str, Optional[str]]] = {
                        rr["Id"]: (rr["Target"], rr.get("TargetMode"))
                        for rr in d_rels
                        if rr["Type"] == REL_TYPES["hyperlink"]
                    }

                    try:
                        droot = self._read_xml(z, drawing_part)
                    except Exception:
                        logger.exception("Failed to parse drawing %s", drawing_part)
                        continue

                    # look for any a:hlinkClick anywhere (best-effort)
                    for elem in droot.findall(".//a:hlinkClick", NS):
                        rid = elem.get(f"{{{NS['r']}}}id")
                        if not rid or rid not in d_rid_to_link:
                            continue
                        url, _m = d_rid_to_link[rid]
                        code, status, desc = _check_url(url)
                        rows.append({
                            "part": label,
                            "type": "External",
                            "link": url,
                            "status": status,
                            "code": code,
                            "description": desc,
                        })

        return rows
