# docx_ext/extractors/links.py
"""
DocxLinkExtractor — extract and check links in DOCX.

We scan these parts:
- word/document.xml
- word/header*.xml
- word/footer*.xml
- word/footnotes.xml (if present)
- word/endnotes.xml  (if present)

Sources:
- <w:hyperlink r:id="rIdX"> with a Relationship of type 'hyperlink'
- Field codes (HYPERLINK "...") in:
    * <w:fldSimple w:instr="...">
    * <w:instrText> nodes (complex fields)

For external links (http/https), we perform a HEAD with fallback to GET.
For internal (bookmarks, file rels), we mark as "Internal".

Output schema:
[
  {"part": "document", "link": "http...", "type":"External", "status":"OK", "code":200, "description":"..."},
  {"part": "header1",  "link": "#bookmark", "type":"Internal", ...},
  ...
]
"""

from __future__ import annotations

import re
import logging
import posixpath
from typing import List, Dict, Tuple, Optional
from lxml import etree

import requests

from .base import DocxBaseExtractor, NS, REL_TYPES

logger = logging.getLogger(__name__)

HYPERLINK_RE = re.compile(r'HYPERLINK\s+"([^"]+)"', re.IGNORECASE)


def http_code_meaning(code: Optional[int]) -> str:
    if code is None:
        return "No response"
    try:
        code = int(code)
    except Exception:
        return "Unknown"
    if 200 <= code < 300:
        return "OK"
    elif 300 <= code < 400:
        return "Redirect"
    elif 400 <= code < 500:
        return "Client Error"
    elif 500 <= code < 600:
        return "Server Error"
    else:
        return "Other"


def check_url(url: str) -> Tuple[Optional[int], str, str]:
    headers = {
        "User-Agent": "pptx-extended-parser/1.0",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
    }
    try:
        resp = requests.head(url, allow_redirects=True, timeout=5, headers=headers)
        if resp.status_code in (403, 405, 501):
            resp = requests.get(url, allow_redirects=True, timeout=5, stream=True, headers=headers)
            resp.close()
        return resp.status_code, http_code_meaning(resp.status_code), resp.reason
    except Exception as e:
        return None, "Bad link", str(e)


class DocxLinkExtractor(DocxBaseExtractor):
    """Extracts hyperlinks from DOCX and optionally checks http(s)."""

    def _links_from_rels(self, z, part: str, root) -> List[Dict]:
        """Collect links declared by <w:hyperlink r:id=...> using rels table."""
        rels = {r["Id"]: r for r in self._read_rels(z, part)}
        results: List[Dict] = []

        for hl in root.findall(".//w:hyperlink", NS):
            rid = hl.get("{%s}id" % NS["r"])
            anchor = hl.get("anchor")  # internal bookmark
            if anchor and not rid:
                # purely internal bookmark link
                results.append({
                    "part": self._label_for_part(part),
                    "link": f"#{anchor}",
                    "type": "Internal",
                    "status": "OK",
                    "code": "",
                    "description": "Bookmark link",
                })
                continue

            if not rid or rid not in rels:
                continue
            rel = rels[rid]
            href = rel["Target"]
            is_external = rel["TargetMode"] == "External" or href.startswith("http")
            if is_external:
                code, status, reason = check_url(href)
                results.append({
                    "part": self._label_for_part(part),
                    "link": href,
                    "type": "External",
                    "status": status,
                    "code": code,
                    "description": reason,
                })
            else:
                results.append({
                    "part": self._label_for_part(part),
                    "link": href,
                    "type": "Internal",
                    "status": "OK",
                    "code": "",
                    "description": "Internal target",
                })
        return results

    def _links_from_field_codes(self, part_label: str, root) -> List[Dict]:
        """Parse field codes like HYPERLINK "http://..."."""
        out: List[Dict] = []

        # Simple fields
        for fld in root.findall(".//w:fldSimple", NS):
            instr = fld.get("instr") or ""
            for url in HYPERLINK_RE.findall(instr):
                out.append(self._external(url, part_label))

        # Complex fields: gather contiguous instrText runs
        instr_chunks: List[str] = []
        for node in root.findall(".//w:instrText", NS):
            text = (node.text or "").strip()
            if text:
                instr_chunks.append(text)
        if instr_chunks:
            big = " ".join(instr_chunks)
            for url in HYPERLINK_RE.findall(big):
                out.append(self._external(url, part_label))

        return out

    def _external(self, url: str, part_label: str) -> Dict:
        code, status, reason = check_url(url)
        return {
            "part": part_label,
            "link": url,
            "type": "External",
            "status": status,
            "code": code,
            "description": reason,
        }

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

    def extract(self, docx_bytes: bytes) -> List[Dict]:
        results: List[Dict] = []
        with self._zip(docx_bytes) as z:
            parts = [self._doc_part()]
            parts += self._header_parts(z)
            parts += self._footer_parts(z)
            fn = self._footnotes_part(z)
            en = self._endnotes_part(z)
            if fn: parts.append(fn)
            if en: parts.append(en)

            for p in parts:
                if not self._exists(z, p):
                    continue
                root = self._read_xml(z, p)
                label = self._label_for_part(p)
                results.extend(self._links_from_rels(z, p, root))
                results.extend(self._links_from_field_codes(label, root))
        return results
