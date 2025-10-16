# docx_ext/extractors/links.py
from __future__ import annotations

import io
import re
import posixpath
import zipfile
import logging
from typing import Dict, List, Optional, Tuple

from lxml import etree

# Namespaces (local copy)
NS = {
    "w":  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

REL_TYPES = {
    "hyperlink": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
}

logger = logging.getLogger(__name__)

# --- Optional: simple HTTP checker (same spirit as your PPTX checker) ---
def _http_check(url: str, timeout: float = 6.0) -> Tuple[str, Optional[int], str]:
    """
    Return (status, code, description) for external HTTP(S) URLs.
    status: "OK" | "Bad link" | "Error"
    """
    try:
        import requests
    except Exception:
        return ("Error", None, "requests not available")

    try:
        # try HEAD first
        r = requests.head(url, allow_redirects=True, timeout=timeout)
        if 200 <= r.status_code < 400:
            return ("OK", r.status_code, "OK")
        # fallback to GET (some servers dislike HEAD)
        r = requests.get(url, allow_redirects=True, timeout=timeout)
        if 200 <= r.status_code < 400:
            return ("OK", r.status_code, "OK")
        return ("Bad link", r.status_code, r.reason or "HTTP error")
    except requests.exceptions.SSLError as e:
        return ("Bad link", None, f"SSL error: {e}")
    except requests.exceptions.ConnectionError as e:
        return ("Bad link", None, f"Connection error: {e}")
    except requests.exceptions.Timeout:
        return ("Bad link", None, "Timeout")
    except Exception as e:
        return ("Error", None, str(e))


# ---------- helpers ----------
def _rels_path(part_path: str) -> str:
    d, b = posixpath.split(part_path)
    return posixpath.join(d, "_rels", b + ".rels")

def _norm_join(base_part: str, target: str) -> str:
    base_dir = posixpath.dirname(base_part)
    p = posixpath.normpath(posixpath.join(base_dir, target))
    return p.lstrip("/")

def _read_xml(zf: zipfile.ZipFile, path: str):
    with zf.open(path) as f:
        return etree.fromstring(f.read())

def _read_rels(zf: zipfile.ZipFile, part_path: str) -> Dict[str, Dict[str, str]]:
    """Return {rId: {Type, Target, TargetMode}} for a given part."""
    rp = _rels_path(part_path)
    out: Dict[str, Dict[str, str]] = {}
    if rp not in zf.namelist():
        return out
    rels = _read_xml(zf, rp)
    for rel in rels.findall(f".//{{{REL_NS}}}Relationship"):
        rid = rel.get("Id")
        rtype = rel.get("Type")
        raw_target = rel.get("Target") or ""
        tmode = rel.get("TargetMode")
        # external if TargetMode=External OR has a scheme like http://, mailto:, etc.
        is_external = (tmode and tmode.lower() == "external") or ":" in raw_target
        target = raw_target if is_external else _norm_join(part_path, raw_target)
        out[rid] = {"Type": rtype, "Target": target, "TargetMode": tmode}
    return out

def _collect_text(el: etree._Element) -> str:
    texts = el.findall(".//w:t", namespaces=NS)
    return "".join((t.text or "") for t in texts).strip()

# Parse instr string like: HYPERLINK "https://..."   or   HYPERLINK \\l "Bookmark"
_INSTR_RE = re.compile(r'HYPERLINK\b(?P<rest>.*)', re.IGNORECASE | re.DOTALL)
_QUOTED_RE = re.compile(r'"([^"]+)"')

def _parse_hyperlink_instr(instr: str) -> Tuple[Optional[str], bool]:
    """
    Returns (target, is_internal).
    Examples:
      'HYPERLINK "https://example.com"'           -> ("https://example.com", False)
      'HYPERLINK \\l "BookmarkName"'              -> ("BookmarkName", True)
      'HYPERLINK  \l  "Heading_1"'                -> ("Heading_1", True)
    """
    m = _INSTR_RE.search(instr or "")
    if not m:
        return (None, False)
    rest = m.group("rest")
    # internal if \l appears anywhere before a quoted arg
    is_int = "\\l" in rest or r"\l" in rest
    qm = _QUOTED_RE.search(rest)
    if not qm:
        return (None, is_int)
    target = qm.group(1).strip()
    return (target, is_int)

# --- Scheme classifier for external targets ---
_SCHEME_RE = re.compile(r'^[a-zA-Z][a-zA-Z0-9+.\-]*:')

def _classify_and_check_target(target: str) -> Tuple[str, str, Optional[int], str]:
    """
    Decide link 'type' and perform checks when appropriate.

    Returns: (type, status, code, description_note)
      - type: "External" | "Email"
      - status: "OK" | "Bad link" | "Error" | "Skipped"
      - code: HTTP status code or None
      - description_note: short note to append when we skip checks, etc.
    """
    if not target:
        return ("External", "Error", None, "Empty target")

    tl = target.lower()

    # Email links: mailto:
    if tl.startswith("mailto:"):
        return ("Email", "OK", None, "Email link")

    # HTTP(S) -> run checker
    if tl.startswith("http://") or tl.startswith("https://"):
        status, code, desc = _http_check(target)
        return ("External", status, code, desc or "")

    # Any other scheme (tel:, ftp:, file:, etc.) -> don't HTTP check
    if _SCHEME_RE.match(target):
        return ("External", "Skipped", None, "Non-HTTP scheme")

    # Fallback: treat as external but don't check (unlikely in DOCX rels)
    return ("External", "Skipped", None, "Unrecognized scheme")

def _merge_description(visible: str, note: str, status: str) -> str:
    """
    Append note to visible text when status isn't OK and a note is present.
    Keeps UI noise low.
    """
    visible = visible or "(no text)"
    if note and status not in ("OK",):
        return f"{visible} ({note})"
    return visible


class DocxLinkExtractor:
    """
    Finds both external and internal (in-document) links in:
      - word/document.xml
      - word/header*.xml
      - word/footer*.xml

    Output rows look like:
      {
        "part": "document" | "header1" | "footer1",
        "link": "https://..." or "#Bookmark",
        "type": "External" | "Internal" | "Email",
        "status": "OK" | "Bad link" | "Error" | "Skipped",
        "code": 200 | None,
        "description": "<visible text or brief note>"
      }
    """

    def extract(self, docx_bytes: bytes) -> List[Dict]:
        rows: List[Dict] = []
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
            names = set(z.namelist())
            parts = ["word/document.xml"]
            parts += sorted(p for p in names if p.startswith("word/header") and p.endswith(".xml"))
            parts += sorted(p for p in names if p.startswith("word/footer") and p.endswith(".xml"))

            for part in parts:
                try:
                    root = _read_xml(z, part)
                except Exception:
                    logger.exception("Failed to parse %s", part)
                    continue

                rels = _read_rels(z, part)
                part_label = self._label_for_part(part)

                # --- 1) <w:hyperlink ...> ---
                for h in root.findall(".//w:hyperlink", namespaces=NS):
                    rid = h.get(f"{{{NS['r']}}}id")
                    anchor = h.get("{" + NS["w"] + "}anchor") or h.get("anchor")
                    visible = _collect_text(h) or "(no text)"

                    if anchor:
                        # Internal (bookmark/cross-ref)
                        rows.append({
                            "part": part_label,
                            "link": f"#{anchor}",
                            "type": "Internal",
                            "status": "OK",
                            "code": None,
                            "description": visible,
                        })
                        continue

                    if rid and rid in rels and rels[rid]["Type"] == REL_TYPES["hyperlink"]:
                        url = rels[rid]["Target"]
                        link_type, status, code, note = _classify_and_check_target(url)
                        rows.append({
                            "part": part_label,
                            "link": url,
                            "type": link_type,  # "External" or "Email"
                            "status": status,
                            "code": code,
                            "description": _merge_description(visible, note, status),
                        })

                # --- 2) <w:fldSimple w:instr="HYPERLINK ..."> ---
                for fld in root.findall(".//w:fldSimple", namespaces=NS):
                    instr = fld.get(f"{{{NS['w']}}}instr") or ""
                    target, is_internal = _parse_hyperlink_instr(instr)
                    if not target:
                        continue
                    visible = _collect_text(fld) or "(no text)"
                    if is_internal:
                        rows.append({
                            "part": part_label,
                            "link": f"#{target}",
                            "type": "Internal",
                            "status": "OK",
                            "code": None,
                            "description": visible,
                        })
                    else:
                        link_type, status, code, note = _classify_and_check_target(target)
                        rows.append({
                            "part": part_label,
                            "link": target,
                            "type": link_type,  # "External" or "Email"
                            "status": status,
                            "code": code,
                            "description": _merge_description(visible, note, status),
                        })

                # --- 3) Complex fields: w:fldChar + w:instrText ---
                rows.extend(self._scan_complex_fields(root, part_label))

        return rows

    def _label_for_part(self, part_path: str) -> str:
        """
        word/document.xml -> 'document'
        word/header1.xml  -> 'header1'
        word/footer2.xml  -> 'footer2'
        """
        base = posixpath.basename(part_path)
        if base == "document.xml":
            return "document"
        if base.startswith("header"):
            return base[:-4]  # strip .xml
        if base.startswith("footer"):
            return base[:-4]
        return base

    def _scan_complex_fields(self, root: etree._Element, part_label: str) -> List[Dict]:
        """
        Handle fields that look like:
          <w:fldChar w:fldCharType="begin"/>
          <w:instrText> HYPERLINK "\\l" "Bookmark" </w:instrText>
          <w:fldChar w:fldCharType="separate"/>
          ... link display text runs ...
          <w:fldChar w:fldCharType="end"/>

        We accumulate instrText between begin..(separate|end),
        and collect the visible text between (separate)..end.
        """
        out: List[Dict] = []

        # Walk the tree in document order
        def walk(e):
            yield e
            for c in e:
                yield from walk(c)

        collecting_instr = False
        instr_buf: List[str] = []
        collecting_result = False
        result_text_nodes: List[str] = []

        def flush():
            nonlocal instr_buf, result_text_nodes
            instr = "".join(instr_buf) if instr_buf else ""
            target, is_internal = _parse_hyperlink_instr(instr)
            visible = "".join(result_text_nodes).strip() or "(no text)"
            if target:
                if is_internal:
                    out.append({
                        "part": part_label,
                        "link": f"#{target}",
                        "type": "Internal",
                        "status": "OK",
                        "code": None,
                        "description": visible,
                    })
                else:
                    link_type, status, code, note = _classify_and_check_target(target)
                    out.append({
                        "part": part_label,
                        "link": target,
                        "type": link_type,  # "External" or "Email"
                        "status": status,
                        "code": code,
                        "description": _merge_description(visible, note, status),
                    })
            # reset buffers
            instr_buf = []
            result_text_nodes = []

        for node in walk(root):
            tag = node.tag if isinstance(node.tag, str) else ""
            if tag == f"{{{NS['w']}}}fldChar":
                ftype = node.get(f"{{{NS['w']}}}fldCharType")
                if ftype == "begin":
                    collecting_instr = True
                    collecting_result = False
                    instr_buf = []
                    result_text_nodes = []
                elif ftype == "separate":
                    collecting_instr = False
                    collecting_result = True
                elif ftype == "end":
                    if collecting_instr or collecting_result:
                        flush()
                    collecting_instr = False
                    collecting_result = False

            elif tag == f"{{{NS['w']}}}instrText" and collecting_instr:
                if node.text:
                    instr_buf.append(node.text)

            elif tag == f"{{{NS['w']}}}t" and collecting_result:
                if node.text:
                    result_text_nodes.append(node.text)

        # if document ended in a field without closing 'end'
        if collecting_instr or collecting_result:
            flush()

        return out
