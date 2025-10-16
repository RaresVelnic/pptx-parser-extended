"""
links.py — OOP wrapper for extracting & checking links inside PPTX files.

Behavior preserved from original `check_links_in_pptx` + helpers:
- Enumerates slide relationships and resolves external (HTTP/S) & internal links.
- External links: HEAD with cache-busting headers; fallback to GET on 403/405/501.
- Internal links: normalized path existence checks within the ZIP.
- Output: list of dicts with slide, link, type, status, code/description (same as before).
"""

from __future__ import annotations

import io
import zipfile
import logging
import requests
from lxml import etree
import posixpath
from typing import Dict, List, Tuple, Optional, Any

logger = logging.getLogger(__name__)

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class LinkCheckExtractor:
    """Extract & check links in a PPTX archive."""

    # ---------------------- Public API ----------------------

    def extract(self, pptx_bytes: bytes) -> List[Dict[str, Any]]:
        """
        Extracts and checks all links in a PPTX file.

        - Checks external HTTP(S) links (status code, reachable).
        - Flags mailto: links as Email (OK, no HTTP check).
        - Skips checks for other non-HTTP schemes (tel:, ftp:, file:, …).
        - Checks internal links (to slides/media/embeddings) by verifying the target exists.

        Returns rows like:
        {
          "slide": int,
          "link": str,
          "type": "External" | "Email" | "Internal Slide" | "Internal File" | "Internal",
          "status": str,
          "code": int|str|None,
          "description": str
        }
        """
        results: List[Dict[str, Any]] = []
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
            all_files = set(z.namelist())
            slide_files = sorted(
                [n for n in all_files if n.startswith("ppt/slides/slide") and n.endswith(".xml")],
                key=lambda x: int("".join(filter(str.isdigit, x))),
            )

            for slide_idx, slide_file in enumerate(slide_files, start=1):
                logger.info(f"Checking Slide {slide_idx}")
                rels_path = slide_file.replace("slides/", "slides/_rels/") + ".rels"
                rels = self._get_relationships(z, rels_path)
                tree = etree.fromstring(z.read(slide_file))

                for elem in tree.iter():
                    r_id = elem.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    if not r_id or r_id not in rels:
                        continue

                    rel = rels[r_id]
                    target = rel["target"] or ""
                    target_mode = rel.get("target_mode", None)

                    # If relationship is External or the target clearly has a scheme, treat as external/email.
                    if self._is_external_like(target_mode, target):
                        link_type, status, code, note = self._classify_and_check_target(target)
                        results.append({
                            "slide": slide_idx,
                            "link": target,
                            "type": link_type,                   # "External" or "Email"
                            "status": status,                    # "OK" / "Bad link" / "Skipped" / "Error"
                            "code": code if code is not None else "",
                            "description": note,
                        })
                    else:
                        # Internal references: normalize and check existence
                        slide_dir = posixpath.dirname(slide_file)
                        normalized_path = posixpath.normpath(posixpath.join(slide_dir, target))
                        exists = normalized_path in all_files

                        if "/slides/" in normalized_path:
                            link_type = "Internal Slide"
                        elif "/media/" in normalized_path or "/embeddings/" in normalized_path:
                            link_type = "Internal File"
                        else:
                            link_type = "Internal"

                        result = {"slide": slide_idx, "link": target, "type": link_type}
                        if exists:
                            result.update(
                                {"status": "OK", "code": "", "description": f"Target exists: {normalized_path}"}
                            )
                        else:
                            result.update(
                                {"status": "Broken/Missing", "code": "", "description": f"Target missing: {normalized_path}"}
                            )
                        results.append(result)

        return results

    # ---------------------- Internals (helpers) ----------------------

    def _get_relationships(self, z: zipfile.ZipFile, rels_path: str) -> Dict[str, Dict[str, Optional[str]]]:
        """
        Extracts relationships from a .rels XML file within a pptx archive.
        Returns: { rId: { 'target': str, 'type': str, 'target_mode': str|None } }
        """
        rels: Dict[str, Dict[str, Optional[str]]] = {}
        if rels_path in z.namelist():
            tree = etree.fromstring(z.read(rels_path))
            for rel in tree.findall(f".//{{{REL_NS}}}Relationship"):
                rel_id = rel.get("Id")
                target = rel.get("Target")
                rel_type = rel.get("Type")
                target_mode = rel.get("TargetMode")  # "External" or None
                if rel_id:
                    rels[rel_id] = {"target": target, "type": rel_type, "target_mode": target_mode}
        return rels

    def _is_external_like(self, target_mode: Optional[str], target: str) -> bool:
        """
        Decide if a relationship target should be treated as an external-style link.
        """
        t = (target or "").lower()
        return (
            (target_mode or "").lower() == "external" or
            t.startswith("http://") or t.startswith("https://") or
            t.startswith("mailto:") or
            ":" in t  # any scheme (tel:, ftp:, file:, etc.)
        )

    def _http_code_meaning(self, code: Optional[int]) -> str:
        """Converts HTTP status code to a human-readable meaning."""
        if code is None:
            return "No response"
        try:
            code_int = int(code)
        except Exception:
            return "Unknown"
        if 200 <= code_int < 300:
            return "OK"
        elif 300 <= code_int < 400:
            return "Redirect"
        elif 400 <= code_int < 500:
            return "Client Error"
        elif 500 <= code_int < 600:
            return "Server Error"
        else:
            return "Other"

    def _check_url(self, url: str) -> Tuple[Optional[int], str, str]:
        """
        Checks the status of an external (HTTP/HTTPS) URL.
        HEAD first with cache-busting headers; fallback to GET on 403/405/501.
        Returns (status_code, status_text, description).
        """
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
            return resp.status_code, self._http_code_meaning(resp.status_code), resp.reason
        except Exception as e:
            return None, "Bad link", str(e)

    def _classify_and_check_target(self, target: str) -> Tuple[str, str, Optional[int], str]:
        """
        Classify target + run checks when appropriate.
        Returns (type, status, code, note)
          - type: "External" | "Email"
          - status: "OK" | "Bad link" | "Error" | "Skipped"
          - code: HTTP status code or None
          - note: short note (e.g., 'Email link' / 'Non-HTTP scheme' / reason)
        """
        if not target:
            return ("External", "Error", None, "Empty target")

        t = target.lower()

        if t.startswith("mailto:"):
            return ("Email", "OK", None, "Email link")

        if t.startswith("http://") or t.startswith("https://"):
            code, status, desc = self._check_url(target)
            return ("External", status, code, desc or "")

        # Any other scheme (tel:, ftp:, file:, etc.) -> don't HTTP check
        return ("External", "Skipped", None, "Non-HTTP scheme")
