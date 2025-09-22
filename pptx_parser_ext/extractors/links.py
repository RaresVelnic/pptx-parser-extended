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
        - Checks internal links (to other slides, images, embedded files) by verifying the target exists.

        Args:
            pptx_bytes: Binary PPTX file content.

        Returns:
            list[dict]: Each describing a found link and its status:
                {
                  "slide": int,
                  "link": str,
                  "type": "External" | "Internal Slide" | "Internal File" | "Internal",
                  "status": str,   # e.g., "OK", "Redirect", "Client Error", "Server Error", "Broken/Missing", "Bad link"
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
                    target = rel["target"]
                    rel_type = rel["type"]  # not used for categorization in original, but kept
                    target_mode = rel.get("target_mode", None)

                    # External web links
                    if (target_mode == "External" and target.startswith("http")) or target.startswith("http"):
                        result = {"slide": slide_idx, "link": target, "type": "External"}
                        code, status, desc = self._check_url(target)
                        result.update({"status": status, "code": code, "description": desc})
                        results.append(result)
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

        Returns:
            Mapping from relationship Id to dict with keys: 'target', 'type', 'target_mode'
        """
        rels: Dict[str, Dict[str, Optional[str]]] = {}
        if rels_path in z.namelist():
            tree = etree.fromstring(z.read(rels_path))
            for rel in tree.findall(f".//{{{REL_NS}}}Relationship"):
                rel_id = rel.get("Id")
                target = rel.get("Target")
                rel_type = rel.get("Type")
                target_mode = rel.get("TargetMode")  # "External" or "Internal" (or None)
                if rel_id:
                    rels[rel_id] = {"target": target, "type": rel_type, "target_mode": target_mode}
        return rels

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

        - HEAD first with cache-busting headers; fallback to GET on 403/405/501.
        - Returns (status_code, status_text, description).
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
