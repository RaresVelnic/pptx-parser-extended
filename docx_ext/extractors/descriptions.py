# docx_ext/extractors/descriptions.py
"""
DocxDescriptionExtractor — extract alt text ("descriptions") from DOCX.

We scan these parts:
- word/document.xml
- word/header*.xml
- word/footer*.xml

Sources of alt text:
- DrawingML pictures: <wp:docPr descr="..." title="...">
- (Optional) VML shapes: <v:shape o:title="..."> (legacy)

Output schema:
[
  {"part": "document", "descriptions": ["..", ...]},
  {"part": "header1",  "descriptions": ["..", ...]},
  {"part": "footer1",  "descriptions": ["..", ...]},
  ...
]
(Parts with no descriptions are omitted — keeps UI clean later.)
"""

from __future__ import annotations

import logging
from typing import List, Dict
from lxml import etree

from .base import DocxBaseExtractor, NS

logger = logging.getLogger(__name__)


class DocxDescriptionExtractor(DocxBaseExtractor):
    """Extract alt text (descriptions) from DOCX drawings/shapes."""

    def extract(self, docx_bytes: bytes) -> List[Dict]:
        rows: List[Dict] = []
        with self._zip(docx_bytes) as z:
            targets = [("document", self._doc_part())]
            targets += [(f"header{idx+1}", p) for idx, p in enumerate(self._header_parts(z))]
            targets += [(f"footer{idx+1}", p) for idx, p in enumerate(self._footer_parts(z))]

            for label, part in targets:
                if not self._exists(z, part):
                    continue
                root = self._read_xml(z, part)
                found: List[str] = []

                # DrawingML: <wp:docPr descr="..." title="...">
                for node in root.findall(".//wp:docPr", NS):
                    descr = node.get("descr") or node.get("title")
                    if descr:
                        found.append(descr.strip())

                # Legacy VML: <v:shape o:title="...">
                for node in root.findall(".//v:shape", NS):
                    title = node.get("{%s}title" % NS["o"])
                    if title:
                        found.append(title.strip())

                # Only keep parts that have at least one *real* description string
                found_clean = [s for s in (d.strip() for d in found) if s]
                if found_clean:
                    rows.append({"part": label, "descriptions": found_clean})

        return rows
