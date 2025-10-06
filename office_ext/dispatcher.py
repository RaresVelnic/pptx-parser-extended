# office_ext/dispatcher.py
"""
FileDispatcher — routes uploaded Office files (.pptx, .docx, .xlsx)
to the appropriate OOP extractors based on selected modes.

- Keeps your existing PPTX extractors
- Adds lazy/optional support for DOCX/XLSX (safe if not added yet)
- Returns only the sections requested by `modes`
"""

from __future__ import annotations

import os
import logging
from typing import Any, Dict, List, Optional

# --- Always-available (your existing modules) ---
from pptx_parser_ext.extractors.descriptions import DescriptionExtractor as PptxDescriptionExtractor
from pptx_parser_ext.extractors.links import LinkCheckExtractor as PptxLinkExtractor
from pptx_parser_ext.extractors.fonts import FontUsageExtractor as PptxFontExtractor

logger = logging.getLogger(__name__)


def _try_import_docx_extractors():
    """Lazy-import DOCX extractors; return (desc, links, fonts) or (None, None, None)."""
    try:
        from docx_ext.extractors.descriptions import DocxDescriptionExtractor
        from docx_ext.extractors.links import DocxLinkExtractor
        from docx_ext.extractors.fonts import DocxFontExtractor
        return DocxDescriptionExtractor, DocxLinkExtractor, DocxFontExtractor
    except Exception as e:
        logger.debug("DOCX extractors not available yet: %s", e)
        return None, None, None


def _try_import_xlsx_extractors():
    """Lazy-import XLSX extractors; return (desc, links, fonts) or (None, None, None)."""
    try:
        from xlsx_ext.extractors.descriptions import XlsxDescriptionExtractor
        from xlsx_ext.extractors.links import XlsxLinkExtractor
        from xlsx_ext.extractors.fonts import XlsxFontExtractor
        return XlsxDescriptionExtractor, XlsxLinkExtractor, XlsxFontExtractor
    except Exception as e:
        logger.debug("XLSX extractors not available yet: %s", e)
        return None, None, None


class FileDispatcher:
    """
    Route bytes to the right extractors based on file extension.

    Usage:
        dispatcher = FileDispatcher()
        result = dispatcher.run(filename, content_bytes, modes)
        # result contains optional keys: descriptions, links, fonts, file_type, warnings
    """

    def __init__(self) -> None:
        # PPTX: always wired in from your current project
        self.pptx_desc = PptxDescriptionExtractor()
        self.pptx_links = PptxLinkExtractor()
        self.pptx_fonts = PptxFontExtractor()

        # DOCX/XLSX: optional (import only if present)
        DocxDesc, DocxLinks, DocxFonts = _try_import_docx_extractors()
        XlsxDesc, XlsxLinks, XlsxFonts = _try_import_xlsx_extractors()

        self.docx_desc = DocxDesc() if DocxDesc else None
        self.docx_links = DocxLinks() if DocxLinks else None
        self.docx_fonts = DocxFonts() if DocxFonts else None

        self.xlsx_desc = XlsxDesc() if XlsxDesc else None
        self.xlsx_links = XlsxLinks() if XlsxLinks else None
        self.xlsx_fonts = XlsxFonts() if XlsxFonts else None

    # ---------------- internal helpers ----------------

    @staticmethod
    def _kind(filename: str) -> str:
        ext = os.path.splitext(filename.lower())[1]
        if ext == ".pptx":
            return "pptx"
        if ext == ".docx":
            return "docx"
        if ext == ".xlsx":
            return "xlsx"
        return "unknown"

    # ---------------- public API ----------------

    def run(self, filename: str, content: bytes, modes: List[str]) -> Dict[str, Any]:
        """
        Execute only the requested modes for the detected file type.

        Args:
            filename: original filename (used to detect type)
            content: raw bytes of the uploaded Office file
            modes: list of selected modes (e.g. ["extract_description", "check_links", "analyze_fonts"])

        Returns:
            dict with keys:
                - "file_type": "pptx" | "docx" | "xlsx" | "unknown"
                - Optional sections: "descriptions", "links", "fonts"
                - Optional "warnings": list[str] if a requested mode isn't available
                - Optional "error": str for unsupported file types
        """
        kind = self._kind(filename)
        out: Dict[str, Any] = {"file_type": kind}
        warnings: List[str] = []

        if kind == "unknown":
            out["error"] = "Unsupported file type. Please upload a .pptx, .docx, or .xlsx file."
            return out

        try:
            # ---------- PPTX ----------
            if kind == "pptx":
                if "extract_description" in modes:
                    out["descriptions"] = self.pptx_desc.extract(content)
                if "check_links" in modes:
                    out["links"] = self.pptx_links.extract(content)
                if "analyze_fonts" in modes:
                    out["fonts"] = self.pptx_fonts.extract(content)

            # ---------- DOCX ----------
            elif kind == "docx":
                if "extract_description" in modes:
                    if self.docx_desc:
                        out["descriptions"] = self.docx_desc.extract(content)
                    else:
                        warnings.append("DOCX descriptions not available (module not installed).")
                if "check_links" in modes:
                    if self.docx_links:
                        out["links"] = self.docx_links.extract(content)
                    else:
                        warnings.append("DOCX links not available (module not installed).")
                if "analyze_fonts" in modes:
                    if self.docx_fonts:
                        out["fonts"] = self.docx_fonts.extract(content)
                    else:
                        warnings.append("DOCX fonts not available (module not installed).")

            # ---------- XLSX ----------
            elif kind == "xlsx":
                if "extract_description" in modes:
                    if self.xlsx_desc:
                        out["descriptions"] = self.xlsx_desc.extract(content)
                    else:
                        warnings.append("XLSX descriptions not available (module not installed).")
                if "check_links" in modes:
                    if self.xlsx_links:
                        out["links"] = self.xlsx_links.extract(content)
                    else:
                        warnings.append("XLSX links not available (module not installed).")
                if "analyze_fonts" in modes:
                    if self.xlsx_fonts:
                        out["fonts"] = self.xlsx_fonts.extract(content)
                    else:
                        warnings.append("XLSX fonts not available (module not installed).")

        except Exception as e:
            # Bubble a friendly error and log the stack
            logger.exception("Dispatcher run() failed for %s", filename)
            out["error"] = f"Failed to process {kind.upper()} file: {e}"

        if warnings:
            out["warnings"] = warnings
        return out
