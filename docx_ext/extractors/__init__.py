# docx_ext/extractors/__init__.py
from .descriptions import DocxDescriptionExtractor
from .links import DocxLinkExtractor
from .fonts import DocxFontExtractor

__all__ = [
    "DocxDescriptionExtractor",
    "DocxLinkExtractor",
    "DocxFontExtractor",
]
