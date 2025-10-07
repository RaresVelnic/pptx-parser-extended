# xlsx_ext/extractors/__init__.py
from .links import XlsxLinkExtractor
from .descriptions import XlsxDescriptionExtractor

__all__ = ["XlsxLinkExtractor", "XlsxDescriptionExtractor"]
