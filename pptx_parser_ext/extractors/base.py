from abc import ABC, abstractmethod
from typing import Any

class BaseExtractor(ABC):
    @abstractmethod
    def extract(self, pptx_bytes: bytes) -> Any:
        """Return structured data extracted from the PPTX bytes."""
        raise NotImplementedError
