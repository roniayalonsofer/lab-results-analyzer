"""
base.py
-------
Abstract base class for all lab report parsers.
Every parser must implement `parse(file_obj) -> list[dict]`.
"""

from __future__ import annotations

import io
from abc import ABC, abstractmethod


class BaseParser(ABC):
    """Base class for laboratory report parsers."""

    # Override in subclass with the display name of the lab
    LAB_NAME: str = "Unknown Lab"

    @abstractmethod
    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        """
        Parse a lab report file.

        Parameters
        ----------
        file_obj : io.BytesIO
            Binary file-like object of the uploaded report.

        Returns
        -------
        list[dict]
            Each dict represents one measurement record with keys:
                - compound  (str)  : chemical name
                - cas       (str)  : CAS number
                - value     (float | None) : measured concentration
                - flag      (str)  : '', '<', '>', 'ND'
                - unit      (str)  : unit of measurement (e.g. 'µg/m³')
                - sample_id (str)  : sample identifier (optional)
        """
        ...

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__} lab='{self.LAB_NAME}'>"
