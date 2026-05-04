"""
lab_value_parser.py
-------------------
Utilities for parsing raw lab value strings such as:
  '<0.5', '0.5', 'ND', 'N/A', '1.2E-3', '> 100'

Returns numeric floats, with helpers for detection-limit handling.
"""

import re


# Sentinel for non-detect / below detection limit
NON_DETECT = None
ABOVE_RANGE = None


class LabValueParser:
    """Parse raw laboratory measurement strings into numeric values."""

    # Regex patterns
    _RE_LESS_THAN    = re.compile(r'^[<＜]\s*(\d+\.?\d*(?:[eE][+-]?\d+)?)')
    _RE_GREATER_THAN = re.compile(r'^[>＞]\s*(\d+\.?\d*(?:[eE][+-]?\d+)?)')
    _RE_NUMBER       = re.compile(r'^(\d+\.?\d*(?:[eE][+-]?\d+)?)')

    NON_DETECT_STRINGS = {
        'nd', 'n/a', 'na', '<dl', '<mdl', '<rl', 'bdl', 'not detected',
        'לא זוהה', 'מתחת לגבול', '<גבול', '--', '-', ''
    }

    def __init__(self, default_nd_factor: float = 0.5):
        """
        Parameters
        ----------
        default_nd_factor : float
            When a '<' value is found, multiply the detection limit by this
            factor to assign a numeric value (e.g. 0.5 → half the DL).
        """
        self.nd_factor = default_nd_factor

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------
    def parse(self, raw: str) -> tuple[float | None, str]:
        """
        Parse a raw value string.

        Returns
        -------
        (value, flag) where:
            value : float | None   — numeric value (None = truly ND)
            flag  : str            — '', '<', '>', 'ND'
        """
        if raw is None:
            return None, 'ND'

        s = str(raw).strip()

        if s.lower() in self.NON_DETECT_STRINGS:
            return None, 'ND'

        m = self._RE_LESS_THAN.match(s)
        if m:
            dl = float(m.group(1))
            return dl * self.nd_factor, '<'

        m = self._RE_GREATER_THAN.match(s)
        if m:
            return float(m.group(1)), '>'

        m = self._RE_NUMBER.match(s)
        if m:
            return float(m.group(1)), ''

        # Fallback: try direct float conversion
        try:
            return float(s), ''
        except ValueError:
            return None, 'ND'

    def parse_value(self, raw: str) -> float | None:
        """Convenience: return only the numeric value."""
        val, _ = self.parse(raw)
        return val

    def is_non_detect(self, raw: str) -> bool:
        """Return True if the raw string represents a non-detect."""
        _, flag = self.parse(raw)
        return flag == 'ND'

    def is_below_limit(self, raw: str) -> bool:
        """Return True if value is below detection limit ('<' prefix)."""
        _, flag = self.parse(raw)
        return flag == '<'
