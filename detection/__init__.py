"""
Region-type detectors.

Each detector implements the ``Detector`` base class and provides both a
fast heuristic path and an LLM-assisted path.

The canonical evaluation order is:
  1. HeadingDetector   — small, bold/merged label rows
  2. KeyValueDetector  — 2-column form-like layouts
  3. TextDetector      — free-text / notes / disclaimers
  4. TableDetector     — structured columnar data (default / fallback)
"""

from detection.heading import HeadingDetector
from detection.key_value import KeyValueDetector
from detection.text import TextDetector
from detection.table import TableDetector
from detection.base import Detector

__all__ = [
    "Detector",
    "HeadingDetector",
    "KeyValueDetector",
    "TextDetector",
    "TableDetector",
]
