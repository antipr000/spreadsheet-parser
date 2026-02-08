"""
Base class for all region-type detectors.

Each detector answers two questions about a candidate ``RegionData``:
  1. **detect** (heuristic) — does this region match my type?  Returns a
     block DTO if yes, ``None`` if no.
  2. **detect_with_ai** — ask the LLM to confirm / extract structure.
     Returns a block DTO if yes, ``None`` if no.

The caller decides whether to use heuristic-only, AI-only, or
heuristic-first-then-AI.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Optional

from dto.blocks import Block
from dto.region import RegionData


class Detector(ABC):
    """Interface that every block-type detector must implement."""

    @abstractmethod
    def detect(self, region: RegionData) -> Optional[Block]:
        """
        Pure heuristic detection.

        Returns a fully-populated block DTO if the region matches this
        detector's type, or ``None`` if it does not.
        """
        ...

    @abstractmethod
    def detect_with_ai(self, region: RegionData) -> Optional[Block]:
        """
        AI-assisted detection.

        Sends the region's cell data to the LLM and parses the response
        into a block DTO.  Returns ``None`` if the LLM says this region
        does not match the type.
        """
        ...
