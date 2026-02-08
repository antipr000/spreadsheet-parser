"""
Detector for heading / section-title regions.

Heuristic rules (all should be satisfied):
  - Region spans at most 3 rows
  - At least one cell is bold, has a larger-than-default font, or is merged
  - Few distinct non-empty values (≤ 3)
  - No formulas (headings are plain labels)
"""

from __future__ import annotations

import logging
from typing import Optional

from ai.factory import get_decision_service
from ai.response_parser import parse_llm_json
from detection.base import Detector
from dto.blocks import Block, HeadingBlock
from dto.region import RegionData
from prompts.detection import get_heading_detection_prompt

logger = logging.getLogger(__name__)

# Font sizes at or above this threshold hint at a heading.
_HEADING_FONT_SIZE_THRESHOLD = 12


class HeadingDetector(Detector):

    # ------------------------------------------------------------------
    # Heuristic
    # ------------------------------------------------------------------

    def detect(self, region: RegionData) -> Optional[Block]:
        # Must be a small vertical region
        if region.num_rows > 3:
            return None

        non_empty = region.non_empty_cells
        if not non_empty:
            return None

        # Few distinct values
        distinct_values = {c.value for c in non_empty}
        if len(distinct_values) > 3:
            return None

        # Must NOT contain formulas (headings are labels)
        if any(c.formula for c in non_empty):
            return None

        # At least one "heading-like" signal: bold, large font, or merged
        has_bold = any(c.font_bold for c in non_empty)
        has_large_font = any(
            c.font_size is not None and c.font_size >= _HEADING_FONT_SIZE_THRESHOLD
            for c in non_empty
        )
        has_merge = any(c.merged_with is not None for c in non_empty)

        if not (has_bold or has_large_font or has_merge):
            return None

        # Assemble the heading text
        text = " ".join(c.value for c in non_empty if c.value)

        return HeadingBlock(
            bounding_box=region.bounding_box,
            text=text,
            cells=non_empty,
        )

    # ------------------------------------------------------------------
    # AI-assisted
    # ------------------------------------------------------------------

    def detect_with_ai(self, region: RegionData) -> Optional[Block]:
        prompt = get_heading_detection_prompt(region.non_empty_cells)
        ai = get_decision_service()
        raw = ai.get_decision(prompt)

        parsed = parse_llm_json(raw)
        if not isinstance(parsed, dict):
            return None

        if not parsed.get("is_heading", False):
            return None

        text = parsed.get("text", "")
        if not text:
            # Fallback: join cell values
            text = " ".join(c.value for c in region.non_empty_cells if c.value)

        return HeadingBlock(
            bounding_box=region.bounding_box,
            text=text,
            cells=region.non_empty_cells,
        )
