"""
Detector for free-text / notes / disclaimer regions.

Heuristic rules:
  - Region has at most 2 meaningful columns (or a single merged cell)
  - Content is sentence-length or longer (average word count ≥ 4 per cell)
  - Not structured as a grid (most cells in a row are empty)
  - No heading-like formatting cues (large bold short text)
"""

from __future__ import annotations

import logging
from typing import Optional

from ai.factory import get_decision_service
from ai.response_parser import parse_llm_json
from detection.base import Detector
from dto.blocks import Block, TextBlock
from dto.region import RegionData
from prompts.detection import get_text_detection_prompt

logger = logging.getLogger(__name__)

# If the average word count across non-empty cells meets this threshold
# we consider it "prose-like".
_MIN_AVG_WORDS = 4


class TextDetector(Detector):

    # ------------------------------------------------------------------
    # Heuristic
    # ------------------------------------------------------------------

    def detect(self, region: RegionData) -> Optional[Block]:
        non_empty = region.non_empty_cells
        if not non_empty:
            return None

        # Count columns that actually have data
        occupied_cols = set()
        for c in non_empty:
            col_str = "".join(ch for ch in c.coordinate if ch.isalpha())
            occupied_cols.add(col_str)

        # Text blocks live in 1–2 columns (or a wide merged cell)
        if len(occupied_cols) > 2 and not any(c.merged_with for c in non_empty):
            return None

        # Average word count should suggest prose, not labels
        total_words = sum(len((c.value or "").split()) for c in non_empty)
        avg_words = total_words / len(non_empty) if non_empty else 0
        if avg_words < _MIN_AVG_WORDS:
            return None

        # Should not look like a heading (few cells, bold, short)
        if region.num_rows <= 2 and all(
            c.font_bold for c in non_empty if c.font_bold is not None
        ):
            return None

        text = "\n".join(c.value for c in non_empty if c.value)

        return TextBlock(
            bounding_box=region.bounding_box,
            text=text,
            cells=non_empty,
        )

    # ------------------------------------------------------------------
    # AI-assisted
    # ------------------------------------------------------------------

    def detect_with_ai(self, region: RegionData) -> Optional[Block]:
        prompt = get_text_detection_prompt(region.non_empty_cells)
        ai = get_decision_service()
        raw = ai.get_decision(prompt)

        parsed = parse_llm_json(raw)
        if not isinstance(parsed, dict):
            return None

        if not parsed.get("is_text", False):
            return None

        text = parsed.get("text", "")
        if not text:
            text = "\n".join(c.value for c in region.non_empty_cells if c.value)

        return TextBlock(
            bounding_box=region.bounding_box,
            text=text,
            cells=region.non_empty_cells,
        )
