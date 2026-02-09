"""
Image Extractor — extracts embedded images from the worksheet and
sends them to a vision model for description.
"""

from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from ai.factory import get_decision_for_media_service
from dto.blocks import Block, ImageBlock
from dto.cell_data import CellData

from agentic_flow.dto.plan import PlannedBlock
from agentic_flow.extractors.base import BaseExtractor
from agentic_flow.prompts.image import get_image_description_prompt

logger = logging.getLogger(__name__)


class ImageExtractor(BaseExtractor):

    def extract(
        self,
        planned: PlannedBlock,
        grid: Dict[Tuple[int, int], CellData],
        merge_map: Dict[str, str],
        ws: Worksheet,
        wb: Workbook,
        *,
        computed_values: Optional[Dict[Tuple[str, str], Any]] = None,
        screenshot_bytes: Optional[bytes] = None,
    ) -> List[Block]:
        blocks: List[Block] = []

        # Try to extract embedded images from the worksheet
        images = getattr(ws, "_images", [])
        for img in images:
            description = self._describe_image(img)
            blocks.append(
                ImageBlock(
                    bounding_box=planned.bounding_box,
                    description=description,
                )
            )

        # If no images found but the planner identified an image block,
        # try using the full screenshot for description
        if not blocks:
            description = None
            if screenshot_bytes:
                description = self._describe_from_screenshot(screenshot_bytes)
            blocks.append(
                ImageBlock(
                    bounding_box=planned.bounding_box,
                    description=description or planned.description or "Embedded image",
                )
            )

        return blocks

    @staticmethod
    def _describe_image(img) -> Optional[str]:
        """
        Extract image bytes from an openpyxl Image and send to
        vision model.
        """
        try:
            # openpyxl Image stores data in _data or ref attribute
            img_data = getattr(img, "_data", None)
            if img_data is None:
                ref = getattr(img, "ref", None)
                if ref and hasattr(ref, "read"):
                    img_data = ref.read()

            if img_data is None:
                return None

            # Determine MIME type
            format_hint = getattr(img, "format", "png") or "png"
            mime = f"image/{format_hint.lower()}"

            prompt = get_image_description_prompt()
            ai = get_decision_for_media_service()
            return ai.get_decision_for_media(prompt, img_data, mime_type=mime)
        except Exception:
            logger.warning(
                "  [ImageExtractor] Failed to describe image", exc_info=True
            )
            return None

    @staticmethod
    def _describe_from_screenshot(screenshot_bytes: bytes) -> Optional[str]:
        """Use the full sheet screenshot as a fallback for image description."""
        try:
            prompt = get_image_description_prompt()
            ai = get_decision_for_media_service()
            return ai.get_decision_for_media(
                prompt, screenshot_bytes, mime_type="image/png"
            )
        except Exception:
            logger.warning(
                "  [ImageExtractor] Screenshot description failed", exc_info=True
            )
            return None
