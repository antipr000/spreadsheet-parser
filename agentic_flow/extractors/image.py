"""
Image Extractor — extracts embedded images from the worksheet and
sends the raw image bytes (PNG/JPEG) to the LLM for description.
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
    ) -> List[Block]:
        blocks: List[Block] = []

        # Extract embedded images from the worksheet
        images = getattr(ws, "_images", [])
        for img in images:
            description = self._describe_image(img)
            blocks.append(
                ImageBlock(
                    bounding_box=planned.bounding_box,
                    description=description,
                )
            )

        # If no images found, use the planner's description
        if not blocks:
            blocks.append(
                ImageBlock(
                    bounding_box=planned.bounding_box,
                    description=planned.description or "Embedded image",
                )
            )

        return blocks

    @staticmethod
    def _describe_image(img) -> Optional[str]:
        """
        Extract image bytes from an openpyxl Image and send to the
        vision model.  Raw image bytes (PNG/JPEG/GIF) are supported
        by all providers.
        """
        try:
            img_data = None
            if callable(getattr(img, "_data", None)):
                img_data = img._data()
            elif hasattr(img, "ref") and hasattr(img.ref, "read"):
                img_data = img.ref.read()

            if not img_data:
                return None

            # Determine MIME type from image header bytes
            mime = "image/png"
            if img_data[:3] == b"\xff\xd8\xff":
                mime = "image/jpeg"
            elif img_data[:4] == b"\x89PNG":
                mime = "image/png"
            elif img_data[:4] == b"GIF8":
                mime = "image/gif"

            prompt = get_image_description_prompt()
            ai = get_decision_for_media_service()
            return ai.get_decision_for_media(prompt, img_data, mime_type=mime)
        except Exception:
            logger.warning(
                "  [ImageExtractor] Failed to describe image", exc_info=True
            )
            return None
