"""
Chart Extractor — wraps the existing ChartExtractor and optionally
sends the chart screenshot region to a vision model for a
natural-language description.
"""

from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from ai.factory import get_decision_for_media_service
from dto.blocks import Block, ChartBlock
from dto.cell_data import CellData
from extractors.chart import ChartExtractor as OriginalChartExtractor

from agentic_flow.dto.plan import PlannedBlock
from agentic_flow.extractors.base import BaseExtractor
from agentic_flow.prompts.chart import get_chart_description_prompt

logger = logging.getLogger(__name__)


class AgenticChartExtractor(BaseExtractor):
    """
    Extracts chart data using the existing ChartExtractor and
    enhances with a vision-model description when a screenshot is
    available.
    """

    def __init__(self) -> None:
        self._original = OriginalChartExtractor()

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
        # Use the existing chart extractor to get structured data
        chart_datas = self._original.extract(ws, wb)

        blocks: List[Block] = []
        for cd in chart_datas:
            description: Optional[str] = None

            # Try to get a vision description using the full screenshot
            if screenshot_bytes:
                description = self._describe_chart(
                    screenshot_bytes, cd.title, cd.chart_type,
                    [s.name for s in cd.series if s.name],
                )

            blocks.append(
                ChartBlock(
                    bounding_box=cd.bounding_box,
                    chart_data=cd,
                    description=description,
                )
            )

        # If the existing extractor found no charts but the planner
        # thinks there's one, create a stub block
        if not blocks:
            blocks.append(
                ChartBlock(
                    bounding_box=planned.bounding_box,
                    description=planned.description or "Chart (no data extracted)",
                )
            )

        return blocks

    @staticmethod
    def _describe_chart(
        screenshot_bytes: bytes,
        title: Optional[str],
        chart_type: Optional[str],
        series_names: List[str],
    ) -> Optional[str]:
        """
        Send the screenshot to a vision model with the chart description
        prompt.
        """
        try:
            prompt = get_chart_description_prompt(title, chart_type, series_names)
            ai = get_decision_for_media_service()
            return ai.get_decision_for_media(
                prompt, screenshot_bytes, mime_type="image/png"
            )
        except Exception:
            logger.warning(
                "  [ChartExtractor] Vision description failed", exc_info=True
            )
            return None
